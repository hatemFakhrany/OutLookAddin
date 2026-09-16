using Hangfire;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Mvc;
using WebPortal.Integration.Api.Documents;
using WebPortal.Integration.Api.Folders;
using WebPortal.Models.Archive;
using WebPortal.MyConstants;
using WebPortal.Utilities;

namespace OutlookAddinProject.Controllers;

public partial class ArchiveController : BaseController
    {
        public async Task<ActionResult> CreateNewDocument(string folderid, int docTypeId = 0, string docIntegrationId = null)
        {
            var model = new DocMetaDataModel();
            model.DocObj = new ArchiveDocumentViewModel
            {
                ReceiveDate = DateTime.Now
            };
            int? decFolderid = DecryptFolderIdOrNull(folderid);

            try
            {
                ApplyDocumentMetadataLookups(model, await GetArchiveLookupsAsync(decFolderid));
            }
            catch (Exception exception)
            {
                return new HttpStatusCodeResult(503, exception.Message);
            }

            if (model.SubOrgsList == null || model.SubOrgsList.Count < 2)
            {
                model.DocObj.SubOrgID = 1;
            }
            if (decFolderid.HasValue)
            {
                model.NewRealFolderID = folderid;
                try
                {
                    model.UserPermissions = await GetNewAclConceptFolderPermissionsAsync(decFolderid.Value);
                }
                catch (Exception exception)
                {
                    return new HttpStatusCodeResult(503, exception.Message);
                }
            }
            var defaultClassificationEntity = model.DocClassificationsList.FirstOrDefault(x => x.Selected);
            if (defaultClassificationEntity != null)
            {
                model.DocObj.DOCCLASSIFICATIONID = Convert.ToInt32(defaultClassificationEntity.Value);
            }

            // Bound directly off DocTypesList (DOCTYPE-table-driven via GetDCsFolderWithSecurity)
            // instead of a fresh CORRESPONDANCE_FOLDER_DC_REF_GetBy_DocTypeId DAL call -- DocTypeId
            // is already known (the action's own parameter); only DCID (still needed for the
            // <select>'s bound value and the legacy, non-NewAclConcept create path) requires a
            // lookup, and DocTypesList already carries it for every row.
            model.DocTypeId = docTypeId;
            // Doctypeid drives which <option> the doctype <select> pre-selects
            // (_docMetaDataFieldsPartial.cshtml matches on Doctypeid, not DCID -- see that file's
            // comment); DCID is still resolved below only because the <select>'s posted value and
            // the legacy create path need it.
            model.DocObj.Doctypeid = docTypeId;
            var matchedDocType = model.DocTypesList.FirstOrDefault(dt => dt.DocTypeID == docTypeId.ToString());
            if (matchedDocType != null && decimal.TryParse(matchedDocType.DCID, out var matchedDcId))
            {
                model.DocObj.DCID = matchedDcId;
            }
            model.DocIntegrationId = docIntegrationId;
            if (!docIntegrationId.IsNullOrEmpty())
            {
                model.DocObj.DOCNAME = docIntegrationId;
            }

            return View(model);
        }

        /// <summary>Resolves a client-supplied folder id that may be either a real encrypted id or
        /// one of the small plain sentinel values the folder tree also uses for virtual nodes
        /// (e.g. -3 = Favourites, -1 = "no folder", 0/1 = root) -- same convention
        /// SearchDocuments.cs's getViewSearchModel already applies to RealFolderId. Calling
        /// ProcessingClass.DecryptId directly on one of these sentinels throws (they are not
        /// encrypted data), which is what turned a Favourites-scoped "Add New Document" click into
        /// an uncaught 500. Any sentinel/non-positive result is treated as "no folder selected"
        /// rather than passed on as a real folder id.</summary>
        private static int? DecryptFolderIdOrNull(string folderid)
        {
            if (folderid.IsNullOrEmpty())
                return null;

            var isPlainSentinel = Regex.IsMatch(folderid, @"^-?\d+$") && int.Parse(folderid) <= 1;
            var resolvedFolderId = isPlainSentinel ? int.Parse(folderid) : IdProtection.Decrypt(folderid);
            return resolvedFolderId > 0 ? resolvedFolderId : (int?)null;
        }

        private async Task<WebPortal.Models.Administration.UserPermissionsModel> GetNewAclConceptFolderPermissionsAsync(int folderId)
        {
            if (folderId <= 0)
                throw new ArgumentOutOfRangeException(nameof(folderId));

            var result = await _webPortalApi.AclProxy.GetObjectUserPermissionsAsync(
                PortalCorrespondenceConstants.EntityTypeFolder,
                folderId,
                CurrentUser.UserData.UserId);
            if (result == null)
                throw new InvalidOperationException("WebPortal.Api folder permissions are unavailable.");
            return result;
        }

        /// <summary>Creates a document through WebPortal.Api.</summary>
        [HttpPost]
        public async System.Threading.Tasks.Task<string> CreateNewDocument(DocMetaDataModel model)
        {
            return await CreateNewDocumentViaNewAclConceptAsync(model);
        }

        /// <summary>New-technique document-creation flow: EVERY database operation --
        /// ACL/EntityPermission resolution, the core row, FolderItem placement, audit logging,
        /// external-integration save -- happens server-side in WebPortal.Api's
        /// CreateCorrespondenceDocumentCommand (see its doc comment for the exact mapping), not
        /// via any legacy persistence call in this method. Workflow eligibility is
        /// evaluated by WebPortal.Api after the document transaction commits.
        /// <para>Deliberately does NOT fall back to m_ProcObj.AddCORRESPONDANCE_DOCUMENT -- not on
        /// a transport/API exception, and not on a null result (API base URL unconfigured) either.
        /// DOCUMENT_METADATA is written ONLY by CreateCorrespondenceDocumentCommandHandler, in the
        /// SAME transaction as the CORRESPONDANCE_DOCUMENT row; the legacy proc has no knowledge of
        /// that table at all. A silent legacy fallback here used to leave a document that "saved
        /// successfully" but could never appear in ListFolderDocumentsQuery/the document explorer,
        /// since that read path is DOCUMENT_METADATA-only. Any
        /// inability to write via the API is reported as a failure instead.</para></summary>
        private async System.Threading.Tasks.Task<string> CreateNewDocumentViaNewAclConceptAsync(DocMetaDataModel model)
        {
            int? decFolderid;
            if (string.IsNullOrEmpty(model.NewRealFolderID))
                decFolderid = null;
            else
                decFolderid = IdProtection.Decrypt(model.NewRealFolderID);
            if (model.SubOrgsList == null || model.SubOrgsList.Count < 2)
            {
                model.DocObj.SubOrgID = 1;
            }
            if (!decFolderid.HasValue)
            {
                decFolderid = -1;
            }

            // model.DocTypeId is posted via a hidden field kept in sync with #selectDocType
            // client-side (docTypeIdSync in _docMetaDataFieldsPartial.cshtml); when that sync
            // doesn't fire before submit, model.DocTypeId can still bind to its default (0) even
            // though model.DocObj.DCID (the <select>'s own "DocObj.DCID"-bound value) correctly
            // reflects what the user picked. Re-resolve DocTypeId from DCID instead of trusting the
            // hidden field. model.DocTypesList is never posted (only real inputs are), so unlike
            // the GET action it is null here on POST, so resolve it through the same API lookup
            // slice used by the create form.
            try
            {
                ApplyDocumentMetadataLookups(model, await GetArchiveLookupsAsync());
            }
            catch (Exception exception)
            {
                return new PortalOperationResult<string>
                {
                    Status = UIConstants.Status_Failed,
                    Result = "-1",
                    Message = exception.Message
                }.SerializeToJson();
            }
            var selectedDocType = model.DocTypesList?
                .FirstOrDefault(dt => decimal.TryParse(dt.DCID, out var dcid) && dcid == model.DocObj.DCID);
            if (selectedDocType != null && int.TryParse(selectedDocType.DocTypeID, out var resolvedDocTypeIdFromDcid))
            {
                model.DocTypeId = resolvedDocTypeIdFromDcid;
            }

            PortalOperationResult<string> addResult = new PortalOperationResult<string>()
            {
                Status = UIConstants.Status_Ok,
                Result = "-1",
                Message = "Saved Successfully".ToLocalizedValue()
            };

            if (model.DocObj.EncryptDocOwnerUserId != "-1")
            {
                var decDocOwnerUserId = IdProtection.Decrypt(model.DocObj.EncryptDocOwnerUserId);
                model.DocObj.DOCOWNERUSERID = decDocOwnerUserId;
            }
            var documentFolderId = decFolderid > 0 ? (decimal?)decFolderid.Value : null;
            try
            {
                // model.ACLID is the RAW client-submitted value
                // (setUpdateFieldsOnly copied it from model.ACLID) -- sent to WebPortal.Api as the
                // REQUESTED value; the server decides whether to honor or override it via the same
                // Folder_Doc_DocFile_ApplyACL check legacy runs. Neither a null result (API base
                // URL unconfigured) nor a transport/API exception falls back to legacy -- see this
                // method's doc comment for why; both are reported as a failure below instead.
                try
                {
                    var newAclConceptResult = await _webPortalApi.ArchiveProxy.CreateAsync(new CreateCorrespondenceDocumentInput
                    {
                        DocName = model.DocObj.DOCNAME,
                        DocTypeId = model.DocTypeId,
                        FolderId = documentFolderId,
                        OwnerUserName = CurrentUser.UserData.Username,
                        ReceiveDate = model.DocObj.ReceiveDate,
                        IsDocRestored = model.DocObj.IsDocRestored,
                        FileNumberCode = model.DocObj.FileNumberCode,
                        Refrence = model.DocObj.Refrence,
                        BarCodeFormatId = model.DocObj.BarCodeFormatID,
                        RequestedAclId = model.ACLID,
                        DocOwnerUserId = model.DocObj.DOCOWNERUSERID,
                        DocClassificationId = model.DocObj.DOCCLASSIFICATIONID.HasValue ? (long?)model.DocObj.DOCCLASSIFICATIONID.Value : null,
                        SubOrgId = model.DocObj.SubOrgID ?? 1,
                        CreatingUserId = CurrentUser.UserData.UserId,
                        DocIntegrationId = model.DocIntegrationId,
                        Values = ToFieldValues(model.CustomAttributesList),
                    });

                    if (newAclConceptResult != null)
                    {
                        model.DocObj.DOCID = newAclConceptResult.DocId;
                        model.DocObj.DCID = newAclConceptResult.DocTypeId;
                    }
                    else
                    {
                        return new PortalOperationResult<string>
                        {
                            Status = UIConstants.Status_Failed,
                            Result = "-1",
                            Message = "Error during save".ToLocalizedValue()
                        }.SerializeToJson();
                    }
                }
                catch (Exception)
                {
                    return new PortalOperationResult<string>
                    {
                        Status = UIConstants.Status_Failed,
                        Result = "-1",
                        Message = "Error during save".ToLocalizedValue()
                    }.SerializeToJson();
                }

                try
                {
                    await EvaluateDocumentWorkflowViaApiAsync(
                        decFolderid > 0 ? decFolderid : null,
                        Convert.ToInt64(model.DocObj.DOCID),
                        PortalCorrespondenceConstants.WorkflowTriggerAddDocument,
                        PortalCorrespondenceConstants.DocumentActionAddNewDocument);
                }
                catch { }

                addResult = new PortalOperationResult<string>()
                {
                    Status = UIConstants.Status_Ok,
                    Result = IdProtection.Encrypt(model.DocObj.DOCID).ToString(),
                    Message = "Saved Successfully".ToLocalizedValue()
                };
            }
            catch (Exception exp)
            {
                addResult = new PortalOperationResult<string>()
                {
                    Status = UIConstants.Status_Failed,
                    Result = "-1",
                    Message = string.Format("{0} @ {1} {2}", "Error during save".ToLocalizedValue(),
                    DateTime.Now.ToString("dd/MM/yyyy HH:mm"), exp.Message)
                };
            }


            return addResult.SerializeToJson();
        }

        #region Create New document by AI
        public async Task<ActionResult> CreateNewDocumentByAI(string folderid, int docTypeId = 0, string docIntegrationId = null)
        {
            var model = new DocMetaDataModel();
            model.DocObj = new ArchiveDocumentViewModel
            {
                ReceiveDate = DateTime.Now
            };
            int? decFolderid = DecryptFolderIdOrNull(folderid);

            try
            {
                ApplyDocumentMetadataLookups(model, await GetArchiveLookupsAsync(decFolderid));
            }
            catch (Exception exception)
            {
                return new HttpStatusCodeResult(503, exception.Message);
            }

            if (model.SubOrgsList == null || model.SubOrgsList.Count < 2)
            {
                model.DocObj.SubOrgID = 1;
            }
            if (decFolderid.HasValue)
            {
                model.NewRealFolderID = folderid;
                try
                {
                    model.UserPermissions = await GetNewAclConceptFolderPermissionsAsync(decFolderid.Value);
                }
                catch (Exception exception)
                {
                    return new HttpStatusCodeResult(503, exception.Message);
                }
            }
            var defaultClassificationEntity = model.DocClassificationsList.FirstOrDefault(x => x.Selected);
            if (defaultClassificationEntity != null)
            {
                model.DocObj.DOCCLASSIFICATIONID = Convert.ToInt32(defaultClassificationEntity.Value);
            }

            // See CreateNewDocument's matching comment above -- same DocTypesList-driven binding.
            model.DocTypeId = docTypeId;
            // Doctypeid drives which <option> the doctype <select> pre-selects
            // (_docMetaDataFieldsPartial.cshtml matches on Doctypeid, not DCID -- see that file's
            // comment); DCID is still resolved below only because the <select>'s posted value and
            // the legacy create path need it.
            model.DocObj.Doctypeid = docTypeId;
            var matchedDocType = model.DocTypesList.FirstOrDefault(dt => dt.DocTypeID == docTypeId.ToString());
            if (matchedDocType != null && decimal.TryParse(matchedDocType.DCID, out var matchedDcId))
            {
                model.DocObj.DCID = matchedDcId;
            }
            model.DocIntegrationId = docIntegrationId;
            if (!docIntegrationId.IsNullOrEmpty())
            {
                model.DocObj.DOCNAME = docIntegrationId;
            }

            return View(model);
        }


        #endregion
    }

