using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.Routing;
using WebPortal.ActionFilter;
using WebPortal.Integration.Api;
using WebPortal.Integration.Api.Administration.ConnectedUsers;
using WebPortal.Models.Account;
using WebPortal.Models.Administration;
using WebPortal.TaskflowSdkCoreIntegration;
using WebPortal.Utilities;

namespace OutlookAddinProject.Controllers;

[Internationalization]
    public class BaseControllerAnonymous : ControllerBase
{
        /// <summary>Typed Administration API client used by migrated actions. Resolving it through
        /// the existing MVC dependency resolver keeps legacy controller constructors compatible
        /// while HttpClientFactory still owns the underlying client lifecycle. Exposed here (rather
        /// than on <see cref="BaseController"/>) because AccountController's login flows need it
        /// before a session/[Authorize] context exists.</summary>
        protected IConnectedUsersProxy ConnectedUsersApi =>
            DependencyResolver.Current.GetService<IConnectedUsersProxy>();

        protected IWebPortalApi WebPortalApi =>
            DependencyResolver.Current.GetService<IWebPortalApi>();
    }

    [Authorize]
    [ValidateSessionFilter]
    public class BaseController : BaseControllerAnonymous
    {
        protected CoreProcessingClass _coreProcessing = DependencyResolver.Current.GetService<CoreProcessingClass>();

        public AppUserState AppUserState { get; private set; }

        protected override void Initialize(RequestContext requestContext)
        {
            base.Initialize(requestContext);

            // Grab the user's login information from Identity
            AppUserState appUserState = new AppUserState();
            if (User is ClaimsPrincipal)
            {
                var user = User as ClaimsPrincipal;
                var claims = user.Claims.ToList();

                var name = getClaim(claims, ClaimTypes.Name);
                if (!name.IsNullOrEmpty())
                {
                    var userStateString = getClaim(claims, "userState");
                    //var id = GetClaim(claims, ClaimTypes.NameIdentifier);
                    if (!string.IsNullOrEmpty(userStateString))
                    {
                        appUserState.FromString(userStateString);

                        AppUserState = appUserState;
                        CurrentUser.SetUserData(AppUserState);
                        ViewData["UserState"] = AppUserState;
                    }
                    else
                    {
                        CurrentUser.ClearUserDataAndServices();
                    }
                }
                else
                {
                    //CurrentUser.ClearUserDataAndServices();
                    AppUserState userTest = new AppUserState();
                    userTest.UserId = "1";
                    userTest.Username = "Ecmadmin";
                    userTest.EnFullName = "System Administrator";
                    userTest.FullName = "System Administrator";
                    userTest.UserSessionId = "e51131b7121526";
                    CurrentUser.SetUserData(userTest);
                    ViewData["UserState"] = userTest;
                }
            }


        }

        private string getClaim(List<Claim> claims, string claimType)
        {
            var claim = claims.FirstOrDefault(x => x.Type == claimType);
            return claim == null ? null : claim.Value;
        }

        /// <summary>Document/File per-user permission-action flags, entirely via WebPortal.Api/EF
        /// Core (SecurableObjectsController.GetObjectUserPermissions -> EffectivePermission). Super
        /// users are resolved server-side (UserPermissionCache.IsSuperUserAsync) to full control,
        /// so no client-side super-user short-circuit is needed here.</summary>
        protected Task<WebPortal.Models.Administration.UserPermissionsModel> GetNewAclConceptPermissionsAsync(int entityType, long entityId) =>
            WebPortalApi.AclProxy.GetObjectUserPermissionsAsync(entityType, entityId, CurrentUser.UserData.UserId);
    }


