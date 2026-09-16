using Microsoft.AspNetCore.Mvc;
using OutlookAddinProject.Data;

namespace OutlookAddinProject.Controllers;

public sealed record AttachmentDto(string Name, string ContentType, long SizeBytes, string ContentBase64);

public sealed record EmailDto(
    string Subject,
    string From,
    string To,
    string? Cc,
    string Body,
    string? UserName,
    int? UserId,
    List<AttachmentDto>? Attachments);

[ApiController]
[Route("api/Emails")]
public class EmailController : ControllerBase
{
    private readonly AppDbContext _db;

    public EmailController(AppDbContext db)
    {
        _db = db;
    }

    // POST /api/Emails/CreateDocumnet
    // NOTE: keeping this exact route name ("Documnet") to match the task pane's
    // existing fetch call. Rename both sides together if you want to fix the typo.
    [HttpPost("CreateDocumnet")]
    public async Task<ActionResult<string>> CreateDocument([FromBody] EmailDto dto)
    {
        var record = await SaveEmailAsync(dto);
        return Ok($"Document created. Id={record.Id}");
    }

    [HttpPost("CreateCorrespondence")]
    public async Task<ActionResult<string>> CreateCorrespondence([FromBody] EmailDto dto)
    {
        var record = await SaveEmailAsync(dto);
        return Ok($"Correspondence created. Id={record.Id}");
    }

    private async Task<EmailRecord> SaveEmailAsync(EmailDto dto)
    {
        var record = new EmailRecord
        {
            Subject = dto.Subject,
            FromAddress = dto.From,
            ToAddress = dto.To,
            CcAddress = dto.Cc,
            Body = dto.Body,
            UserName = dto.UserName,
            UserId = dto.UserId,
            CreatedAt = DateTime.UtcNow,
            Attachments = (dto.Attachments ?? new List<AttachmentDto>())
                .Select(a => new AttachmentRecord
                {
                    FileName = a.Name,
                    ContentType = string.IsNullOrEmpty(a.ContentType) ? "application/octet-stream" : a.ContentType,
                    SizeBytes = a.SizeBytes,
                    Content = string.IsNullOrEmpty(a.ContentBase64)
                        ? Array.Empty<byte>()
                        : Convert.FromBase64String(a.ContentBase64)
                })
                .ToList()
        };

        _db.Emails.Add(record);
        await _db.SaveChangesAsync();
        return record;
    }
}
