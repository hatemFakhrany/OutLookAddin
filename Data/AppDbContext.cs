using Microsoft.EntityFrameworkCore;

namespace OutlookAddinProject.Data;

public class EmailRecord
{
    public int Id { get; set; }
    public string Subject { get; set; } = string.Empty;
    public string FromAddress { get; set; } = string.Empty;
    public string ToAddress { get; set; } = string.Empty;
    public string? CcAddress { get; set; }
    public string Body { get; set; } = string.Empty;
    public string? UserName { get; set; }
    public int? UserId { get; set; }
    public DateTime CreatedAt { get; set; }
    public List<AttachmentRecord> Attachments { get; set; } = new();
}

public class AttachmentRecord
{
    public int Id { get; set; }
    public int EmailId { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public byte[] Content { get; set; } = Array.Empty<byte>();
    public EmailRecord? Email { get; set; }
}

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    public DbSet<EmailRecord> Emails => Set<EmailRecord>();
    public DbSet<AttachmentRecord> Attachments => Set<AttachmentRecord>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<EmailRecord>()
            .HasMany(e => e.Attachments)
            .WithOne(a => a.Email)
            .HasForeignKey(a => a.EmailId);

        modelBuilder.Entity<EmailRecord>()
            .Property(e => e.Body)
            .HasColumnType("nvarchar(max)");

        modelBuilder.Entity<AttachmentRecord>()
            .Property(a => a.Content)
            .HasColumnType("varbinary(max)");
    }
}
