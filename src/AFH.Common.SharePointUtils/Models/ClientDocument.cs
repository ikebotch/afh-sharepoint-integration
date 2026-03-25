namespace AFH.Common.SharePointUtils.Models;


public class ClientDocument
{
    /// <summary>SharePoint Item ID</summary>
    public required string Id { get; init; }

    /// <summary>Business client identifier</summary>
    public required string ClientId { get; init; }

    /// <summary>Type of document (e.g. Contract, Invoice, Report)</summary>
    public required string DocumentType { get; init; }

    /// <summary>Original file name</summary>
    public required string FileName { get; init; }

    /// <summary>Status of the document</summary>
    public ClientDocumentStatus Status { get; init; }

    /// <summary>When the document was uploaded</summary>
    public DateTimeOffset UploadedAt { get; init; }

    /// <summary>Who uploaded the document</summary>
    public string? UploadedBy { get; init; }

    /// <summary>Last modified timestamp</summary>
    public DateTimeOffset? ModifiedAt { get; init; }

    /// <summary>Who last modified the document</summary>
    public string? ModifiedBy { get; init; }

    /// <summary>SharePoint drive item ID if stored in a document library</summary>
    public string? DriveItemId { get; init; }

    /// <summary>File size in bytes</summary>
    public long? FileSize { get; init; }

    /// <summary>Direct URL to the file in SharePoint</summary>
    public string? FileUrl { get; init; }

    /// <summary>Optional description or notes</summary>
    public string? Description { get; init; }

    /// <summary>Raw SharePoint field values for extensibility</summary>
    public IDictionary<string, object?> RawFields { get; init; } = new Dictionary<string, object?>();
}
