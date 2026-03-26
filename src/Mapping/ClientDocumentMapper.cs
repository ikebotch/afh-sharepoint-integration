using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Mapping;

public static class ClientDocumentMapper
{
    public static ClientDocument ToDomain(SharePointListItemModel item)
    {
        var fields = item.Fields;

        return new ClientDocument
        {
            Id = item.Id ?? string.Empty,
            ClientId = fields.TryGetValue("ClientId", out var clientId) ? clientId?.ToString() ?? string.Empty : string.Empty,
            DocumentType = fields.TryGetValue("DocumentType", out var docType) ? docType?.ToString() ?? string.Empty : string.Empty,
            FileName = fields.TryGetValue("FileName", out var fileName) ? fileName?.ToString() ?? string.Empty : string.Empty,
            Status = fields.TryGetValue("Status", out var status) && Enum.TryParse<ClientDocumentStatus>(status?.ToString(), out var parsedStatus)
                ? parsedStatus
                : ClientDocumentStatus.Pending,
            UploadedAt = item.CreatedDateTime ?? DateTimeOffset.MinValue,
            UploadedBy = fields.TryGetValue("UploadedBy", out var uploadedBy) ? uploadedBy?.ToString() : null,
            ModifiedAt = item.LastModifiedDateTime,
            ModifiedBy = fields.TryGetValue("ModifiedBy", out var modifiedBy) ? modifiedBy?.ToString() : null,
            DriveItemId = fields.TryGetValue("DriveItemId", out var driveId) ? driveId?.ToString() : null,
            FileSize = fields.TryGetValue("FileSize", out var size) && long.TryParse(size?.ToString(), out var parsedSize) ? parsedSize : null,
            FileUrl = fields.TryGetValue("FileUrl", out var url) ? url?.ToString() : null,
            Description = fields.TryGetValue("Description", out var desc) ? desc?.ToString() : null,
            RawFields = new Dictionary<string, object?>(fields, StringComparer.OrdinalIgnoreCase)
        };
    }

    public static ClientDocument ToDomain(ListItem item)
    {
        return ToDomain(SharePointMapper.ToModel(item));
    }

    public static ListItem ToListItem(ClientDocument doc)
        => SharePointMapper.ToListItem(ToFields(doc));

    public static IReadOnlyDictionary<string, object?> ToFields(ClientDocument doc)
    {
        return new Dictionary<string, object?>
        {
            { "ClientId", doc.ClientId },
            { "DocumentType", doc.DocumentType },
            { "FileName", doc.FileName },
            { "Status", doc.Status.ToString() },
            { "UploadedBy", doc.UploadedBy },
            { "ModifiedBy", doc.ModifiedBy },
            { "DriveItemId", doc.DriveItemId },
            { "FileSize", doc.FileSize },
            { "FileUrl", doc.FileUrl },
            { "Description", doc.Description }
        };
    }
}
