# AFH Common SharePoint SDK

Low-level SharePoint and Microsoft Graph SDK for AFH services.

## Capabilities

- Graph authentication and DI registration
- site and list discovery helpers
- generic list CRUD operations
- filter/select/expand/order query support
- field discovery for list columns
- display-name to internal-name resolution
- mapping profile support for consumer-defined property mappings
- typed field-access helpers for common SharePoint value shapes
- generic list-item mapping into SDK models
- document-library metadata, upload, download, list, rename, and delete operations
- compatibility wrappers for existing `ISharePointListService` and `IClientDocumentService` consumers

## Boundary

This package is provider-focused and business-agnostic. It owns SharePoint mechanics, request shaping, mapping, and reusable models. It does not own adviser ranking, booking orchestration, SQL persistence, or service business rules.

## Configuration

Bind either:

- `AzureAD`
- `SharePointGraph`

and optionally:

- `SharePointListsConfig`
- `ClientDocuments`

## DI

```csharp
services.AddSharePoint(configuration);
```

Registered SDK surface:

- `ISharePointDiscoveryClient`
- `ISharePointListClient`
- `ISharePointDocumentClient`
- `ISharePointFieldResolver`

Compatibility registrations:

- `ISharePointListService`
- `IClientDocumentService`

## Testing

- Unit-style SDK tests run locally without tenant access.
- Tenant-backed integration tests are automatically skipped unless the required SharePoint environment variables are present.

## Field Discovery

```csharp
var fields = await fieldResolver.GetFieldsAsync(siteId, listId, ct);

foreach (var field in fields)
{
    Console.WriteLine($"{field.DisplayName} -> {field.InternalName} ({field.Type})");
}
```

## Mapping Profiles

```csharp
var profile = new SharePointMappingProfile()
    .MapDisplay("ClientName", "Client Name")
    .MapInternal("PolicyNumber", "PolicyNumber");

var resolved = await fieldResolver.ResolveProfileAsync(siteId, listId, profile, ct);
var clientName = item.Fields.GetString(resolved, "ClientName");
```
