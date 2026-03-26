using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Tests.Helpers;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;
using Xunit.Abstractions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
[Trait("Op", "Verify")]
public sealed class ClientDocumentConfigIntegrationTests
{
    private readonly IClientDocumentService _svc;
    private readonly ITestOutputHelper _output;

    public ClientDocumentConfigIntegrationTests(
        SharepointIntegrationFixture fx,
        ITestOutputHelper output)
    {
        _svc = fx.Get<IClientDocumentService>();
        _output = output;
    }



    [RequiresSharePointIntegrationFact]
    public async Task Verify_ClientDocumentListsExist_returns_true()
    {
        // Act
        var result = await _svc.VerifyConfigurationAsync();

        // Assert
        result.Should().BeTrue("both ClientDocuments and Config lists should exist in the test environment");

        _output.WriteLine("ClientDocumentVerifyListConfig returned: " + result);
    }
}