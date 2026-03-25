using System;
using System.IO;
using AFH.Common.SharePointUtils.Extensions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace AFH.Common.SharePointUtils.Tests.Fixtures;

public sealed class SharepointIntegrationFixture : IDisposable
{
    private readonly ServiceProvider _sp;

    public SharepointIntegrationFixture()
    {
        var basePath = AppContext.BaseDirectory;

        // Load .env
        var envPath = Path.Combine(basePath, ".env");
        if (File.Exists(envPath))
        {
            DotNetEnv.Env.Load(envPath);
            Console.WriteLine($"[Fixture] .env loaded from: {envPath}");
        }

        // Build config
        var cfg = new ConfigurationBuilder()
            .SetBasePath(basePath)
            .AddJsonFile("local.settings.json", optional: true)
            .AddEnvironmentVariables()
            .Build();
        DumpConfiguration(cfg);
        // Validate required keys
        Ensure(cfg, "AzureAD:TenantId");
        Ensure(cfg, "AzureAD:ClientId");
        Ensure(cfg, "AzureAD:ClientSecret");
        Ensure(cfg, "AzureAD:AuthorityHost");
        Ensure(cfg, "AzureAD:Scopes:0");

        Ensure(cfg, "SharePointListsConfig:ClientDocuments:SiteId");
        Ensure(cfg, "SharePointListsConfig:ClientDocuments:ListId");

        var services = new ServiceCollection();

        services.AddLogging(b =>
            b.AddConsole().SetMinimumLevel(LogLevel.Information));
        services.AddSingleton<IConfiguration>(cfg);
        services.AddSharePoint(cfg);

        _sp = services.BuildServiceProvider();
    }

    public T Get<T>() where T : notnull => _sp.GetRequiredService<T>();
    public void Dispose() => _sp.Dispose();
    private static void DumpConfiguration(IConfiguration cfg)
    {
        Console.WriteLine("=== BEGIN CONFIGURATION DUMP ===");
        foreach (var kv in cfg.AsEnumerable())
        {
            // Only show the interesting bits
            if (kv.Key.StartsWith("AzureAD", StringComparison.OrdinalIgnoreCase) ||
                kv.Key.StartsWith("SharePointListsConfig", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"{kv.Key} = {kv.Value}");
            }
        }
        Console.WriteLine("=== END CONFIGURATION DUMP ===");
    }
    private static void Ensure(IConfiguration cfg, string key)
    {
        if (string.IsNullOrWhiteSpace(cfg[key]))
            throw new InvalidOperationException($"Missing required config: {key}");
    }
}