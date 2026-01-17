using Microsoft.Extensions.Configuration;
using System.IO;
using System.Reflection;

namespace Audit_B;

public static class ConfigurationHelper
{
    private static IConfiguration? _configuration;

    public static IConfiguration GetConfiguration()
    {
        if (_configuration == null)
        {
            // Read appsettings.json from embedded resource
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "Audit_B.appsettings.json";
            
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException($"Embedded resource '{resourceName}' not found");

                using (var reader = new StreamReader(stream))
                {
                    var builder = new ConfigurationBuilder()
                        .AddJsonStream(stream);
                    
                    _configuration = builder.Build();
                }
            }
        }

        return _configuration;
    }

    public static string GetConnectionString(string name = "DefaultConnection")
    {
        var config = GetConfiguration();
        var connectionString = config.GetConnectionString(name);
        
        if (string.IsNullOrEmpty(connectionString))
        {
            throw new InvalidOperationException($"Connection string '{name}' not found in appsettings.json");
        }

        return connectionString;
    }
}