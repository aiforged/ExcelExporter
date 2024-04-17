using ExcelExporter.Models;

using OfficeOpenXml;

namespace ExcelExporter
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var ENV_PREFIX = "EXLGEN_";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureHostConfiguration(config =>
                {
                    config.AddEnvironmentVariables(prefix: ENV_PREFIX);
                    config.AddUserSecrets<Program>();
                })
                .ConfigureServices((hostContext, services) =>
                {
                    IConfiguration configuration = hostContext.Configuration;
                    string configValue = configuration.GetSection($"{ENV_PREFIX}CONFIG").Value;

                    configValue = configValue.Replace("'", "\"");
                    Config config = System.Text.Json.JsonSerializer.Deserialize<Config>(configValue, new System.Text.Json.JsonSerializerOptions()
                    {
                        AllowTrailingCommas = true,
                        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull | System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingDefault,
                    });

#if DEBUG
                    config.APIKey = configuration["AIForged:ApiKey"];
                    config.ProjectId = Convert.ToInt32(configuration["AIForged:ProjectId"]);
                    config.ServiceId = Convert.ToInt32(configuration["AIForged:ServiceId"]);
                    config.AIForgedEndpoint = configuration["AIForged:EndPoint"];
                    config.MasterParamDefName = configuration["AIForged:MasterParamDefName"];
                    config.InputTemplatePath = configuration["AIForged:InputTemplatePath"];
#endif

                    services.AddSingleton(config);
                    services.AddHostedService<Worker>();
                    services.AddLogging(config => config
                .SetMinimumLevel(LogLevel.Trace)
#if DEBUG
                .AddDebug()
                .AddConsole()
#endif
                .AddOpenTelemetry()
                );
                })
                .Build();

            host.Run();
        }
    }
}