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

                    //Parse our app Config from appsettings.json
                    Config config = configuration.GetSection("Config").Get<Config>();

#if DEBUG
                    //Override our config from user secrets when debugging
                    config.APIKey = configuration["AIForged:ApiKey"];
                    config.ProjectId = Convert.ToInt32(configuration["AIForged:ProjectId"]);
                    config.ServiceId = Convert.ToInt32(configuration["AIForged:ServiceId"]);
                    config.AIForgedEndpoint = configuration["AIForged:EndPoint"];
                    config.MasterParamDefName = configuration["AIForged:MasterParamDefName"];
                    config.InputTemplatePath = configuration["AIForged:InputTemplatePath"];
                    config.InputDocumentStatus = Enum.Parse<AIForged.API.DocumentStatus>(configuration["AIForged:InputDocumentStatus"]);
                    config.ProcessedDocumentStatus = Enum.Parse<AIForged.API.DocumentStatus>(configuration["AIForged:ProcessedDocumentStatus"]);
                    config.EmailClientId = configuration["AIForged:EmailClientId"];
                    config.EmailTenantId = configuration["AIForged:EmailTenantId"];
                    config.EmailClientSecret = configuration["AIForged:EmailClientSecret"];
                    config.EmailFromAddress = configuration["AIForged:EmailFromAddress"];
                    config.EmailRecipients = System.Text.Json.JsonSerializer.Deserialize<List<string>>(configuration["AIForged:EmailRecipients"]);
#endif
                    //Add config as app lifetime service
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