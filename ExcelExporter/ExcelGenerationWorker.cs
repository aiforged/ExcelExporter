using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

using ExcelExporter.Clients;
using ExcelExporter.Generators;
using ExcelExporter.Models;

using LarcAI;

namespace ExcelExporter
{
    public class ExcelGenerationWorker
    {
        ILogger logger;
        IConfig config;
        CancellationToken stoppingToken;

        public ExcelGenerationWorker(ILogger logger, IConfig config, CancellationToken stoppingToken)
        {
            this.logger = logger;
            this.config = config;
            this.stoppingToken = stoppingToken;
        }

        public async Task CheckAndGenerateExcelAsync()
        {
            try
            {
                ExcelGenerator excelGenerator = new ExcelGenerator(logger);
                AIForgedClient aiForgedClient = new AIForgedClient(logger, config);

                var docs = await aiForgedClient.GetProcessedDocsAsync(config.ProjectId, config.ServiceId, config.InputDocumentStatus);

                logger.LogInformation($"{DateTime.Now} CheckAndGenerateExcelAsync: Found {docs.Count} documents to process...");
                foreach (var doc in docs)
                {
                    try
                    {
                        logger.LogInformation($"{DateTime.Now} CheckAndGenerateExcelAsync: Start processing document {doc}...");

                        var docResult = await aiForgedClient.GetDocumentResultsAsync(config.ProjectId, config.ServiceId, doc.Id);
                        var docDataResult = await aiForgedClient.GetDocumentDataAsync(doc.MasterId ?? doc.Id);

                        doc.Data = docDataResult?.ToObservableCollection();

                        if (docResult == null || docResult.Count == 0)
                        {
                            logger.LogWarning($"{DateTime.Now} CheckAndGenerateExcelAsync: Document contains no extracted values, skipping...");
                            await aiForgedClient.SetDocumentStatusErrorAsync(doc.Id);
                            continue;
                        }

                        excelGenerator.GenerateExcelFile(config.InputTemplatePath, 
                            Path.Combine(config.OutputPath, doc.Filename.Replace(doc.FileType, ".xlsx")), 
                            docResult.FirstOrDefault(d => d.ParamDef.Name.Equals(config.MasterParamDefName))?.Children);

                        await aiForgedClient.SetDocumentStatusCompletedAsync(doc.Id);

                        logger.LogInformation($"{DateTime.Now} CheckAndGenerateExcelAsync: Done processing document {doc}...");
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"{DateTime.Now} CheckAndGenerateExcelAsync: Document Processing Exception: Document: {doc} | {ex}");
                        await aiForgedClient.SetDocumentStatusErrorAsync(doc.Id);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{DateTime.Now} CheckAndGenerateExcelAsync: Exception: {ex}");
            }
        }
    }
}
