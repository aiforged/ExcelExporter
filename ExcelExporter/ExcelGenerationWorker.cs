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

using Microsoft.Graph.Models;

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
                ExcelGenerator excelGenerator = new ExcelGenerator(logger, config);
                AIForgedClient aiForgedClient = new AIForgedClient(logger, config);
                EmailClient emailClient = new EmailClient(logger, config);

                List<string> excelFiles = new List<string>();

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

                        string outputPath = Path.Combine(config.OutputPath, $"{doc.Filename.Replace(doc.FileType, string.Empty)}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");

                        excelGenerator.GenerateExcelFile(config.InputTemplatePath,
                            outputPath,
                            docResult);

                        await aiForgedClient.SetDocumentStatusCompletedAsync(doc.Id);
                        await aiForgedClient.SetDocumentStatusCompletedAsync(doc.MasterId);

                        excelFiles.Add(outputPath);

                        logger.LogInformation($"{DateTime.Now} CheckAndGenerateExcelAsync: Done processing document {doc}...");
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"{DateTime.Now} CheckAndGenerateExcelAsync: Document Processing Exception: Document: {doc} | {ex}");
                        await aiForgedClient.SetDocumentStatusErrorAsync(doc.Id);
                        await aiForgedClient.SetDocumentStatusErrorAsync(doc.MasterId);
                    }
                }

                logger.LogInformation($"{DateTime.Now} CheckAndGenerateExcelAsync: Done processing {docs.Count} docs...");

                if (excelFiles is null || excelFiles.Count == 0) return;
                if (config.EmailRecipients is null || config.EmailRecipients.Count == 0) return;

                if (config.SendBulkEmail)
                {
                    List<Attachment> attachments = new List<Attachment>();

                    foreach (var file in excelFiles)
                    {
                        attachments.Add(new FileAttachment()
                        {
                            Name = Path.GetFileName(file),
                            ContentBytes = File.ReadAllBytes(file),
                            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        });
                    }

                    await emailClient.SendEmailAsync("AIForged Processed Documents",
                        "",
                        attachments,
                        config.EmailBccRecipents.ToArray(),
                        config.EmailRecipients.ToArray());
                }
                else
                {
                    foreach (var file in excelFiles)
                    {
                        await emailClient.SendEmailAsync($"AIForged Processed Document - {Path.GetFileName(file)}",
                        "",
                        [new FileAttachment()
                        {
                            Name = Path.GetFileName(file),
                            ContentBytes = File.ReadAllBytes(file),
                            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        }],
                        config.EmailBccRecipents.ToArray(),
                        config.EmailRecipients.ToArray());
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
