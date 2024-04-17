using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using AIForged.API;

namespace ExcelExporter.Clients
{
    public class AIForgedClient
    {
        private readonly string apiKey = "";
        ILogger logger;
        ExcelExporter.Models.IConfig config;
        IContext context;

        public AIForgedClient(ILogger logger, ExcelExporter.Models.IConfig config)
        {
            this.logger = logger;
            this.config = config;

            this.apiKey = config.APIKey;

            Init();
        }

        private void Init()
        {
            Config cfg = new Config();

            cfg.BaseUrl = config.AIForgedEndpoint;
            cfg.Init();
            cfg.HttpClient.DefaultRequestHeaders.Add("X-Api-Key", apiKey);

            context = new AIForged.API.Context(cfg);
        }

        public async Task<bool> IsLoggedInAsync()
        {
            bool isLoggedIn = false;

            try
            {
                var userResult = await context.GetCurrentUserAsync();

                if (userResult != null)
                {
                    isLoggedIn = true;
                }
            }
            catch
            {
            }

            return isLoggedIn;
        }

        //Use to get documents with a processed status from specified project and service combination
        public async Task<ICollection<DocumentViewModel>> GetProcessedDocsAsync(int projectId, int serviceId, DocumentStatus documentStatus = DocumentStatus.Processed)
        {
            //Instantiate our list to contain our document information
            List<DocumentViewModel> docs = new List<DocumentViewModel>();

            try
            {
                logger.LogInformation($"{DateTime.Now} GetProcessedDocsAsync: Start");

                if (!(await IsLoggedInAsync()))
                {
                    logger.LogInformation($"{DateTime.Now} GetProcessedDocsAsync: User not logged in, exiting.");
                    return docs;
                }

                //Get all the available result docs with processed status over the past 2 weeks in descending order
                var docsResult = await context.DocumentClient.GetExtendedAsync(context.CurrentUserId,
                    projectId,
                    serviceId,
                    UsageType.Outbox,
                    [documentStatus],
                    null,
                    null,
                    null,
                    DateTime.Today.AddDays(-14),
                    LarcAI.Utilities.EndOfTheDay(DateTime.Today),
                    null,
                    null,
                    null,
                    null,
                    SortField.Id,
                    SortDirection.Descending,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);

                //Check if the request was successful (might actually fail to the catch if the request was not succesful, but we check anyway).
                logger.LogInformation($"{DateTime.Now} GetProcessedDocsAsync: Got result. Checking...");
                if (docsResult.StatusCode >= 200 && docsResult.StatusCode < 300 && docsResult.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} GetProcessedDocsAsync: Got result. Success");
                    docs = docsResult.Result.ToList();
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{DateTime.Now} GetProcessedDocsAsync: Exception: {ex}");
            }

            //Return our docs, if any
            logger.LogInformation($"{DateTime.Now} GetProcessedDocsAsync: End");
            return docs;
        }

        //Use to get document data from specified document id
        public async Task<ICollection<DocumentDataViewModel>> GetDocumentDataAsync(int documentId)
        {
            //Instantiate our list to contain our document information
            List<DocumentDataViewModel> docResults = new List<DocumentDataViewModel>();

            try
            {
                logger.LogInformation($"{DateTime.Now} GetDocumentDataAsync: Start");

                if (!(await IsLoggedInAsync()))
                {
                    logger.LogInformation($"{DateTime.Now} GetDocumentDataAsync: User not logged in, exiting.");
                    return docResults;
                }

                //Get the latest document extraction results for the given document id
                var docResultsResp = await context.DocumentClient.GetDataAsync(documentId, [DocumentDataType.Image], null, null, null, null, null);

                //Check if the request was successful (might actually fail to the catch if the request was not succesful, but we check anyway).
                logger.LogInformation($"{DateTime.Now} GetDocumentDataAsync: Got result. Checking...");
                if (docResultsResp.StatusCode >= 200 && docResultsResp.StatusCode < 300 && docResultsResp.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} GetDocumentDataAsync: Got result. Success");
                    docResults = docResultsResp.Result.ToList();
                }

            }
            catch (Exception ex)
            {
                logger.LogError($"{DateTime.Now} GetDocumentDataAsync: Exception: {ex}");
            }

            //Return our docs, if any
            logger.LogInformation($"{DateTime.Now} GetDocumentDataAsync: End");
            return docResults;
        }

        //Use to get document extraction results from specified project, service and document combination
        public async Task<ICollection<DocumentParameterViewModel>> GetDocumentResultsAsync(int projectId, int serviceId, int documentId)
        {
            //Instantiate our list to contain our document information
            List<DocumentParameterViewModel> docResults = new List<DocumentParameterViewModel>();

            try
            {
                logger.LogInformation($"{DateTime.Now} GetDocumentResultsAsync: Start");

                if (!(await IsLoggedInAsync()))
                {
                    logger.LogInformation($"{DateTime.Now} GetDocumentResultsAsync: User not logged in, exiting.");
                    return docResults;
                }

                //Get the latest document extraction results for the given document id
                var docResultsResp = await context.ParametersClient.GetHierarchyAsync(documentId, 
                    serviceId, 
                    includeverification: true, 
                    pageIndex: null);

                //Check if the request was successful (might actually fail to the catch if the request was not succesful, but we check anyway).
                logger.LogInformation($"{DateTime.Now} GetDocumentResultsAsync: Got result. Checking...");
                if (docResultsResp.StatusCode >= 200 && docResultsResp.StatusCode < 300 && docResultsResp.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} GetDocumentResultsAsync: Got result. Success");
                    docResults = docResultsResp.Result.ToList();
                }

            }
            catch (Exception ex)
            {
                logger.LogError($"{DateTime.Now} GetDocumentResultsAsync: Exception: {ex}");
            }

            //Return our docs, if any
            logger.LogInformation($"{DateTime.Now} GetDocumentResultsAsync: End");
            return docResults;
        }

        //Use to set a document's status to CustomProcessed given its id
        public async Task<bool> SetDocumentStatusCompletedAsync(int? documentId)
        {
            //Initialize our doc object to null used to update the document's status
            DocumentViewModel doc = null;
            //Initialize our response boolean to false
            bool isSuccess = false;

            if (documentId == null) return isSuccess;
            try
            {
                logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: Start");

                if (!(await IsLoggedInAsync()))
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: User not logged in, exiting.");
                    return isSuccess;
                }

                //First get our document so that we can update its status
                var docResp = await context.DocumentClient.GetDocumentAsync(documentId);

                if (docResp.StatusCode >= 200 && docResp.StatusCode < 300 && docResp.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: Got result. Success");
                    doc = docResp.Result;
                }
                else
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: No result. Failed");
                    return isSuccess;
                }

                //Set the status of the given document to CustomProcessed
                doc.Status = config.ProcessedDocumentStatus;

                var docUpdateResp = await context.DocumentClient.UpdateAsync(doc);

                //Check if the request was successful (might actually fail to the catch if the request was not succesful, but we check anyway).
                logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: Got result. Checking...");
                if (docUpdateResp.StatusCode > 200 && docUpdateResp.StatusCode < 300 && docUpdateResp.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: Got result. Success");
                    isSuccess = true;
                }

            }
            catch (Exception ex)
            {
                logger.LogError($"{DateTime.Now} SetDocumentStatusCompletedAsync: Exception: {ex}");
            }

            //Return our docs, if any
            logger.LogInformation($"{DateTime.Now} SetDocumentStatusCompletedAsync: End");
            return isSuccess;
        }

        //Use to set a document's status to CustomError given its id
        public async Task<bool> SetDocumentStatusErrorAsync(int documentId)
        {
            //Initialize our doc object to null used to update the document's status
            DocumentViewModel doc = null;
            //Initialize our response boolean to false
            bool isSuccess = false;

            try
            {
                logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: Start");

                if (!(await IsLoggedInAsync()))
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: User not logged in, exiting.");
                    return isSuccess;
                }

                //First get our document so that we can update its status
                var docResp = await context.DocumentClient.GetDocumentAsync(documentId);

                if (docResp.StatusCode >= 200 && docResp.StatusCode < 300 && docResp.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: Got result. Success");
                    doc = docResp.Result;
                }
                else
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: No result. Failed");
                    return isSuccess;
                }

                //Set the status of the given document to CustomProcessed
                doc.Status = DocumentStatus.CustomError;

                var docUpdateResp = await context.DocumentClient.UpdateAsync(doc);

                //Check if the request was successful (might actually fail to the catch if the request was not succesful, but we check anyway).
                logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: Got result. Checking...");
                if (docUpdateResp.StatusCode > 200 && docUpdateResp.StatusCode < 300 && docUpdateResp.Result != null)
                {
                    logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: Got result. Success");
                    isSuccess = true;
                }

            }
            catch (Exception ex)
            {
                logger.LogError($"{DateTime.Now} SetDocumentStatusErrorAsync: Exception: {ex}");
            }

            //Return our docs, if any
            logger.LogInformation($"{DateTime.Now} SetDocumentStatusErrorAsync: End");
            return isSuccess;
        }
    }
}
