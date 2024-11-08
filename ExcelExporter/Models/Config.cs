﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using AIForged.API;

namespace ExcelExporter.Models
{
    public interface IConfig
    {
        string APIKey { get; set; }
        int ProjectId { get; set; }
        int ServiceId { get; set; }
        string AIForgedEndpoint { get; set; }

        string InputTemplatePath { get; set; }
        string OutputPath { get; set; }

        string MasterParamDefName { get; set; }
        float ValueCompareConfidence { get; set; }

        DocumentStatus InputDocumentStatus { get; set; }
        DocumentStatus ProcessedDocumentStatus { get; set; }

        string EmailClientId { get; set; }
        string EmailTenantId { get; set; }
        string EmailClientSecret { get; set; }
        string EmailFromAddress { get; set; }
        List<string> EmailRecipients { get; set; }
        List<string> EmailBccRecipents { get; set; }
        bool SendBulkEmail {  get; set; }
    }

    public class Config : IConfig
    {
        public Config() { }

        public string APIKey { get; set; }
        public int ProjectId { get; set; }
        public int ServiceId { get; set; }
        public string AIForgedEndpoint { get; set; }

        public string InputTemplatePath { get; set; }
        public string OutputPath { get; set; }

        public string MasterParamDefName { get; set; }
        public float ValueCompareConfidence { get; set; } = 0.9F;

        public DocumentStatus InputDocumentStatus { get; set; } = DocumentStatus.InterimProcessed;
        public DocumentStatus ProcessedDocumentStatus { get; set; } = DocumentStatus.Processed;

        public string EmailClientId { get; set; }
        public string EmailTenantId { get; set; }
        public string EmailClientSecret { get; set; }
        public string EmailFromAddress { get; set; }
        public List<string> EmailRecipients { get; set; }
        public List<string> EmailBccRecipents { get; set; }
        public bool SendBulkEmail { get; set; } = false;
    }
}
