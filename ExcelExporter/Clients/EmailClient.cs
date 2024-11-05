using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Azure.Identity;

using ExcelExporter.Models;

using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.Identity.Client;

namespace ExcelExporter.Clients
{
    public partial class EmailClient
    {
        private IConfig _config;
        private ILogger _logger;

        ClientSecretCredential _credential;
        GraphServiceClient _graphClient;

        public EmailClient(ILogger logger, IConfig config)
        {
            _config = config;
            _logger = logger;

            Init();
        }

        private void Init()
        {
            _credential = new ClientSecretCredential(_config.EmailTenantId, _config.EmailClientId, _config.EmailClientSecret);
            _graphClient = new GraphServiceClient(_credential);
        }

        public async Task SendEmailAsync(string subject, string body, List<Attachment> attachments = null, string[] bcc = null, params string[] recipients)
        {
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
                },
                ToRecipients = new List<Recipient>(),
                Attachments = attachments
            };

            foreach (string recipient in recipients)
            {
                message.ToRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient
                    }
                });
            }

            if (bcc is not null)
            {
                foreach (string recipient in bcc)
                {
                    message.BccRecipients.Add(new Recipient()
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    });
                }
            }

            var sendMailPostRequestBody = new SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = false
            };

            var user = await _graphClient.Users[_config.EmailFromAddress].GetAsync();

            await _graphClient.Users[user.Id].SendMail.PostAsync(sendMailPostRequestBody);
        }
    }
}
