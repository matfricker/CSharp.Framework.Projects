using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;

namespace Console.EWS.Example
{
    internal class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            // using Microsoft.Identity.Client
            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(ConfigurationManager.AppSettings["ClientId"])
                .WithClientSecret(ConfigurationManager.AppSettings["ClientSecret"])
                .WithTenantId(ConfigurationManager.AppSettings["TenantId"])
                .Build();

            var exchangeScopes = new string[] { "https://outlook.office365.com/.default" };

            try
            {
                const int maxFolders = 25;
                const int maxItems = 10;

                var authResult = await confidentialClientApplication.AcquireTokenForClient(exchangeScopes).ExecuteAsync();

                // configure the ExchangeService with the access token
                var exchangeService = new ExchangeService
                {
                    Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
                    Credentials = new OAuthCredentials(authResult.AccessToken),
                    ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, ConfigurationManager.AppSettings["ImpersonatedUser"])
                };

                // include x-anchormailbox header
                exchangeService.HttpHeaders.Add("X-AnchorMailbox", ConfigurationManager.AppSettings["ImpersonatedUser"]);

                // make an EWS call
                var folders = exchangeService.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(maxFolders));

                foreach (var folder in folders)
                {
                    if (folder.DisplayName != WellKnownFolderName.Inbox.ToString()) continue;

                    WriteLine($"Folder: {folder.DisplayName}");

                    var inboxMessages = exchangeService.FindItems(folder.Id, new ItemView(maxItems));
                    if (inboxMessages.TotalCount > 0)
                        WriteLine($"Count:{inboxMessages.TotalCount}");

                    var inboxSubfolders = exchangeService.FindFolders(folder.Id, new FolderView(maxFolders));
                    if (inboxSubfolders.TotalCount <= 0) continue;

                    foreach (var subfolder in inboxSubfolders)
                    {
                        WriteLine($"> Sub-Folder: {subfolder.DisplayName}");
                        var subfolderMessages = exchangeService.FindItems(subfolder.Id, new ItemView(maxItems));

                        if (subfolderMessages.TotalCount <= 0) continue;

                        WriteLine($"  Count: {subfolderMessages.TotalCount}");
                        var item = subfolderMessages.Items[0];
                        if (!item.HasAttachments) continue;

                        var emailMessage = EmailMessage.Bind(exchangeService, item.Id, new PropertySet(ItemSchema.Attachments));
                        foreach (var attachment in emailMessage.Attachments)
                        {
                            WriteLine($"  Attachment: {attachment.Name}");
                        }
                    }
                }
            }
            catch (MsalException ex)
            {
                WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                WriteLine($"Error: {ex}");
            }

            if (System.Diagnostics.Debugger.IsAttached)
            {
                WriteLine("");
                WriteLine("Hit any key to exit...");
                System.Console.ReadKey();
            }
        }

        private static void WriteLine(string message)
        {
            System.Console.WriteLine(message);
        }
    }
}
