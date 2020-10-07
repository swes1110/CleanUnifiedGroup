using System;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;
using Microsoft.Identity.Client;
using System.Collections.ObjectModel;

namespace CleanUnifiedGroup
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).Wait();

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Hit any key to exit...");
                Console.ReadKey();
            }
        }

        static async System.Threading.Tasks.Task MainAsync(string[] args)
        {
            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ConfigurationManager.AppSettings["appId"],
                TenantId = ConfigurationManager.AppSettings["tenantId"],
                RedirectUri = ConfigurationManager.AppSettings["authRedirectURI"]
            };

            var pca = PublicClientApplicationBuilder.CreateWithApplicationOptions(pcaOptions).Build();

            var ewsScopes = new string[] { "https://outlook.office.com/EWS.AccessAsUser.All" };

            try
            {
                //Make the interactive token request
                var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService();
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);

                // Bind to wellknown folder inbox
                string groupAddress = ConfigurationManager.AppSettings["smtpAddressOfUnifiedGroup"];
                Console.WriteLine("Unified Group: {0}",groupAddress);
                FolderId inboxId = new FolderId(WellKnownFolderName.Inbox, groupAddress);
                Folder inbox = Folder.Bind(ewsClient, inboxId);
                Console.WriteLine("Found {0} items in the INBOX", inbox.TotalCount);

                //Bind to wellknown deleted items folder
                FolderId deletedItemsId = new FolderId(WellKnownFolderName.DeletedItems, groupAddress);
                Folder deletedItems = Folder.Bind(ewsClient, deletedItemsId);

                //Find items older than x
                TimeSpan tsOneDay = new TimeSpan(1, 0, 0, 0);
                FindItemsResults<Item> oldItems = inbox.FindItems(new SearchFilter.IsLessThanOrEqualTo(
                    ItemSchema.DateTimeReceived, DateTime.Now.Subtract(tsOneDay)), new ItemView(1000));
                Console.WriteLine("Found [{0}] old items", oldItems.TotalCount);

                //Loop while items older than time span exist in folder
                while (oldItems.TotalCount > 0)
                {
                    //Instantiate collection of ItemIds
                    Collection<ItemId> messageItems = new Collection<ItemId>();
                    foreach (Item item in oldItems)
                    {
                        //Add ItemIds to collection
                        messageItems.Add(item.Id);
                    }

                    //Call BatchDeleteEmailItems function
                    Console.WriteLine("Deleting {0} messages", messageItems.Count);
                    BatchDeleteEmailItems(ewsClient, messageItems);

                    while (deletedItems.TotalCount > 0)
                    {
                        //Empty Deleted Items
                        Console.WriteLine("Emptying Deleted Items Folder");
                        //Update DeletedItems count
                        deletedItems = Folder.Bind(ewsClient, deletedItemsId);
                        Console.WriteLine("There are {0} items in deleted items folder", deletedItems.TotalCount);
                        FindItemsResults<Item> rsDeletedItems = deletedItems.FindItems(new ItemView(1000));
                        Collection<ItemId> colDeletedItemsId = new Collection<ItemId>();
                        foreach (Item deletedItem in rsDeletedItems)
                        {
                            //Add ItemIDs to collection
                            colDeletedItemsId.Add(deletedItem.Id);
                        }

                        //Call BatchDeleteEmailItems function
                        BatchDeleteDeletedItems(ewsClient, colDeletedItemsId);
                    }

                    //Update oldItems variable
                    //Find items older than x
                    oldItems = inbox.FindItems(new SearchFilter.IsLessThanOrEqualTo(
                        ItemSchema.DateTimeReceived, DateTime.Now.Subtract(tsOneDay)), new ItemView(1000));
                    Console.WriteLine("Found [{0}] old items", oldItems.TotalCount);
                }
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex.ToString()}");
                Console.WriteLine("Press return to exit");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
                Console.WriteLine("Press return to exit");
                Console.ReadLine();
            }
        }

        public static void BatchDeleteEmailItems(ExchangeService service, Collection<ItemId> itemIds)
        {
            // Delete the batch of email message objects.
            // This method call results in an DeleteItem call to EWS.
            ServiceResponseCollection<ServiceResponse> response = service.DeleteItems(itemIds, DeleteMode.HardDelete, null, AffectedTaskOccurrence.AllOccurrences);

            // Check for success of the DeleteItems method call.
            // DeleteItems returns success even if it does not find all the item IDs.
            if (response.OverallResult == ServiceResult.Success)
            {
                Console.WriteLine("Email messages deleted successfully.\r\n");
            }
            // If the method did not return success, print a message.
            else
            {
                Console.WriteLine("Not all email messages deleted successfully.\r\n");
            }
        }

        public static void BatchDeleteDeletedItems(ExchangeService service, Collection<ItemId> itemIds)
        {
            // Delete the batch of email message objects.
            // This method call results in an DeleteItem call to EWS.
            ServiceResponseCollection<ServiceResponse> response = service.DeleteItems(itemIds, DeleteMode.HardDelete, null, AffectedTaskOccurrence.AllOccurrences);

            // Check for success of the DeleteItems method call.
            // DeleteItems returns success even if it does not find all the item IDs.
            if (response.OverallResult == ServiceResult.Success)
            {
                Console.WriteLine("Email messages deleted successfully.\r\n");
            }
            // If the method did not return success, print a message.
            else
            {
                Console.WriteLine("Not all email messages deleted successfully.\r\n");
            }
        }
    }
}