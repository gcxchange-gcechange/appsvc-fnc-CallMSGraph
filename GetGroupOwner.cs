using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.WindowsAzure.Storage;
using Microsoft.Extensions.Configuration;
using System.Text;
using Azure.Storage.Blobs;
using Azure.Identity;

namespace appsvc_fnc_CallMSGraph
{

    // Is it possible for us to get a list of all the community owners
    // We are gathering users to do more research on ProB to ensure that it aligns with what users want and envision from the Pro B functionality.
    // we just need to know the current list to start recruiting. 
    // The email is the most important information, but if we can add the name of the sites they own next to them, that will be ideal

    public static class GetGroupOwner
    {
        [FunctionName("GetGroupOwner")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string AzureWebJobsStorage = config["AzureWebJobsStorage"];

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            
            var result = await GetGroupList(graphAPIAuth, log);

            //await StoreGroupList(graphAPIAuth, AzureWebJobsStorage, log);

            return new OkResult();
        }
        public static async Task<bool> GetGroupList(GraphServiceClient graphServiceClient, ILogger log)
        {
            bool result = true;
            string groupInfo;
            string ownerInfo;

            var groups = await graphServiceClient.Groups.Request().GetAsync();

            foreach (var group in groups)
            {
                ownerInfo = string.Empty;

                var owners = await graphServiceClient.Groups[group.Id].Owners.Request().GetAsync();
                foreach (var owner in owners)
                {
                    var user = await graphServiceClient.Users[owner.Id].Request().GetAsync();
                    ownerInfo = string.Concat(ownerInfo, $"{user.DisplayName}, {user.Mail}; ");
                }

                groupInfo = $"{group.DisplayName}: {ownerInfo}";
                log.LogInformation(groupInfo);
            }

            return result;
        }

        public static async Task<bool> StoreGroupList(GraphServiceClient graphServiceClient,string AzureWebJobsStorage, ILogger log)
        {
            bool result = true;

            string accountName = "dstgraphstd";
            string containerName = "community-list";
            string containerEndpoint = string.Format("https://{0}.blob.core.windows.net/{1}", accountName, containerName);

            //CloudStorageAccount storageAccount = CloudStorageAccount.Parse(AzureWebJobsStorage);

            BlobContainerClient containerClient = new BlobContainerClient(new Uri(containerEndpoint), new DefaultAzureCredential());
            // containerClient.Properties.ContentType = "application/json";

            var stringInsideTheFile = $"{{\"B2BGroupSyncAlias\": \"group_alias\",\"groupAliasToUsersMapping\":{{ resultUserList }} }}";
            byte[] byteArray = Encoding.ASCII.GetBytes(stringInsideTheFile);

            try {
                using (MemoryStream stream = new MemoryStream(byteArray))
                {
                    await containerClient.UploadBlobAsync("my_file_name.json", stream);
                }
            }
            catch (Exception e) {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException != null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            return result;
        }

    }
}