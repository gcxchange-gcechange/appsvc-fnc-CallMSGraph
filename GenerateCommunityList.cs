using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Text;
using Azure.Storage.Blobs;
using Azure.Identity;

namespace appsvc_fnc_CallMSGraph
{
    public static class GenerateCommunityList
    {
        public struct Community
        {
            public string DisplayName;
            public List<Owner> OwnerInfo;

            public Community() {
                DisplayName = string.Empty;
                OwnerInfo = new List<Owner>();  
            }
        }
        public struct Owner
        {
            public string DisplaName;
            public string EmailAddress;
        }

        [FunctionName("GenerateCommunityList")]
        public static async Task<bool> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            List<Community> communityList = await GetCommunityList(graphAPIAuth, log);

            if (communityList != null)
            {
                var result = await StoreCommunityList(communityList, log);
                return result;
            }
            else
                return false;
        }
        public static async Task<List<Community>> GetCommunityList(GraphServiceClient graphServiceClient, ILogger log)
        {
            List<Community> communityList = new List<Community>();
            Community community;
            Owner owner;

            try {
                var groups = await graphServiceClient.Groups.Request(new List<QueryOption>() { new QueryOption("$count", "true") }).Header("ConsistencyLevel", "eventual").Filter("groupTypes/any(c:c eq 'Unified')").OrderBy("displayName asc").GetAsync();
                

                foreach (var group in groups)
                {
                    community = new Community();
                    community.DisplayName = group.DisplayName;

                    var owners = await graphServiceClient.Groups[group.Id].Owners.Request().GetAsync();
                    foreach (var o in owners)
                    {
                        var user = await graphServiceClient.Users[o.Id].Request().GetAsync();
                        owner = new Owner();
                        owner.DisplaName = user.DisplayName;
                        owner.EmailAddress = user.Mail;
                        community.OwnerInfo.Add(owner);
                    }

                    communityList.Add(community);
                }
                while (groups.NextPageRequest != null && (groups = await groups.NextPageRequest.GetAsync()).Count > 0) ;
            }
            catch (Exception e) {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException != null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                return null;
            }

            return communityList;
        }

        public static async Task<bool> StoreCommunityList(List<Community> communityList, ILogger log)
        {
            bool result = true;

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string storageAccountName = config["storageAccountName"];
            string containerName = config["containerName"];
            string containerEndpoint = string.Format("https://{0}.blob.core.windows.net/{1}", storageAccountName, containerName);
            string filename = $"community_list_{DateTime.Now.ToString("yyyy-MM-dd")}.json";

            BlobContainerClient containerClient = new BlobContainerClient(new Uri(containerEndpoint), new DefaultAzureCredential());

            byte[] byteArray = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(communityList));

            try {
                await containerClient.DeleteBlobIfExistsAsync(filename);

                using (MemoryStream stream = new MemoryStream(byteArray))
                {
                    await containerClient.UploadBlobAsync(filename, stream);
                }
            }
            catch (Exception e) {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException != null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            return result;
        }

    }
}