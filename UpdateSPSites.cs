using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.SharePoint.Client;
using System.Linq;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using PnP.Framework;
using Microsoft.Graph;
using Group = Microsoft.SharePoint.Client.Group;
using Microsoft.Extensions.Configuration;

namespace appsvc_fnc_CallMSGraph
{
    public static class UpdateSPSites
    {
        [FunctionName("UpdateSPSites")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {

            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");
            var appOnlyId = config["appOnlyId"];
            var appOnlySecret = config["appOnlySecret"];
            var assignedGroupName = config["assignedGroupName"];
            var excludeIds = config["excludeIds"];

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            var result = await CallMSFunction(graphAPIAuth, assignedGroupName, appOnlyId, appOnlySecret, excludeIds, log);
           
            string responseMessage = result
                ? "Work as it should"
                : $"Something went wrong. Check the logs";

            return new OkObjectResult(responseMessage);
        }

        public static async Task<bool> CallMSFunction(GraphServiceClient graphServiceClient, string assignedGroupName, string appOnlyId, string appOnlySecret, string excludeIds, ILogger log)
        {
            bool result = false;
            log.LogInformation($"in callmsfunction");
            var siteQueryOptions = new List<QueryOption>()
            {
                new QueryOption("search", "*")
            };
            var sites = await graphServiceClient.Sites.Request(siteQueryOptions).GetAsync();
            log.LogInformation($"after get sites");

            foreach (var site in sites)
            {
                string[] siteIds = site.Id.Split(',');

                if (siteIds.Any(x => excludeIds.Contains(x)) == false)
                {
                    string siteURL = site.WebUrl;

                    try
                    {
                        // SharePoint App only
                        ClientContext ctx = new PnP.Framework.AuthenticationManager().GetACSAppOnlyContext(siteURL, appOnlyId, appOnlySecret, AzureEnvironment.Production);
                        log.LogInformation($"after apponly identification");

                        Web web = ctx.Web;

                        Microsoft.SharePoint.Client.User adGroup = ctx.Web.EnsureUser(assignedGroupName);

                        ctx.Load(adGroup);

                        Group spGroup = ctx.Web.AssociatedMemberGroup;
                        spGroup.Users.AddUser(adGroup);

                        Microsoft.SharePoint.Client.RoleDefinition writeDefinition = ctx.Web.RoleDefinitions.GetByName("Read");
                        RoleDefinitionBindingCollection roleDefCollection = new RoleDefinitionBindingCollection(ctx);
                        roleDefCollection.Add(writeDefinition);
                        Microsoft.SharePoint.Client.RoleAssignment newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                        ctx.Load(spGroup, x => x.Users);
                        ctx.ExecuteQuery();

                        result = true;
                    }
                    catch (ServiceException ex)
                    {
                        log.LogInformation($"Error in msgraph function : {ex.Message}");
                        result = false;
                    };
                }              
            }            
            return result;
        }
    }
}



