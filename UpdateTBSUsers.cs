using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

namespace appsvc_fnc_CallMSGraph
{
    public static class UpdateTBSUsers
    {
        [FunctionName("UpdateTBSUsers")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {

            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");
            var tbsgroup = config["tbsgroup"];
            var newdepartment = config["rgcodetbs"];

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            var result = await CallMSFunction(graphAPIAuth, tbsgroup, newdepartment, log);
           
            string responseMessage = result
                ? "Work as it should"
                : $"Something went wrong. Check the logs";

            return new OkObjectResult(responseMessage);
        }

        public static async Task<bool> CallMSFunction(GraphServiceClient graphServiceClient, string tbsgroup, string newdepartment, ILogger log)
        {
            var groupMembers = await graphServiceClient.Groups[tbsgroup].Members.Request().GetAsync();
            var users = new List<Microsoft.Graph.User>(); 
            while (groupMembers.Count > 0) 
                { 
                    foreach (var member in groupMembers) 
                    { 
                        if (member is Microsoft.Graph.User user) 
                        { 
                            users.Add(user); 
                        } 
                    } 
                    if (groupMembers.NextPageRequest == null) break; 
                    groupMembers = await groupMembers.NextPageRequest.GetAsync(); 
                }

                foreach (var user in users) 
                    { 
                    user.Department = newdepartment; 
                    await graphServiceClient.Users[user.Id].Request().UpdateAsync(user); 
                    log.LogInformation($"Updated department for user {user.DisplayName}"); 
                    }
            return true;
        }
            
    }
}




