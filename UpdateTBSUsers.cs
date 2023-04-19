using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System;

namespace appsvc_fnc_CallMSGraph
{
    public static class UpdateTBSUsers
    {
        [FunctionName("UpdateTBSUsers")]
        public static async Task Run(
            //[HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
            [TimerTrigger("0 0 1 * * *")] TimerInfo myTimer, ILogger log)
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
           
        }

        public static async Task<bool> CallMSFunction(GraphServiceClient graphServiceClient, string tbsgroup, string newdepartment, ILogger log)
        {
            var groupMembers = await graphServiceClient.Groups[tbsgroup].Members.Request()
                    .Header("ConsistencyLevel", "eventual")
                    .Select("department,id,displayName")
                    .GetAsync();
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
                    log.LogInformation(user.Department);
                if (user.Department != newdepartment)
                {
                    user.Department = newdepartment;
                    await graphServiceClient.Users[user.Id].Request().UpdateAsync(user);
                    log.LogInformation($"Updated department for user {user.DisplayName}");
                }
            }
            return true;
        }
            
    }
}




