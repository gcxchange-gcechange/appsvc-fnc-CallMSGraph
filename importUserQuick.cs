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
using System.IO;
using Microsoft.SharePoint.News.DataModel;
using Newtonsoft.Json;
using System;

namespace appsvc_fnc_CallMSGraph
{
    public static class ImportUserQuick
    {
        [FunctionName("ImportUserQuick")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {

            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");

            var assignedGroupName = config["assignedGroupID"];
            var welcomeGroup = config["welcomeGroup"];
            var CIOGroup = config["CIOGroup"];
            var redirectLink = config["redirectLink"];
            var UsersList = await new StreamReader(req.Body).ReadToEndAsync();

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);
            var result = await ImportBulkUserQuick(graphAPIAuth, UsersList, assignedGroupName, welcomeGroup, CIOGroup, redirectLink, log);
           
            string responseMessage = result
                ? "Work as it should"
                : $"Something went wrong. Check the logs";

            return new OkObjectResult(responseMessage);
        }

        public static async Task<bool> ImportBulkUserQuick(GraphServiceClient graphServiceClient,string UsersList, string assignedGroupName, string welcomeGroup, string CIOGroup, string redirectLink, ILogger log)
        {
                List<Item> items = JsonConvert.DeserializeObject<List<Item>>(UsersList);

                foreach (var user in items)
                {
                    log.LogInformation($"{user.FirstName}");
                    try
                    {
                        var invitation = new Invitation
                        {
                            SendInvitationMessage = false,
                            InvitedUserEmailAddress = user.Email,
                            InvitedUserType = "Member",
                            InvitedUserDisplayName = $"{user.FirstName} {user.LastName}",
                            InviteRedirectUrl = redirectLink
                        };

                        var userInvite = await graphServiceClient.Invitations.Request().AddAsync(invitation);

                        var userID = userInvite.InvitedUser.Id;

                        try
                        {
                            var guestUser = new Microsoft.Graph.User
                            {
                                UserType = "Member"
                            };

                            await graphServiceClient.Users[userID].Request().UpdateAsync(guestUser);
                            log.LogInformation("User update successfully");

                        }
                        catch (Exception ex)
                        {
                            log.LogInformation($"Error Updating User : {ex.Message}");
                        }

                        try
                        {
                            var directoryObject = new DirectoryObject
                            {
                                Id = userID
                            };
                            await graphServiceClient.Groups[CIOGroup].Members.References
                                .Request()
                                .AddAsync(directoryObject);
                            log.LogInformation("User add to CIO group group successfully");
                            await graphServiceClient.Groups[assignedGroupName].Members.References
                                .Request()
                                .AddAsync(directoryObject);
                            log.LogInformation("User add to GCX_Assigned group successfully");
                            await graphServiceClient.Groups[welcomeGroup].Members.References
                                .Request()
                                .AddAsync(directoryObject);
                            log.LogInformation("User add to welcome group successfully");
                        }
                        catch (Exception ex)
                        {
                            log.LogInformation($"Error adding User to groups : {ex.Message}");
                        }
                        log.LogInformation(@"User invite successfully - {userInvite.InvitedUser.Id}");
                    }
                    catch (ServiceException ex)
                    {
                        log.LogInformation($"Error Creating User Invite : {ex.Message}");
                    }
                }
           return true;
        }
        public class Item
        {
            public string FirstName;
            public string LastName;
            public string Email;
        }
    }
}



