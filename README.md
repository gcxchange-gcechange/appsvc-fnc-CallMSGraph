# appsvc-fnc-CallMSGraph

This is use as a script and it get trigger on call.

This script add an azure group to all site (on exception) with read access.

Site seeting that need to be add to the function app

clientSecret = client secret of app registration
clientId = client id of app registration with site full controle api permission
tenantid = Id of the tenant
appOnlyId =  App only id with
appOnlySecret = App only secret full site controle permisson
assignedGroupName = Group that will be add to all site with read access
excludeIds = list of sharepoint site id that will not apply the new group
