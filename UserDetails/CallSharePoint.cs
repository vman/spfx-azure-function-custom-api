using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace UserDetails
{
    public static class CallSharePoint
    {
        [FunctionName("CurrentUserFromSharePoint")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string ClientId = "<client-id-of-custom-aad-app>";
            string ClientSecret = "<client-secret-of-custom-aad-app>";

            string spRootResourceUrl = "https://tenant.sharepoint.com";
            string spSiteUrl = $"{spRootResourceUrl}/sites/comms";

            //Get the tenant id from the current claims
            string tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;
            string authority = $"https://login.microsoftonline.com/{tenantId}";

            //Access token coming from SPFx AadHttpClient with "user_impersonation" Scope
            var userImpersonationAccessToken = req.Headers.Authorization.Parameter;

            //Exchange the SPFx access token with another Access token containing the delegated "AllSites.Manage" scope for the SharePoint resource
            ClientCredential clientCred = new ClientCredential(ClientId, ClientSecret);
            UserAssertion userAssertion = new UserAssertion(userImpersonationAccessToken);
            //For production, use a Token Cache like Redis https://blogs.msdn.microsoft.com/mrochon/2016/09/19/using-redis-as-adal-token-cache/
            var authContext = new AuthenticationContext(authority);
            AuthenticationResult authResult = await authContext.AcquireTokenAsync(spRootResourceUrl, clientCred, userAssertion);
            var spAccessToken = authResult.AccessToken;

            //Get CSOM ClientContext using the SharePoint Access Token. OfficeDevPnP nuget package is needed for using the extention methods.
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            var clientContext = authManager.GetAzureADAccessTokenAuthenticatedContext(spSiteUrl, spAccessToken);

            //The usual CSOM stuff.
            var web = clientContext.Web;
            var currentUser = web.CurrentUser;
            clientContext.Load(web);
            clientContext.Load(currentUser);
            clientContext.ExecuteQuery();

            var result = new Dictionary<string, string>();
            result.Add("Current Web in SharePoint", web.Title);
            result.Add("Current User in SharePoint", currentUser.Title);
            return req.CreateResponse(HttpStatusCode.OK, result, JsonMediaTypeFormatter.DefaultMediaType);

        }
    }
}
