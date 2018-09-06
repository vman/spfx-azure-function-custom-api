using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace UserDetails
{
    public static class CurrentUserFromGraph
    {
        private static HttpClient httpClient = new HttpClient();

        [FunctionName("CurrentUserFromGraph")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string ClientId = "<client-id-of-custom-aad-app>";
            string ClientSecret = "<client-secret-of-custom-aad-app>";

            string msGraphResourceUrl = "https://graph.microsoft.com";

            //Get the tenant id from the current claims
            string tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid")?.Value;
            string authority = $"https://login.microsoftonline.com/{tenantId}";

            //Access token coming from SPFx AadHttpClient with "user_impersonation" Scope
            var userImpersonationAccessToken = req.Headers.Authorization.Parameter;

            //Exchange the SPFx access token with another Access token containing the delegated scope for Microsoft Graph
            ClientCredential clientCred = new ClientCredential(ClientId, ClientSecret);
            UserAssertion userAssertion = new UserAssertion(userImpersonationAccessToken);
            //For production, use a Token Cache like Redis https://blogs.msdn.microsoft.com/mrochon/2016/09/19/using-redis-as-adal-token-cache/
            var authContext = new AuthenticationContext(authority);
            AuthenticationResult authResult = await authContext.AcquireTokenAsync(msGraphResourceUrl, clientCred, userAssertion);
            var graphAccessToken = authResult.AccessToken;

            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var request = new HttpRequestMessage(HttpMethod.Get, $"https://graph.microsoft.com/v1.0/me");

            var response = await httpClient.SendAsync(request);

            var content = await response.Content.ReadAsStringAsync();

            var result = new Dictionary<string, string>();
            result.Add("Current User through Microsoft Graph", content);
            return req.CreateResponse(HttpStatusCode.OK, result, JsonMediaTypeFormatter.DefaultMediaType);

        }
    }
}
