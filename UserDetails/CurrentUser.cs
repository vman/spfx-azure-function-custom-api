using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Security.Claims;

namespace UserDetails
{
    public static class CurrentUser
    {
        [FunctionName("CurrentUser")]
        public static HttpResponseMessage Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            var result = new Dictionary<string, string>();

            result.Add("source", "Authenticated Azure Function!");

            //Current user claims
            foreach (Claim claim in ClaimsPrincipal.Current.Claims)
            {
                result.Add(claim.Type, claim.Value);
            }

            return req.CreateResponse(HttpStatusCode.OK, result, JsonMediaTypeFormatter.DefaultMediaType);
        }
    }
}
