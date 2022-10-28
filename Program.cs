using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsGraphApi
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var client = GetAuthenticatedGraphClient();

            var graphRequest = client.Users.Request();

            var results = graphRequest.GetAsync().Result;

            //Display User Id, Name, Email
            foreach (var user in results)
            {
                Console.WriteLine(user.Id + ": " + user.DisplayName + " | " + user.Mail);
            }
            string domainName = "xxxxxxx.sharepoint.com";
            string siteId = "xxxxxxxxxxxxxxxxxxxx";
            string webId = "xxxxxxxxxxxxxxxxxxxxx";

            string siteDetail = string.Format("{0},{1},{2}", domainName, siteId, webId);
            var lists = client.Sites[siteDetail].Lists
                                .Request()
                                .GetAsync().Result;

            //Display all lists name in the site
            foreach (var list in lists)
            {
                Console.WriteLine("lists : " + list.Name);
            }
            Console.ReadLine();

        }

        private static IAuthenticationProvider CreateAuthorizationProvider()
        {
            string clientId = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
            string clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
            string tenantId = "xxxxxxxxxxxxxxxxxxxxxxxxxxxx";
            string authority = string.Format("https://login.microsoftonline.com/{0}/v2.0", tenantId);

            string redirectUri = "https://xxxxxxxx.sharepoint.com/sites/contoso";
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient()
        {
            var authenticationProvider = CreateAuthorizationProvider();
            var graphClient = new GraphServiceClient(authenticationProvider);
            return graphClient;
        }
    }
}
