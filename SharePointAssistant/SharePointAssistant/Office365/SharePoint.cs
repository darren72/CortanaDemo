using Microsoft.Office365.Discovery;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.UI.Xaml;
using Windows.Web.Http;
using Windows.Web.Http.Headers;

namespace SharePointAssistant.Office365
{
    public static class SharePoint
    {
        public static async Task<List<string>> GetListItems(string listName)
        {
            // Use the discovery service to get the URL of the root site and an access token to talk to it.
            string accessToken = await GetAccessTokenForResource("https://api.office.com/discovery/");
            DiscoveryClient discoveryClient = new DiscoveryClient(() => accessToken);
            CapabilityDiscoveryResult result = await discoveryClient.DiscoverCapabilityAsync("RootSite");
            var sharePointAccessToken = await GetAccessTokenForResource(result.ServiceResourceId);
            var sharePointServiceEndpointUri = result.ServiceEndpointUri.ToString();

            // Construct an HTTP request to bring back te contents of the Announcements list in the root site.
            string url = string.Format("{0}/web/lists/GetByTitle('{1}')/items", sharePointServiceEndpointUri, listName);
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, new Uri(url));
            request.Headers.Add("Accept", "text/xml");

            // Bring back the list of announcements
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new HttpCredentialsHeaderValue("Bearer", sharePointAccessToken);
            var response = await client.SendRequestAsync(request);

            // Quick fix to bring back just the titles - would need to be parsed properly in a real app!
            XElement reponseAsXml = XElement.Parse(response.Content.ToString());
            XNamespace dNameSpace = "http://schemas.microsoft.com/ado/2007/08/dataservices";
            return reponseAsXml.Descendants(dNameSpace + "Title")
                           .Select(x => (string)x)
                           .ToList();
        }

        /// <summary>
        /// Attempts to get an access token for a resource. Siltently if possible, otherwise displays Office 365 login screen.
        /// </summary>
        /// <param name="resource">The name of the resource to get the access token for.</param>
        /// <returns>The access token.</returns>
        public static async Task<string> GetAccessTokenForResource(string resource)
        {
            string token = null;

            //first try to get the token silently
            WebAccountProvider aadAccountProvider
                = await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.windows.net");
            WebTokenRequest webTokenRequest
                = new WebTokenRequest(aadAccountProvider, String.Empty, Application.Current.Resources["ida:ClientId"].ToString(), WebTokenRequestPromptType.Default);
            webTokenRequest.Properties.Add("authority", "https://login.windows.net");
            webTokenRequest.Properties.Add("resource", resource);
            WebTokenRequestResult webTokenRequestResult
                = await WebAuthenticationCoreManager.GetTokenSilentlyAsync(webTokenRequest);
            if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success)
            {
                WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
                token = webTokenResponse.Token;
            }
            else if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.UserInteractionRequired)
            {
                //get token through prompt
                webTokenRequest
                    = new WebTokenRequest(aadAccountProvider, String.Empty, Application.Current.Resources["ida:ClientId"].ToString(), WebTokenRequestPromptType.ForceAuthentication);
                webTokenRequest.Properties.Add("authority", "https://login.windows.net");
                webTokenRequest.Properties.Add("resource", resource);
                webTokenRequestResult
                    = await WebAuthenticationCoreManager.RequestTokenAsync(webTokenRequest);
                if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success)
                {
                    WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
                    token = webTokenResponse.Token;
                }
            }

            return token;
        }
    }
}
