using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OAuth;
using Microsoft.Office365.OutlookServices;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

using MailSim.Common;
using Newtonsoft.Json.Linq;
using System.Net.Http;

namespace MailSim.ProvidersSDK
{
    /// <summary>
    /// Provides clients for the different service endpoints.
    /// </summary>
    internal static class AuthenticationHelper
    {
        private static readonly string ClientID = Resources.ClientID;

        private static readonly Uri ReturnUri = new Uri(Resources.ReturnUri);

        // Properties used for communicating with your Windows Azure AD tenant.
        private static readonly string CommonAuthority = "https://login.microsoftonline.com/Common";
        private static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        private const string DiscoveryResourceId = "https://api.office.com/discovery/";
        private const string AadServiceResourceId = "https://graph.windows.net/";

        private const string ModuleName = "AuthenticationHelper";

        //Static variables store the clients so that we don't have to create them more than once.
        private static ActiveDirectoryClient _graphClient = null;
        private static OutlookServicesClient _outlookClient = null;

        //Property for storing the authentication context.
        private static AuthenticationContext _authenticationContext { get; set; }

        //Property for storing and returning the authority used by the last authentication.
        private static string LastAuthority { get; set; }
        //Property for storing the tenant id so that we can pass it to the ActiveDirectoryClient constructor.
        private static string TenantId { get; set; }
        // Property for storing the logged-in user so that we can display user properties later.
        internal static string LoggedInUser { get; set; }

        /// <summary>
        /// Checks that a Graph client is available.
        /// </summary>
        /// <returns>The Graph client.</returns>
        public static async Task<ActiveDirectoryClient> GetGraphClientAsync()
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_graphClient != null)
            {
                return _graphClient;
            }

            // Active Directory service endpoints
            Uri AadServiceEndpointUri = new Uri(AadServiceResourceId);

            try
            {
                //First, look for the authority used during the last authentication.
                //If that value is not populated, use CommonAuthority.
                string authority = String.IsNullOrEmpty(LastAuthority) ? CommonAuthority : LastAuthority;

                // Create an AuthenticationContext using this authority.
                _authenticationContext = new AuthenticationContext(authority);

                var token = await GetTokenHelperAsync(_authenticationContext, AadServiceResourceId);

                // Check the token
                if (string.IsNullOrEmpty(token))
                {
                    // User cancelled sign-in
                    throw new Exception("Sign-in cancelled");  // assuming we don't want to continue
                }
                else
                {
                    // Create our ActiveDirectory client.
                    _graphClient = new ActiveDirectoryClient(
                        new Uri(AadServiceEndpointUri, TenantId),
                        async () => await GetTokenHelperAsync(_authenticationContext, AadServiceResourceId));

                    return _graphClient;
                }
            }
            catch (Exception)
            {
                _authenticationContext.TokenCache.Clear();
                throw;
            }
        }

        /// <summary>
        /// Checks that an OutlookServicesClient object is available. 
        /// </summary>
        /// <returns>The OutlookServicesClient object. </returns>
        public static async Task<OutlookServicesClient> GetOutlookClientAsync(string capability)
        {
#if false

                 using (var client = new HttpClient())
                {
                    string endpointUri, resourceId;

                    using (var discoveryRequest = new HttpRequestMessage(HttpMethod.Get,
                      "https://api.office.com/discovery/v1.0/me/services('Mail@O365_EXCHANGE')"))
                    {
                        discoveryRequest.Headers.Add("Authorization",
                          "Bearer {token:https://api.office.com/discovery/}");

                        using (var discoveryResponse = await client.SendAsync(discoveryRequest))
                        {
                            var discoverContent = await discoveryResponse.Content.ReadAsStringAsync();
                            var serviceInfo = JObject.Parse(discoverContent);
                            endpointUri = serviceInfo["serviceEndpointUri"].ToString();
                            resourceId = serviceInfo["serviceResourceId"].ToString();
                        }
                    }

                    using (var request = new HttpRequestMessage(HttpMethod.Get,
                      endpointUri + "/Me/Folders('Inbox')/Messages?$orderby=DateTimeReceived%20desc"))
                    {
                        request.Headers.Add("Authorization", "Bearer {token:" + resourceId + "}");

                        using (var response = await client.SendAsync(request))
                        {
                            var content = await response.Content.ReadAsStringAsync();
                            foreach (var item in JObject.Parse(content)["value"])
                            {
                                Console.WriteLine("Message \"{0}\" received at \"{1}\"",
                                  item["Subject"],
                                  item["DateTimeReceived"]);
                            }
                        }
                    }
                }
            return null;
#else
            if (_outlookClient != null)
            {
                return _outlookClient;
            }
            
            try
            {
                // Now get the capability that you are interested in.
                var discoveryClient = new DiscoveryClient(DiscoveryServiceEndpointUri,
                                    async () => await GetTokenHelperAsync(_authenticationContext, DiscoveryResourceId));

                CapabilityDiscoveryResult result = await discoveryClient.DiscoverCapabilityAsync(capability);

                _outlookClient = new OutlookServicesClient(
                    result.ServiceEndpointUri,
                    async () => await GetTokenHelperAsync(_authenticationContext, result.ServiceResourceId));

                return _outlookClient;
            }
            catch (Exception ex)
            {
                _authenticationContext.TokenCache.Clear();
                Log.Out(Log.Severity.Error, ModuleName, ex.ToString());
                throw;
            }
#endif
    }

        internal static string GetOutlookToken()
        {
            string resourceId = "https://outlook.office365.com/";
            return GetTokenHelper(_authenticationContext, resourceId);
        }

    // Get an access token for the given context and resourceId. An attempt is first made to 
    // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
    // TODO: Find a way to call context.AcquireTokenAsync directly.
        private static async Task<string> GetTokenHelperAsync(AuthenticationContext context, string resourceId)
        {
            return await Task.Run(() => GetTokenHelper(context, resourceId));
        }

        private static string GetTokenHelper(AuthenticationContext context, string resourceId)
        {
            string accessToken = null;
            
            try
            {
                AuthenticationResult result = context.AcquireToken(resourceId, ClientID, ReturnUri);

                accessToken = result.AccessToken;

                LoggedInUser = result.UserInfo.UniqueId;
                TenantId = result.TenantId;
                LastAuthority = context.Authority;
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Warning, ModuleName, ex.ToString());
            }

            return accessToken;
        }
    }
}
