//#define USE_UNIFIED

using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Threading.Tasks;

using MailSim.Common;
using System.Web;
using System.Collections.Generic;


namespace MailSim.ProvidersREST
{
    /// <summary>
    /// Provides clients for the different service endpoints.
    /// </summary>
    internal static class AuthenticationHelper
    {
#if USE_UNIFIED   // using Unified App registration
        private static readonly string ClientID = "c6de72e6-5aff-4491-97e5-b1b7a419d592";
        private static string TenantId = "702cfb5c-c600-4b2d-962a-21ceb2c260ae";
        private static string SecretKey = "W2RM2RMxM2a1VIP8VT1X/4muQEOL3AnqlZXQiLpSCEg=";
        private const string AadServiceResourceId = "https://graph.microsoft.com/";
#else
        private static readonly string ClientID = Resources.ClientID;
        private static string TenantId { get; set; }
        private const string AadServiceResourceId = "https://graph.windows.net/";
#endif
        private static readonly Uri ReturnUri = new Uri(Resources.ReturnUri);

        // Properties used for communicating with your Windows Azure AD tenant.
        private const string CommonAuthority = "https://login.microsoftonline.com/Common";
        internal const string OfficeResourceId = "https://outlook.office365.com/";

        private const string ModuleName = "AuthenticationHelper";

        //Static variables store the clients so that we don't have to create them more than once.
        private static ActiveDirectoryClient _graphClient = null;

        //Property for storing the authentication context.
        private static AuthenticationContext _authenticationContext { get; set; }

        //Property for storing and returning the authority used by the last authentication.
        private static string LastAuthority { get; set; }
        //Property for storing the tenant id so that we can pass it to the ActiveDirectoryClient constructor.
        // Property for storing the logged-in user so that we can display user properties later.
        internal static string LoggedInUser { get; set; }

        private static string UserName { get; set; }
        private static string Password { get; set; }

        private class AuthResponse
        {
            public bool admin_consent { get; set; }
            public string code { get; set; }
            public string session_state { get; set; }
            public string state { get; set; }
        }

        private class AccessTokenResponse
        {
            public string access_token { get; set; }
            public int expires_in { get; set; }
            public int expires_on { get; set; }
            public string id_token { get; set; }
            public string refresh_token { get; set; }
            public string resource { get; set; }
            public string scope { get; set; }
            public string token_type { get; set; }
        }

        private class IdToken
        {
            public string tid { get; set; }
        }

        private static bool _useHttp = true;
        private static IDictionary<string, AccessTokenResponse> _tokenResponses = new Dictionary<string, AccessTokenResponse>();

        /// <summary>
        /// Checks that a Graph client is available.
        /// </summary>
        /// <returns>The Graph client.</returns>
        internal static async Task<ActiveDirectoryClient> GetGraphClientAsync(string userName, string password)
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_graphClient != null)
            {
                return _graphClient;
            }

            UserName = userName;
            Password = password;

            // Active Directory service endpoints
            Uri aadServiceEndpointUri = new Uri(AadServiceResourceId);

            if (_useHttp)
            {
#if false
                string uri = string.Format("{0}/authorize/?response_type=code&client_id={1}&redirect_uri={2}",
                                                            oauthUri,
                                                            Resources.ClientID,
                                                            Resources.ReturnUri);
                var res = HttpUtilSync.GetItem<AuthResponse>(uri);
#endif

                var token = GetTokenHelperHttp(AadServiceResourceId, false);
                var tokenResponse = _tokenResponses[AadServiceResourceId];

                var id_token = tokenResponse.id_token;

                IdToken idToken = null;

                try
                {
                    //                    var json = JWT.JsonWebToken.Decode(id_token, string.Empty, verify:false);
                    idToken = JWT.JsonWebToken.DecodeToObject<IdToken>(id_token, string.Empty, verify: false);
                }
                catch (Exception ex)
                {
                }

                // Check the token
                if (string.IsNullOrEmpty(token))
                {
                    // User cancelled sign-in
                    throw new Exception("Sign-in cancelled");  // assuming we don't want to continue
                }
                else
                {
#if USE_UNIFIED
                    string uri = "https://graph.microsoft.com/beta/me";
                    var xxx = HttpUtil.GetItemAsync<UserHttp>(uri, GetAADToken).Result;

                    uri = "https://graph.microsoft.com/beta/" + idToken.tid + "/users/" + "admin@mailsimdemo.onmicrosoft.com";
                    var yyy = HttpUtil.GetItemAsync<UserHttp>(uri, GetAADToken).Result;

                    uri = "https://graph.microsoft.com/beta/" + idToken.tid + "/users/";
                    var zzz = HttpUtil.GetItemsAsync<List<UserHttp>>(uri, GetAADToken).Result;
#else
                    // Create our ActiveDirectory client.
                    string authority = String.IsNullOrEmpty(LastAuthority) ? CommonAuthority : LastAuthority;
#if true
//                    string uri = "https://graph.windows.net/mailsimdemo.onmicrosoft.com/users?api-version=beta";
                    string uri = "https://graph.windows.net/myorganization/users?api-version=beta";
                    var zzz = HttpUtil.GetItemsAsync<List<UserHttp>>(uri, GetAADToken).Result;
#else
                    // Create an AuthenticationContext using this authority.
                    var authenticationContext = new AuthenticationContext(authority);

                    // Get TenandId
                    var token2 = await GetTokenHelperAsync(authenticationContext, AadServiceResourceId);

                    _graphClient = new ActiveDirectoryClient(
                            new Uri(aadServiceEndpointUri, TenantId),
                            async () => await GetTokenHelperAsync(authenticationContext, AadServiceResourceId));
#endif
#endif
                    return _graphClient;
                }
            }
            else
            {
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
                            new Uri(aadServiceEndpointUri, TenantId),
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
        }

        private static async Task<string> GetTokenHelperAsync2(string resourceId)
        {
            return await Task.Run(() => GetTokenHelperHttp(resourceId, false));
        }

        private static string GetTokenHelperHttp(string resourceId, bool isRefresh)
        {
            AccessTokenResponse tokenResponse;

            if (_tokenResponses.TryGetValue(resourceId, out tokenResponse) == false)
            {
                _tokenResponses[resourceId] = QueryTokenResponse(resourceId);
            }
            else if (isRefresh)
            {
                string uri = CommonAuthority + "/oauth2/" + "token";
                var authResponse = _tokenResponses[resourceId];

                string body = string.Format("grant_type=refresh_token&refresh_token={0}&client_id={1}&resource={2}",
                                                HttpUtility.UrlEncode(authResponse.refresh_token),
                                                HttpUtility.UrlEncode(ClientID),
                                                HttpUtility.UrlEncode(resourceId)
                                                );

                Log.Out(Log.Severity.Info, "", "Sending request for new token:" + body);

                var newAuthResponse = HttpUtil.DoHttp<string, AccessTokenResponse>("POST", uri, body, (dummy) => null).Result;
                _tokenResponses[resourceId] = newAuthResponse;

                Log.Out(Log.Severity.Info, "", "Got new access token:" + _tokenResponses[resourceId].access_token);
            }

            string accessToken = _tokenResponses[resourceId].access_token;

            return accessToken;
        }

        private static AccessTokenResponse QueryTokenResponse(string resourceId)
        {
#if USE_UNIFIED
            string oauthUri = CommonAuthority + "/oauth2/";

            string uri = string.Format("{0}token", oauthUri);

            string body = string.Format("resource={0}&client_id={1}&grant_type=password&username={2}&password={3}&client_secret={4}&scope=openid",
                                            HttpUtility.UrlEncode(resourceId),
                                            HttpUtility.UrlEncode(ClientID),
                                            HttpUtility.UrlEncode(UserName),
                                            HttpUtility.UrlEncode(Password),
                                            HttpUtility.UrlEncode(SecretKey)
                                            );
#else
            string oauthUri = CommonAuthority + "/oauth2/";

            string uri = string.Format("{0}token", oauthUri);

            string body = string.Format("resource={0}&client_id={1}&grant_type=password&username={2}&password={3}&scope=openid",
                                            HttpUtility.UrlEncode(resourceId),
                                            HttpUtility.UrlEncode(ClientID),
                                            HttpUtility.UrlEncode(UserName),
                                            HttpUtility.UrlEncode(Password));
#endif
            return HttpUtil.DoHttp<string, AccessTokenResponse>("POST", uri, body, (isRefresh) => null).Result;
        }

        internal static string GetAADToken(bool isRefresh)
        {
            return GetTokenHelper(_authenticationContext, AadServiceResourceId, isRefresh);
        }

        internal static string GetToken(string resourceId, bool isRefresh)
        {
            return GetTokenHelper(_authenticationContext, resourceId, isRefresh);
        }

        internal static async Task<string> GetTokenAsync(string resourceId)
        {
            return await GetTokenHelperAsync(_authenticationContext, resourceId);
        }

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        // TODO: Find a way to call context.AcquireTokenAsync directly.
        private static async Task<string> GetTokenHelperAsync(AuthenticationContext context, string resourceId)
        {
            return await Task.Run(() => GetTokenHelper(context, resourceId, false));
        }

        private static string GetTokenHelper(AuthenticationContext context, string resourceId, bool isRefresh)
        {
            string accessToken = null;

            if (context == null)
            {
                return GetTokenHelperHttp(resourceId, isRefresh);
            }
            
            try
            {
                AuthenticationResult result;

                if (string.IsNullOrEmpty(UserName) || string.IsNullOrEmpty(Password))
                {
                    result = context.AcquireToken(resourceId, ClientID, ReturnUri);
                }
                else
                {
                    result = context.AcquireToken(resourceId, ClientID, new UserCredential(UserName, Password));
                }

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

        private class UserHttp
        {
            public string UserPrincipalName { get; set; }
            public string DisplayName { get; set; }
            public string GivenName { get; set; }
            public string SurName { get; set; }
        }
    }
}
