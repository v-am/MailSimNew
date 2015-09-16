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
        private static readonly string ClientID = Resources.ClientID;

        private static readonly Uri ReturnUri = new Uri(Resources.ReturnUri);

        // Properties used for communicating with your Windows Azure AD tenant.
        private const string CommonAuthority = "https://login.microsoftonline.com/Common";
        private const string AadServiceResourceId = "https://graph.windows.net/";
        private const string OfficeResourceId = "https://outlook.office365.com/";

        private const string ModuleName = "AuthenticationHelper";

        //Static variables store the clients so that we don't have to create them more than once.
        private static ActiveDirectoryClient _graphClient = null;

        //Property for storing the authentication context.
        private static AuthenticationContext _authenticationContext { get; set; }

        //Property for storing and returning the authority used by the last authentication.
        private static string LastAuthority { get; set; }
        //Property for storing the tenant id so that we can pass it to the ActiveDirectoryClient constructor.
        private static string TenantId { get; set; }
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

        private static bool _useHttp = false;
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

                //                string oauthUri = CommonAuthority + "/oauth2/";
#if false
                string uri = string.Format("{0}/authorize/?response_type=code&client_id={1}&redirect_uri={2}",
                                                            oauthUri,
                                                            Resources.ClientID,
                                                            Resources.ReturnUri);
                var res = HttpUtilSync.GetItem<AuthResponse>(uri);
#endif
                var token = GetTokenHelper2(AadServiceResourceId);

                var tokenResponse = _tokenResponses[AadServiceResourceId];
                var id_token = tokenResponse.id_token;

                //                var xxx = JWT.JsonWebToken.Base64UrlDecode(id_token);
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
                    // Create our ActiveDirectory client.
                    // TODO: Is TenantID required?
                    _graphClient = new ActiveDirectoryClient(
                        new Uri(aadServiceEndpointUri, idToken.tid),
                        async () => await GetTokenHelperAsync2(AadServiceResourceId));

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
            return await Task.Run(() => GetTokenHelper2(resourceId));
        }

        private static string GetTokenHelper2(string resourceId)
        {
            AccessTokenResponse tokenResponse;

            if (_tokenResponses.TryGetValue(resourceId, out tokenResponse) == false)
            {
                _tokenResponses[resourceId] = QueryTokenResponse(resourceId);
            }

            string accessToken = _tokenResponses[resourceId].access_token;

            return accessToken;
        }

        private static AccessTokenResponse QueryTokenResponse(string resourceId)
        {
            string oauthUri = CommonAuthority + "/oauth2/";

            string uri = string.Format("{0}token", oauthUri);

            string body = string.Format("resource={0}&client_id={1}&grant_type=password&username={2}&password={3}&scope=openid",
                                            HttpUtility.UrlEncode(resourceId),
                                            HttpUtility.UrlEncode(ClientID),
                                            HttpUtility.UrlEncode(UserName),
                                            HttpUtility.UrlEncode(Password));

            return HttpUtil.DoHttp2<string, AccessTokenResponse>("POST", uri, body).Result;
        }

        internal static string GetOutlookToken2()
        {
            return GetTokenHelper2(OfficeResourceId);
        }

        internal static string GetOutlookToken()
        {
            return GetTokenHelper(_authenticationContext, OfficeResourceId);
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
            return await Task.Run(() => GetTokenHelper(context, resourceId));
        }

        private static string GetTokenHelper(AuthenticationContext context, string resourceId)
        {
            string accessToken = null;
            
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
    }
}
