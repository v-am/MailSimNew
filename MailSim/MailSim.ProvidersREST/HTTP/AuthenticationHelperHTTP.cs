//#define USE_UNIFIED

using System;

using MailSim.Common;
using System.Web;
using System.Collections.Generic;


namespace MailSim.ProvidersREST
{
    /// <summary>
    /// Provides clients for the different service endpoints.
    /// </summary>
    internal static class AuthenticationHelperHTTP
    {
#if USE_UNIFIED   // using Unified App registration
        private const string ClientID = "c6de72e6-5aff-4491-97e5-b1b7a419d592";
        private const string TenantId = "702cfb5c-c600-4b2d-962a-21ceb2c260ae";
        private const string SecretKey = "W2RM2RMxM2a1VIP8VT1X/4muQEOL3AnqlZXQiLpSCEg=";
        private const string AadServiceResourceId = "https://graph.microsoft.com/";
#else
        private static readonly string ClientID = Resources.ClientID;
        private static string TenantId { get; set; }
        private const string AadServiceResourceId = "https://graph.windows.net/";
#endif
        // Properties used for communicating with your Windows Azure AD tenant.
        private const string CommonAuthority = "https://login.microsoftonline.com/Common";

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

        private static IDictionary<string, AccessTokenResponse> _tokenResponses = new Dictionary<string, AccessTokenResponse>();

        internal static void Initialize(string userName, string password)
        {
            UserName = userName;
            Password = password;
#if false
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
#endif
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

                string uri = "https://graph.windows.net/myorganization/users?api-version=1.5";
                var zzz = HttpUtil.GetItemsAsync<List<UserHttp>>(uri, GetAADToken).Result;
#endif
            }
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

            return _tokenResponses[resourceId].access_token;
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
            return GetTokenHelper(AadServiceResourceId, isRefresh);
        }

        internal static string GetToken(string resourceId, bool isRefresh)
        {
            return GetTokenHelper(resourceId, isRefresh);
        }

        private static string GetTokenHelper(string resourceId, bool isRefresh)
        {
            return GetTokenHelperHttp(resourceId, isRefresh);
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
