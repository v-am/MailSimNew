﻿using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Net.Http.Formatting;
using System.Dynamic;
using System.Linq;
using MailSim.Common;
using System.Text;

namespace MailSim.ProvidersREST
{
    static class HttpUtil
    {
        private const string baseOutlookUri = @"https://outlook.office.com/api/v1.0/Me/";

        internal static async Task<T> GetItemAsync<T>(string uri)
        {
            return await DoHttp<EmptyBody,T>(HttpMethod.Get, uri, null);
        }

        internal static async Task<T> GetItemsAsync<T>(string uri)
        {
            var coll = await GetCollectionAsync<T>(uri);

            return coll.value;
        }

        internal static IEnumerable<T> EnumerateCollection<T>(string uri, int count)
        {
            while (count > 0 && uri != null)
            {
                var msgsColl = GetCollectionAsync<IEnumerable<T>>(uri).Result;

                foreach (var m in msgsColl.value)
                {
                    if (--count < 0)
                    {
                        yield break;
                    }
                    yield return m;
                }

                uri = msgsColl.NextLink;
            }
        }

        private static async Task<ODataCollection<T>> GetCollectionAsync<T>(string uri)
        {
            return await DoHttp<EmptyBody, ODataCollection<T>>(HttpMethod.Get, uri, null);
        }

        internal static async Task<T> PostItemAsync<T>(string uri, T item=default(T))
        {
            return await DoHttp<T, T>(HttpMethod.Post, uri, item);
        }

        internal static async Task<T> PostItemDynamicAsync<T>(string uri, dynamic body)
        {
            return await DoHttp<ExpandoObject, T>(HttpMethod.Post, uri, body);
        }

        internal static async Task DeleteItemAsync(string uri)
        {
            using (HttpClient client = GetHttpClient())
            {
                var response = await client.DeleteAsync(BuildUri(uri)).ConfigureAwait(false);

                response.EnsureSuccessStatusCode();
            }
        }

        internal static async Task<T> PatchItemAsync<T>(string uri, T item)
        {
            return await DoHttp<T,T>("PATCH", uri, item);
        }
#if false
        internal static async Task<TResult> DoHttp2<TBody, TResult>(string methodName, string uri, string body)
        {
            return await DoHttp2<TBody, TResult>(new HttpMethod(methodName), uri, body);

        }
        private static HttpClient GetHttpClient2()
        {
            HttpClient client = new HttpClient();

            return client;
        }

        private static async Task<TResult> DoHttp2<TBody, TResult>(HttpMethod method, string uri, string body)
        {
            Log.Out(Log.Severity.Info, "DoHttp", string.Format("Uri=[{0}]", uri));

            HttpResponseMessage response;
            var request = new HttpRequestMessage(method, BuildUri(uri));

//            request.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            if (body != null)
            {
                request.Content = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded");
            }

            using (HttpClient client = GetHttpClient2())
            {
                response = await client.SendAsync(request).ConfigureAwait(false);
            }

            string jsonResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            Log.Out(Log.Severity.Info, "DoHttp", "Got response!");

            if (response.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<TResult>(jsonResponse);
            }
            else
            {
                var errorDetail = JsonConvert.DeserializeObject<ODataError>(jsonResponse);
                throw new System.Exception(errorDetail.error.message);
            }
        }
#endif
        private static async Task<TResult> DoHttp<TBody, TResult>(HttpMethod method, string uri, TBody body)
        {
            Log.Out(Log.Severity.Info, "DoHttp", string.Format("Uri=[{0}]", uri));

            var request = new HttpRequestMessage(method, BuildUri(uri));

            if (body != null)
            {
                if (body is string)
                {
                    request.Content = new StringContent(body as string, Encoding.UTF8, "application/x-www-form-urlencoded");
                }
                else
                {
                    request.Content = new ObjectContent<TBody>(body, new JsonMediaTypeFormatter());
                }
            }

            string token = AuthenticationHelper.GetOutlookToken();

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            HttpResponseMessage response;

            using (HttpClient client = GetHttpClient())
            {
                response = await client.SendAsync(request).ConfigureAwait(false);
            }

            string jsonResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            Log.Out(Log.Severity.Info, "DoHttp", "Got response!");

            if (response.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<TResult>(jsonResponse);
            }
            else
            {
                var errorDetail = JsonConvert.DeserializeObject<ODataError>(jsonResponse);
                throw new System.Exception(errorDetail.error.message);
            }
        }

        private class ODataError
        {
            public class Error
            {
                public string code { get; set; }
                public string message { get; set; }
            }
            public Error error { get; set; }
        }

        internal static async Task<TResult> DoHttp<TBody, TResult>(string methodName, string uri, TBody body)
        {
            return await DoHttp<TBody,TResult>(new HttpMethod(methodName), uri, body);
        }

        private static string BuildUri(string subUri)
        {
            if (subUri.StartsWith("http"))
            {
                return subUri;
            }

            return baseOutlookUri + subUri;
        }

        private static HttpClient GetHttpClient()
        {
            HttpClient client = new HttpClient();
#if false
            string token = AuthenticationHelper.GetOutlookToken();

            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);
#endif
            return client;
        }

        internal class ODataCollection<TCollection>
        {
            [Newtonsoft.Json.JsonProperty("@odata.nextLink")]
            public string NextLink { get; set; }

            public TCollection value { get; set; }
        }

        private class EmptyBody { }
    }
}