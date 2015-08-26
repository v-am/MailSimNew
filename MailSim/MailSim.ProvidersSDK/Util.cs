using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using MailSim.Common;
using System.Net.Http;
using System.Net.Http.Headers;

namespace MailSim.ProvidersSDK
{
    class Util
    {
        internal static void Synchronize(Func<Task> func)
        {
            var task = func();

            try
            {
                task.Wait();
            }
            catch (AggregateException)
            {
                throw;
            }
        }

        private const string baseUri = @"https://outlook.office365.com/api/v1.0/Me/";

        internal static async Task<string> GetItemAsync(string subUri)
        {
            HttpContent content;
            string uri = baseUri + subUri;

            using (HttpClient client = new HttpClient())
            {
                string token = AuthenticationHelper.GetOutlookToken();

                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token);

                HttpResponseMessage response = await client.GetAsync(uri);
                response.EnsureSuccessStatusCode();
                content = response.Content;
            }

            return await content.ReadAsStringAsync();
        }
    }
}
