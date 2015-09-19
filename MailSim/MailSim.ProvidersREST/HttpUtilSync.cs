using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.ProvidersREST
{
    internal class HttpUtilSync
    {
        private readonly string _resourceId;

        internal HttpUtilSync(string resourceId=AuthenticationHelper.OfficeResourceId)
        {
            _resourceId = resourceId;
        }

        private string GetToken(bool isRefresh)
        {
            return AuthenticationHelper.GetToken(_resourceId, isRefresh);
        }

        internal T GetItem<T>(string uri)
        {
            return HttpUtil.GetItemAsync<T>(uri, GetToken).GetResult();
        }

        internal T GetItems<T>(string uri)
        {
            return HttpUtil.GetItemsAsync<T>(uri, GetToken).GetResult();
        }

        internal T PostItem<T>(string uri, T item = default(T))
        {
            return HttpUtil.PostItemAsync<T>(uri, item, GetToken).GetResult();
        }

        internal T PostItemDynamic<T>(string uri, dynamic body)
        {
            // Can't use extensions with dynamic types...
            return HttpUtil.PostItemDynamicAsync<T>(uri, body, new HttpUtil.TokenFunc(GetToken))
                            .ConfigureAwait(false)
                            .GetAwaiter()
                            .GetResult();
        }

        internal void DeleteItem(string uri)
        {
            HttpUtil.DeleteItemAsync(uri, GetToken).GetResult();
        }

        internal T PatchItem<T>(string uri, T item)
        {
            return HttpUtil.PatchItemAsync<T>(uri, item, GetToken).GetResult();
        }

        internal IEnumerable<T> EnumerateCollection<T>(string uri, int count)
        {
#if true
            while (count > 0 && uri != null)
            {
                var msgsColl = HttpUtil.GetCollectionAsync<IEnumerable<T>>(uri, GetToken).GetResult();

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
#else
            return HttpUtil.EnumerateCollection<T>(uri, count, GetToken).GetResult();
#endif
        }
    }
}

