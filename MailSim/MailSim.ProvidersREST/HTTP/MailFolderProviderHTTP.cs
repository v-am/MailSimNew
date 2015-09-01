using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using System.Dynamic;
using System.Net;
using MailSim.Common;

namespace MailSim.ProvidersREST
{
    class MailFolderProviderHTTP : IMailFolder
    {
        private readonly Folder _folder;
//        private string _subscriptionId;

        internal MailFolderProviderHTTP(Folder folder, string name = null)
        {
            _folder = folder;

            if (name != null)
            {
                Name = name;
            }
            else if (_folder != null)
            {
                Name = _folder.DisplayName;
            }
        }

        public string Name { get; private set; }

        public string FolderPath
        {
            get
            {
                return Name;    // TODO: is it the right thing to do?
            }
        }

        internal string Handle
        {
            get
            {
                return _folder.Id;
            }
        }
 
        public int MailItemsCount
        {
            get
            {
                return GetMailCount();
            }
        }

        private int GetMailCount()
        {
            return HttpUtil.GetItemAsync<int>(Uri + "/Messages/$count").Result;
        }

        public int SubFoldersCount
        {
            get
            {
                return GetChildFolderCount();
            }
        }

        private int GetChildFolderCount()
        {
            if (_folder == null)
            {
                return HttpUtil.GetItemAsync<int>("Folders/$count").Result;
            }
            else
            {
                return HttpUtil.GetItemAsync<int>(Uri + "/ChildFolders/$count").Result;
            }
        }

        public IEnumerable<IMailItem> MailItems
        {
            get
            {
                return GetMailItems(string.Empty, GetMailCount());
            }
        }

        private string Uri
        {
            get
            {
                return string.Format("Folders/{0}", _folder.Id);
            }
        }

        public void Delete()
        {
            HttpUtil.DeleteAsync(Uri).Wait();
        }

        public IEnumerable<IMailItem> GetMailItems(string filter, int count)
        {
            var msgs = GetMessages(filter, count);

            var items = msgs.Select(x => new MailItemProviderHTTP(x));

            filter = filter ?? string.Empty;

            return items.Where(i => i.Subject.ContainsCaseInsensitive(filter));
        }

        private IEnumerable<MailItemProviderHTTP.Message> GetMessages(string filter, int count)
        {
            string uri = Uri + string.Format("/Messages?$search=\"{1}\"&$top={0}", count, filter);

            return HttpUtil.EnumerateCollection<MailItemProviderHTTP.Message>(uri, count);
        }

        public IMailFolder AddSubFolder(string name)
        {
            dynamic folderName = new ExpandoObject();
            folderName.DisplayName = name;

            Folder newFolder = HttpUtil.PostDynamicAsync<Folder>(Uri + "/ChildFolders", folderName).Result;

            return new MailFolderProviderHTTP(newFolder);
        }

        public IEnumerable<IMailFolder> SubFolders
        {
            get
            {
                return GetSubFolders();
            }
        }

        private IEnumerable<IMailFolder> GetSubFolders()
        {
            string uri = _folder == null ? "Folders" : Uri + "/ChildFolders";

            var folders = HttpUtil.EnumerateCollection<Folder>(uri, int.MaxValue);

            return folders.Select(f => new MailFolderProviderHTTP(f));
        }

        public void RegisterItemAddEventHandler(Action<IMailItem> callback)
        {
#if false
            string baseUri = "https://outlook.office.com/api/beta/me";

            string uri = baseUri + "/subscriptions";

            var res = Util.DoHttp<SubscriptionRequest, SubscriptionResponse>("POST", uri, new SubscriptionRequest()
            {
                ResourceURL = string.Format("{0}/{1}/messages", baseUri, Uri),
                Type = "#Microsoft.OutlookServices.PushSubscription",
                CallbackURL = "https://webhook.azurewebsites.net/api/send/myNotifyClient",
                ChangeType = "Created",
                ClientState = "3250be24-1282-4b46-a41e-0e53b4cae73f"    // GUID
            }).Result;

            _subscriptionId = res.Id;

            StartNotificationListener(res.Id, callback);
#endif
        }

        private static void StartNotificationListener(string id, Action<IMailItem> callback)
        {
        }

        public void UnRegisterItemAddEventHandler()
        {
#if false
            string baseUri = "https://outlook.office.com/api/beta/me";

            string uri = string.Format("{0}/subscriptions('{1}')", baseUri, _subscriptionId);

            Util.DeleteAsync(uri).Wait();
#endif
        }

        internal class Folder
        {
            public string Id { get; set; }
            public string DisplayName { get; set; }
            public int ChildFolderCount { get; set; }
        }

        private class SubscriptionResponse
        {
            public string Id { get; set; }
            public string ChangeType { get; set; }
            public DateTime ExpirationTime { get; set; }
        }

        private class SubscriptionRequest
        {
            [Newtonsoft.Json.JsonProperty("@odata.type")]
            public string Type { get; set; }
            public string ResourceURL { get; set; }
            public string CallbackURL { get; set; }
            public string ChangeType { get; set; }
            public string ClientState { get; set; }
        }
    }
}
