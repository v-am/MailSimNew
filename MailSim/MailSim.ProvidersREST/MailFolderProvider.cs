using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using System.Dynamic;
using System.Net;

namespace MailSim.ProvidersREST
{
    class MailFolderProvider : IMailFolder
    {
        private readonly Folder _folder;
//        private string _subscriptionId;

        internal MailFolderProvider(Folder folder)
        {
            _folder = folder;
        }

        public string Name
        {
            get
            {
                return _folder.DisplayName;
            }
        }

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
            return Util.GetItemAsync<int>(Uri + "/Messages/$count").Result;
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
            return Util.GetItemAsync<int>(Uri + "/ChildFolders/$count").Result;
        }

        public IEnumerable<IMailItem> MailItems
        {
            get
            {
                return GetMailItems(GetMailCount());
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
            Util.DeleteAsync(Uri).Wait();
        }

        private IEnumerable<IMailItem> GetMailItems(int count)
        {
            var msgs = GetMessages(count);

            return msgs.Select(x => new MailItemProvider(x));
        }

        private IEnumerable<MailItemProvider.Message> GetMessages(int count)
        {
            int pageSize = 0;

            while (count > 0)
            {
                var uri = Uri + string.Format("/Messages?$skip={1}&$top={0}", count, pageSize);
                var msgs = Util.GetItemsAsync<IEnumerable<MailItemProvider.Message>>(uri).Result;

                pageSize = msgs.Count();

                foreach (var m in msgs)
                {
                    yield return m;
                }

                count -= pageSize;
            }
        }

        public IMailFolder AddSubFolder(string name)
        {
            dynamic folderName = new ExpandoObject();
            folderName.DisplayName = name;

            Folder newFolder = Util.PostDynamicAsync<Folder>(Uri + "/ChildFolders", folderName).Result;

            return new MailFolderProvider(newFolder);
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

            IEnumerable<Folder> folders = Util.GetItemsAsync<List<Folder>>(uri).Result;

            return folders.Select(f => new MailFolderProvider(f));
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
