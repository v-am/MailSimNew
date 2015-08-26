using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office365.OutlookServices;
using MailSim.Common.Contracts;
using Microsoft.OData.ProxyExtensions;
using Newtonsoft.Json;

namespace MailSim.ProvidersSDK
{
    class MailFolderProvider : IMailFolder
    {
        private IFolder _folder;
        private readonly IFolderFetcher _folderFetcher;
        private readonly OutlookServicesClient _outlookClient;
        private readonly bool _isRoot;

        internal MailFolderProvider(OutlookServicesClient outlookClient, string name)
        {
            _outlookClient = outlookClient;
            _folderFetcher = _outlookClient.Me.Folders.GetById(name);

            Util.Synchronize(async () => _folder = await _folderFetcher.ExecuteAsync());
        }

        internal MailFolderProvider(OutlookServicesClient outlookClient, IFolderFetcher folderFetcher, bool isRoot)
        {
            _folderFetcher = folderFetcher;
            _outlookClient = outlookClient;
            _isRoot = isRoot;

            if (isRoot == false)
            {
                Util.Synchronize(async () => _folder = await _folderFetcher.ExecuteAsync());
            }
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
                return Name;    // TODO: could we do that?
            }
        }

        internal string Handle
        {
            get
            {
                return _folder.Id;
            }
        }

        private async Task<long> MailCountRequest(string folderId)
        {
            string uri = string.Format("Folders/{0}/Messages/$count", folderId);

            var json = await Util.GetItemAsync(uri);
            long count = long.Parse(json);

            return count;
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
            long count = 0;
#if true
            Util.Synchronize(async () => count = await MailCountRequest(_folder.Id));
#else
            // TODO: This call fails
            Util.Synchronize(async () => count = await _folderFetcher.Messages.CountAsync());
#endif
            return (int) count;
        }

        public int SubFoldersCount
        {
            get
            {
                IFolder folder = null;
                Util.Synchronize(async () => folder = await _folderFetcher.ExecuteAsync());
                return folder.ChildFolderCount ?? 0;
            }
        }

        public IEnumerable<IMailItem> MailItems
        {
            get
            {
                // Generate a request that gets all mail items in one shot
                return GetMailItems(GetMailCount());
            }
        }

        public void Delete()
        {
            Util.Synchronize(async () => await _folder.DeleteAsync());
        }

        private IEnumerable<IMailItem> GetMailItems()
        {
            IPagedCollection<IMessage> all = null;

            Util.Synchronize(async () => all = await _folderFetcher.Messages
                .ExecuteAsync());

            return PagedToMailItemEnumerable(all);
        }

        private IEnumerable<IMailItem> GetMailItems(int count)
        {
            IPagedCollection<IMessage> all = null;

            Util.Synchronize(async () => all = await _folderFetcher.Messages
                .Take(count)
                .ExecuteAsync());

            return PagedToMailItemEnumerable(all);
        }

        public IMailFolder AddSubFolder(string name)
        {
            Folder newFolder = new Folder
            {
                DisplayName = name
            };

            Util.Synchronize(async () => await _folderFetcher.ChildFolders.AddFolderAsync(newFolder));

            return new MailFolderProvider(_outlookClient, newFolder.Id);
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
            var folderCollection = _isRoot ? _outlookClient.Me.Folders : _folderFetcher.ChildFolders;

            IPagedCollection<IFolder> folders = null;

            Util.Synchronize(async () => folders = await folderCollection.ExecuteAsync());
            
            var allFolders = PagedToEnumerable(folders);

            return allFolders.Select(f => new MailFolderProvider(_outlookClient, f.Id));
        }

        private IEnumerable<IMailItem> PagedToMailItemEnumerable(IPagedCollection<IMessage> pages)
        {
            foreach (var item in pages.CurrentPage)
            {
                yield return new MailItemProvider(_outlookClient, item);
            }

            while (pages.MorePagesAvailable)
            {
                Util.Synchronize(async () => pages = await pages.GetNextPageAsync());

                foreach (var item in pages.CurrentPage)
                {
                    yield return new MailItemProvider(_outlookClient, item);
                }
            }
        }

        // Iterate over all pages to get the whole collection
        private IEnumerable<T> PagedToEnumerable<T>(IPagedCollection<T> pages)
        {
            foreach (var item in pages.CurrentPage)
            {
                yield return item;
            }

            while (pages.MorePagesAvailable)
            {
                Util.Synchronize(async () => pages = await pages.GetNextPageAsync());

                foreach (var item in pages.CurrentPage)
                {
                    yield return item;
                }
            }
        }
 
        public void RegisterItemAddEventHandler(Action<IMailItem> callback)
        {
            // TODO: Implement this
        }

        public void UnRegisterItemAddEventHandler()
        {
            // TODO: Implement this
        }
    }
}
