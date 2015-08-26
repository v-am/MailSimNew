using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Newtonsoft.Json.Linq;

namespace MailSim.ProvidersREST
{
    public class MailStoreProviderREST : IMailStore
    {
        private readonly ActiveDirectoryClient _adClient;
        private IMailFolder _rootFolder;

        private static IDictionary<string, string> _predefinedFolders = new Dictionary<string, string>
        {
            {"olFolderInbox",       "Inbox"},
            {"olFolderDeletedItems","Deleted Items"},
            {"olFolderDrafts",      "Drafts"},
            {"olFolderJunk",        "Junk Email"},
            {"olFolderOutbox",      "Outbox"},
            {"olFolderSentMail",    "Sent Items"},
        };

        public MailStoreProviderREST(string mailboxName)
        {
            _adClient = AuthenticationHelper.GetGraphClientAsync().Result;

            var user = Util.GetItemAsync<User>(string.Empty).Result;
            DisplayName = user.Id;
        }

        public IMailItem NewMailItem()
        {
            var body = new MailItemProvider.ItemBody
            {
                Content = "New Body",
                ContentType = "HTML"
            };

            var message = new MailItemProvider.Message
            {
                Subject = "New Subject",
                Body = body,
                ToRecipients = new List<MailItemProvider.Recipient>(),
                Importance = "High"
            };

            // Save the draft message. Saving to Me.Messages saves the message in the Drafts folder.
            var newMessage = Util.PostItemAsync("Messages", message).Result;

            return new MailItemProvider(newMessage);
        }

        public string DisplayName { get; private set; }

        public IMailFolder RootFolder
        {
            get
            {
                if (_rootFolder == null)
                {
                    _rootFolder = new MailFolderProvider(null);
                }
                
                return _rootFolder;
            }
        }

        public IMailFolder GetDefaultFolder(string name)
        {
            string folderName;

            if (_predefinedFolders.TryGetValue(name, out folderName) == false)
            {
                return null;
            }

            return RootFolder.SubFolders.FirstOrDefault(x => x.Name == folderName);
        }

        public IAddressBook GetGlobalAddressList()
        {
            return new AddressBookProvider(_adClient);
        }

        private class User
        {
            public string Id { get; set; }
        }
    }
}
