using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Office365.OutlookServices;
using Newtonsoft.Json.Linq;

namespace MailSim.ProvidersSDK
{
    public class MailStoreProvider : IMailStore
    {
        private readonly ActiveDirectoryClient _adClient;
        private readonly OutlookServicesClient _outlookClient;
        private readonly Microsoft.Office365.OutlookServices.IUser _user;
        private IMailFolder _rootFolder;

        private static IDictionary<string, string> _predefinedFolders = new Dictionary<string, string>
        {
            {"olFolderInbox", "Inbox"},
            {"olFolderDeletedItems", "Deleted Items"},
            {"olFolderDrafts", "Drafts"},
            {"olFolderJunk", "Junk Email"},
            {"olFolderOutbox", "Outbox"},
            {"olFolderSentMail", "Sent Items"},
        };

        public MailStoreProvider(string mailboxName)
        {
            _adClient = AuthenticationHelper.GetGraphClientAsync().Result;

            _outlookClient = AuthenticationHelper.GetOutlookClientAsync("Mail").Result;

            _user = _outlookClient.Me.ExecuteAsync().Result;

            var json = Util.GetItemAsync("").Result;

            DisplayName = JObject.Parse(json)["Id"].ToString();
        }

        public IMailItem NewMailItem()
        {
            ItemBody body = new ItemBody
            {
                Content = "New Body",
                ContentType = BodyType.HTML
            };

            Message message = new Message
            {
                Subject = "New Subject",
                Body = body,
                ToRecipients = new List<Recipient>(),
                Importance = Importance.High
            };

            // Save the draft message. Saving to Me.Messages saves the message in the Drafts folder.
            Util.Synchronize(async () => await _outlookClient.Me.Messages.AddMessageAsync(message));

            return new MailItemProvider(_outlookClient, message);
        }

        public string DisplayName { get; private set; }

        public IMailFolder RootFolder
        {
            get
            {
                if (_rootFolder == null)
                {
                    _rootFolder = new MailFolderProvider(_outlookClient, _outlookClient.Me.RootFolder, isRoot:true);
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
    }
}
