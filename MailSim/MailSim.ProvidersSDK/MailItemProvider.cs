using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using System.IO;

namespace MailSim.ProvidersSDK
{
    class MailItemProvider : IMailItem
    {
        private readonly OutlookServicesClient _outlookClient;
        private readonly IMessage _message;

        public MailItemProvider(OutlookServicesClient outlookClient, IMessage msg)
        {
            _outlookClient = outlookClient;
           _message = msg;
        }

        public string Subject
        {
            get
            {
                return _message.Subject;
            }

            set
            {
                SetAndUpdate((message) => message.Subject = value);
            }
        }

        public string Body
        {
            get
            {
                return _message.Body.Content;
            }

            set
            {
                SetAndUpdate((message) =>
                    message.Body = new ItemBody
                    {
                        Content = value,
                        ContentType = BodyType.HTML
                    });
            }
        }
 
        public string SenderName
        {
            get
            {
                return _message.Sender.EmailAddress.Address;
            }
        }

        private void SetAndUpdate(Action<IMessage> action)
        {
            IMessage message = null;
            
            Util.Synchronize(async() => message = await _outlookClient.Me.Messages[_message.Id].ExecuteAsync());

            action(message);
            
            Util.Synchronize(async() => await message.UpdateAsync());
        }

        public void AddRecipient(string recipient)
        {
            SetAndUpdate((message) => message.ToRecipients.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }));
        }

        public void AddAttachment(string filepath)
        {
            using (var reader = new StreamReader(filepath))
            {
                var contents = reader.ReadToEnd();

                var msgFetcher = _outlookClient.Me.Messages.GetById(_message.Id);

                var bytes = System.Text.Encoding.Unicode.GetBytes(contents);
                var name = filepath.Split('\\').Last();

                var fileAttachment = new FileAttachment
                {
                    ContentBytes = bytes,
                    Name = name,
                    Size = bytes.Length
                };

                Util.Synchronize(async() => await msgFetcher.Attachments.AddAttachmentAsync(fileAttachment));
            }
        }

        public void AddAttachment(IMailItem mailItem)
        {
#if false   // TODO: This doesn't work
            var itemProvider = mailItem as MailItemProvider;

            var msgFetcher = _outlookClient.Me.Messages.GetById(_message.Id);

            var itemAttachment = new ItemAttachment
            {
                Item = itemProvider.Handle as Message,

                Name = "Item Attachment!!!",
            };

           Util.Synchronize(async () => await msgFetcher.Attachments.AddAttachmentAsync(itemAttachment));
#endif
        }

        internal Item Handle
        {
            get
            {
                return _message as Item;
            }
        }
        
        private IMessageFetcher Message
        {
            get
            {
                return _outlookClient.Me.Messages[_message.Id];
            }
        }

        // Create a reply message
        public IMailItem Reply(bool replyAll)
        {
            IMessage replyMsg = null;

            if (replyAll)
            {
                Util.Synchronize(async () => replyMsg = await Message.CreateReplyAllAsync());
            }
            else
            {
                Util.Synchronize(async () => replyMsg = await Message.CreateReplyAsync());
            }

            return new MailItemProvider(_outlookClient, replyMsg);
        }

        public IMailItem Forward()
        {
            IMessage msg = null;

            Util.Synchronize(async () => msg = await Message.CreateForwardAsync());

            return new MailItemProvider(_outlookClient, msg);
        }

        public void Send()
        {
            // This generates Me/SendMail; all the data should be in the body
            Util.Synchronize(async () => await Message.SendAsync());
        }
        
        // TODO: Should this method return a IMailItem?
        public void Move(IMailFolder newFolder)
        {
            MailFolderProvider folderProvider = newFolder as MailFolderProvider;

            var folderId = folderProvider.Handle;
            Util.Synchronize(async () => await Message.MoveAsync(folderId));
        }

        public void Delete()
        {
            Util.Synchronize(async () => await _message.DeleteAsync());
        }

        public bool ValidateRecipients()
        {
            // TODO: Implement this
            return true;
        }
    }
}
