using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MailSim.Common.Contracts;
using MailSim.ProvidersOM;
//using MailSim.ProvidersSDK;

namespace MailSim
{
    class ProviderFactory
    {
        public static IMailStore CreateMailStore(string mailboxName)
        {
            if (false)
            {
                return new MailStoreProviderOM(mailboxName);
            }
            else
            {
//                return new ProvidersSDK.MailStoreProvider(mailboxName);
                return new ProvidersREST.MailStoreProviderREST(mailboxName);
            }
        }
    }
}
