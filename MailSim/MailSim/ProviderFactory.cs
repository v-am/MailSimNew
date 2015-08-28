using MailSim.Common.Contracts;

namespace MailSim
{
    class ProviderFactory
    {
        public static IMailStore CreateMailStore(string mailboxName, MailSimSequence seq = null)
        {
            if (false)
            {
                return new ProvidersOM.MailStoreProviderOM(mailboxName, seq == null ? false : seq.DisableOutlookPrompt);
            }
            else
            {
//                return new ProvidersREST.MailStoreProviderSDK();
                return new ProvidersREST.MailStoreProviderHTTP();
            }
        }
    }
}
