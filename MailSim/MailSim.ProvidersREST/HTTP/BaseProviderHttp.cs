using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.ProvidersREST.HTTP
{
    internal class BaseProviderHttp
    {
        static HttpUtilSync _httpUtilSync = new HttpUtilSync(AuthenticationHelperHTTP.OfficeResourceId);

        internal HttpUtilSync HttpUtilSync
        {
            get
            {
                return _httpUtilSync;
            }
        }
    }
}
