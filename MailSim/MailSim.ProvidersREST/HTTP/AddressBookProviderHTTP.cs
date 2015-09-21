using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using MailSim.Common.Contracts;
using MailSim.Common;

namespace MailSim.ProvidersREST
{
    class AddressBookProviderHTTP : IAddressBook
    {
        private const string BaseUri = "https://graph.windows.net/myorganization/";
        private const string ApiVersion = "api-version=1.5";

        public IEnumerable<string> GetUsers(string match, int count)
        {
            return EnumerateUsers(match, count);
        }

        public IEnumerable<string> GetDLMembers(string dLName, int count)
        {

            if (string.IsNullOrEmpty(dLName))
            {
                return Enumerable.Empty<string>();
            }

            string uri = BaseUri + "groups";
            uri += '?';     // we always have at least api version parameter

            uri = AddFilters(uri, dLName,
                            "displayName"
                            );

            uri += '&';
            uri += ApiVersion;

            var httpProxy = new HttpUtilSync(Constants.AadServiceResourceId);

            var groups = httpProxy.EnumerateCollection<GroupHttp>(uri, 100);

            // Look for the group with exact name match
            var group = groups.FirstOrDefault((g) => g.DisplayName.EqualsCaseInsensitive(dLName));

            if (group == null)
            {
                return Enumerable.Empty<string>();
            }

            uri = BaseUri + "groups/" + group.ObjectId + "/members?" + ApiVersion;

            var members = httpProxy.EnumerateCollection<UserHttp>(uri, count);

            return members.Select(x => x.UserPrincipalName);
        }

        private IEnumerable<string> EnumerateUsers(string match, int count)
        {
            string uri = BaseUri + "users";
            uri += '?';     // we always have at least api version parameter

            if (string.IsNullOrEmpty(match) == false)
            {
                uri = AddFilters(uri, match,
                            "userPrincipalName",
                            "displayName",
                            "givenName"/*, "surName"*/);

                uri += '&';
            }

            uri += ApiVersion;

            var users = new HttpUtilSync(Constants.AadServiceResourceId)
                    .EnumerateCollection<UserHttp>(uri, count);

            return users.Select(x => x.UserPrincipalName);
        }

        private static string AddFilters(string uri, string match, params string[] fields)
        {
            var sb = new StringBuilder(uri);

            sb.Append("$filter=");

            for (int i = 0; i < fields.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append("%20or%20");
                }

                sb.AppendFormat("startswith({0}, '{1}')", fields[i], match);
            }

            return sb.ToString();
        }
#if false

        private IEnumerable<T> GetFilteredItems<T>(IPagedCollection<T> pages, int count, Func<T, bool> filter)
        {
            foreach (var item in pages.CurrentPage)
            {
                if (--count < 0)
                {
                    yield break;
                }
                else if (filter(item))
                {
                    yield return item;
                }
            }

            while (count > 0 && pages.MorePagesAvailable)
            {
                pages = pages.GetNextPageAsync().Result;

                foreach (var item in pages.CurrentPage)
                {
                    if (--count < 0)
                    {
                        yield break;
                    }
                    else if (filter(item))
                    {
                        yield return item;
                    }
                }
            }
        }
#endif
        private class UserHttp
        {
            public string UserPrincipalName { get; set; }
            public string DisplayName { get; set; }
            public string GivenName { get; set; }
            public string SurName { get; set; }
        }

        private class GroupHttp
        {
            public string DisplayName { get; set; }
            public string ObjectId { get; set; }
        }
    }
}
