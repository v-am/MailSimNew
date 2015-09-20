using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using MailSim.Common.Contracts;

namespace MailSim.ProvidersREST
{
    class AddressBookProviderHTTP : IAddressBook
    {
        public AddressBookProviderHTTP()
        {
        }

        public IEnumerable<string> GetUsers(string match, int count)
        {
            return EnumerateUsers(match, count);
        }

        public IEnumerable<string> GetDLMembers(string dLName, int count)
        {
            string uri = "https://graph.windows.net/myorganization/groups?api-version=beta";

            if (string.IsNullOrEmpty(dLName))
            {
                return Enumerable.Empty<string>();
            }
#if true
            return Enumerable.Empty<string>();
#else
            var pagedGroups = _adClient.Groups
                    .Where(g => g.DisplayName.StartsWith(dLName))
                    .ExecuteAsync()
                    .Result;

            // Look for the group with exact match
            var group = GetFilteredItems(pagedGroups, int.MaxValue,
                               (g) => g.DisplayName.EqualsCaseInsensitive(dLName))
                               .FirstOrDefault();

            if (group == null)
            {
                return Enumerable.Empty<string>();
            }

            var groupFetcher = group as IGroupFetcher;

            var pagedMembers = groupFetcher.Members.ExecuteAsync().Result;

            var members = GetFilteredItems(pagedMembers, count, (member) => member is User);

            return members.Select(m => (m as User).UserPrincipalName);
#endif
        }

        private string AppendVersion(string uri)
        {
            return uri + "api-version=beta";
        }

        private IEnumerable<string> EnumerateUsers(string match, int count)
        {
            string uri = "https://graph.windows.net/myorganization/users";
            uri += '?';     // we always have parameters

            if (string.IsNullOrEmpty(match) == false)
            {
                uri = AddFilters(uri, match,
                            "userPrincipalName",
                            "displayName",
                            "givenName"/*, "surName"*/);

                uri += '&';
            }

            uri = AppendVersion(uri);

//            var users = HttpUtil.GetItemsAsync<IEnumerable<UserHttp>>(uri, AuthenticationHelperHTTP.GetAADToken).Result;
            var users = new HttpUtilSync(Constants.AadServiceResourceId).EnumerateCollection<UserHttp>(uri, count);

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
        private IEnumerable<string> EnumerateUsers(string match, int count)
        {
            IPagedCollection<IUser> pagedUsers;

            if (string.IsNullOrEmpty(match))
            {
                pagedUsers = _adClient.Users
                    .ExecuteAsync()
                    .Result;
            }
            else
            {
                // Apply server-side filtering
                pagedUsers = _adClient.Users
                    .Where(x =>
                        x.UserPrincipalName.StartsWith(match) ||
                        x.DisplayName.StartsWith(match) ||
                        x.GivenName.StartsWith(match) ||
                        x.Surname.StartsWith(match)
                    )
                    .ExecuteAsync()
                    .Result;
            }

            var users = GetFilteredItems(pagedUsers, count, (u) => true);

            return users.Select(u => u.UserPrincipalName);
        }

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
    }
}
