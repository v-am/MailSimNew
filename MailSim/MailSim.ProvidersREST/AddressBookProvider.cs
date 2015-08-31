using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;

namespace MailSim.ProvidersREST
{
    class AddressBookProvider : IAddressBook
    {
        private readonly ActiveDirectoryClient _adClient;

        public AddressBookProvider(ActiveDirectoryClient adClient)
        {
            _adClient = adClient;
        }

        /// <summary>
        /// Builds list of addresses for all users in the Address List that have display name match
        /// </summary>
        /// <param name="match"> string to match in user name or null to return all users in the GAL</param>
        /// <returns>List of SMTP addresses of matching users in the address list. The list will be empty if no users exist or match.</returns>
        public IEnumerable<string> GetUsers(string match, int count)
        {
            match = match ?? string.Empty;

            return EnumerateUsers(match, count);
        }

        private IEnumerable<string> EnumerateUsers(string match, int count)
        {
            var pagedUsers = _adClient.Users
                .Where(
                    x => x.UserPrincipalName.StartsWith(match) ||
                    x.DisplayName.StartsWith(match) ||
                    x.GivenName.StartsWith(match) ||
                    x.Surname.StartsWith(match)
                )
                .ExecuteAsync()
                .Result;

            var users = GetFilteredItems(pagedUsers, count, (u) => true);

            return users.Select(u => u.UserPrincipalName);
        }

        /// <summary>
        /// Builds list of addresses for all members of Exchange Distribution list in the Address List
        /// </summary>
        /// <param name="dLName">Exchane Distribution List Name</param>
        /// <returns>List of SMTP addresses of DL members or null if DL is not found. Nesting DLs are not expanded. </returns>
        public IEnumerable<string> GetDLMembers(string dLName, int count)
        {
            var groups = _adClient.Groups
                .Where(g => g.Mail.StartsWith(dLName))
                .ExecuteAsync()
                .Result
                .CurrentPage;   // assume we are going to use the first matching group

            if (groups.Any() == false)
            {
                return Enumerable.Empty<string>();
            }

            var group = groups.First() as Group;
            IGroupFetcher groupFetcher = group;

            var pages = groupFetcher.Members.ExecuteAsync().Result;

            var members = GetFilteredItems(pages, count, (member) => member is User);

            return members.Select(m => (m as User).UserPrincipalName);
        }

        private IEnumerable<T> GetFilteredItems<T>(IPagedCollection<T> pages, int count, Func<T, bool> filter)
        {
            foreach (var item in pages.CurrentPage)
            {
                if (filter(item) && count-- > 0)
                {
                    yield return item;
                }
            }

            while (count > 0 && pages.MorePagesAvailable)
            {
                pages = pages.GetNextPageAsync().Result;

                foreach (var item in pages.CurrentPage)
                {
                    if (filter(item) && count-- > 0)
                    {
                        yield return item;
                    }
                }
            }
        }
    }
}
