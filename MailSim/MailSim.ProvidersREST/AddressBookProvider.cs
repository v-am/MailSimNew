using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;

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
        public IEnumerable<string> GetUsers(string match)
        {
            match = match ?? string.Empty;

            return EnumerateUsers()
                .Where(x => x.IndexOf(match, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private IEnumerable<string> EnumerateUsers()
        {
            var pagedUsers = _adClient.Users
                .Where(x => x.UserPrincipalName.StartsWith("oi"))
                .ExecuteAsync().Result;

            foreach (var item in pagedUsers.CurrentPage)
            {
                yield return item.UserPrincipalName;
            }

            while (pagedUsers.MorePagesAvailable)
            {
                pagedUsers = pagedUsers.GetNextPageAsync().Result;

                foreach (var item in pagedUsers.CurrentPage)
                {
                    yield return item.UserPrincipalName;
                }
            }
        }

        /// <summary>
        /// Builds list of addresses for all members of Exchange Distribution list in the Address List
        /// </summary>
        /// <param name="dLName">Exchane Distribution List Name</param>
        /// <returns>List of SMTP addresses of DL members or null if DL is not found. Nesting DLs are not expanded. </returns>
        public IEnumerable<string> GetDLMembers(string dLName)
        {
            var groups = _adClient.Groups
                .Where(g => g.Mail.StartsWith(dLName))
                .ExecuteAsync().Result.CurrentPage;

            if (groups.Any() == false)
            {
                yield break;
            }

            var group = groups.First() as Group;
            IGroupFetcher groupFetcher = group;

            var members = groupFetcher.Members.ExecuteAsync().Result;

            foreach (var member in members.CurrentPage)
            {
                if (member is User)
                {
                    var user = member as User;
                    yield return user.UserPrincipalName;
                }
            }
        }
    }
}
