//#define MEETING_TIMES

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using MailSim.Common.Contracts;
using MailSim.Common;

namespace MailSim.ProvidersREST
{
    class AddressBookProviderHTTP : HTTP.BaseProviderHttp, IAddressBook
    {
        private const string BaseUri = "https://graph.windows.net/myorganization/";
        private const string ApiVersion = "api-version=1.5";

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

            var groups = httpProxy.GetItems<GroupHttp>(uri, 100);

            // Look for the group with exact name match
            var group = groups.FirstOrDefault((g) => g.DisplayName.EqualsCaseInsensitive(dLName));

            if (group == null)
            {
                return Enumerable.Empty<string>();
            }

            uri = BaseUri + "groups/" + group.ObjectId + "/members?" + ApiVersion;

            var members = httpProxy.GetItems<UserHttp>(uri, count);

            return members.Select(x => x.UserPrincipalName);
        }

        public IEnumerable<string> GetUsers(string match, int count)
        {
#if MEETING_TIMES
            MeetingTimes times = new MeetingTimes();

            MeetingTimeSlot timeSlot = new MeetingTimeSlot();
            timeSlot.Start = new TimeDesc("2015-10-08", "1:00:00", "GMT Standard Time");
            timeSlot.End = new TimeDesc("2015-10-08", "23:00:00", "GMT Standard Time");

            AttendeeBase att1 = new AttendeeBase();
//            att1.EmailAddress.Address = "andreida@microsoft.com";
            att1.EmailAddress.Address = "jill@ointerop.onmicrosoft.com";

            AttendeeBase att2 = new AttendeeBase();
            att2.EmailAddress.Address = "v-am@microsoft.com";

            times.TimeConstraint.Timeslots.Add(timeSlot);
            times.LocationConstraint = new LocationConstraint();

            times.Attendees.Add(att1);
//            times.Attendees.Add(att2);

            var httpProxy = new HttpUtilSync(Constants.OfficeResourceId);
            String uri = "https://outlook.office365.com/api/beta/me/findmeetingtimes";

            var res = httpProxy.PostItem2<MeetingTimes, MeetingTimeCandidates>(uri, times);

            return null;
#else
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
                    .GetItems<UserHttp>(uri, count);

            return users.Select(x => x.UserPrincipalName);
#endif
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

                sb.AppendFormat("startswith({0},'{1}')", fields[i], match);
            }

            return sb.ToString();
        }

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
#if MEETING_TIMES

        private class MeetingTimes
        {
            public List<AttendeeBase> Attendees { get; set; }
            public TimeConstraint TimeConstraint { get; set; }
            public string MeetingDuration = "PT15M";
            public LocationConstraint LocationConstraint;

            public MeetingTimes()
            {
                Attendees = new List<AttendeeBase>();
                TimeConstraint = new TimeConstraint();
            }
        }

        private class LocationConstraint
        {
            public bool IsRequired;
            public bool SuggestLocation;
            public List<Location> Locations { get; set; }

            public LocationConstraint()
            {
                Locations = new List<Location>();
                Locations.Add(new Location());
            }
        }

        private class Location
        {
            public string DisplayName = "Starbucks";
        }

        private class TimeConstraint
        {
            public List<MeetingTimeSlot> Timeslots { get; set; }

            public TimeConstraint()
            {
                Timeslots = new List<MeetingTimeSlot>();
            }
        }

        private class MeetingTimeSlot
        {
            public TimeDesc Start { get; set; }
            public TimeDesc End { get; set; }
        }

        private class AttendeeBase
        {
            public EmailAddress EmailAddress { get; set; }
            public string Type = "Optional";

            public AttendeeBase()
            {
                EmailAddress = new EmailAddress();
            }
        }

        private class TimeDesc
        {
            public String Date { get; set; }
            public String Time { get; set; }
            public String TimeZone { get; set; }

            public TimeDesc(String date, String time, String timeZone)
            {
                Date = date;
                Time = time;
                TimeZone = timeZone;
            }
        }

        private class EmailAddress
        {
            public string Address { get; set; }
            public string Name { get; set; }
        }

        private class MeetingTimeCandidates
        {
            public List<MeetingTimeCandidate> value;
        }

        private class MeetingTimeCandidate
        {
            public MeetingTimeSlot MeetingTimeSlot;
        }
#endif
    }
}
