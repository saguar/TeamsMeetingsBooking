using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Microsoft.Graph.Auth;
using System.Web.Http;
using System.Net;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using TeamsMeetingBookFunc.Models;
using TeamsMeetingBookFunc.Helpers;
using Microsoft.Graph.Extensions;

namespace TeamsMeetingBookFunc.Services
{
    class BookingService
    {
        private readonly IConfigurationRoot configuration;
        private readonly GraphServiceClient graphServiceClient;

        public BookingService(IConfigurationRoot configuration)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

            var authProvider = GetAuthenticationProvider(configuration);

            graphServiceClient = new GraphServiceClient(authProvider);
        }

        internal async Task<Event> CreateCalendarEventAsync(RequestModel requestModel, string bodyText)
        {
            var attendeeList = new List<Attendee>();

            var newEvent = new Event
            {
                Subject = requestModel.Subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = bodyText
                },
                Attendees = attendeeList,
                Start = requestModel.StartDateTime.Value.ToDateTimeTimeZone(),
                End = requestModel.EndDateTime.Value.ToDateTimeTimeZone(),
            };

            if (!String.IsNullOrWhiteSpace(requestModel.DoctorEmailAddress))
            {
                attendeeList.Add(new Attendee
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = requestModel.DoctorEmailAddress
                    },
                    Type = AttendeeType.Required
                }
                );
            }

            if (!String.IsNullOrWhiteSpace(requestModel.PatientEmailAddress))
            {
                attendeeList.Add(new Attendee
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = requestModel.PatientEmailAddress
                    },
                    Type = AttendeeType.Required
                }
                );
            }

            foreach (var emailAddress in requestModel.OptionalAttendees)
            {
                attendeeList.Add(new Attendee
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = emailAddress
                    },
                    Type = AttendeeType.Resource
                }
                );
            }

            return await graphServiceClient.Me.Events.Request()
                .AddAuthenticationToRequest(configuration.GetConnectionStringOrSetting(ConfigConstants.UserEmailCfg), configuration.GetConnectionStringOrSetting(ConfigConstants.UserPasswordCfg))
                .WithMaxRetry(5)
                .AddAsync(newEvent).ConfigureAwait(false);
        }

        internal async Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel requestModel)
        {
            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = requestModel.StartDateTime,
                EndDateTime = requestModel.EndDateTime,
                Subject = requestModel.Subject
            };

            var meeting = await graphServiceClient.Me.OnlineMeetings.Request()
                .AddAuthenticationToRequest(configuration.GetConnectionStringOrSetting(ConfigConstants.UserEmailCfg), configuration.GetConnectionStringOrSetting(ConfigConstants.UserPasswordCfg))
                .WithMaxRetry(5)
                .AddAsync(onlineMeeting).ConfigureAwait(false);
            return meeting;
        }

        private static UsernamePasswordProvider GetAuthenticationProvider(IConfigurationRoot config)
        {
            Uri authority = new Uri($"https://login.microsoftonline.com/{config.GetConnectionStringOrSetting(ConfigConstants.TenantIdCfg)}");

            var app = PublicClientApplicationBuilder.Create(config.GetConnectionStringOrSetting(ConfigConstants.ClientIdCfg))
                .WithAuthority(authority)
                .Build();

            return new UsernamePasswordProvider(app);
        }
    }
}
