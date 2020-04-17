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
using Microsoft.Azure.Services.AppAuthentication;
using TeamsMeetingBookFunc.Authentication;

namespace TeamsMeetingBookFunc.Services
{
    class BookingService : IBookingService
    {
        private readonly IConfiguration configuration;
        private readonly IGraphServiceClient graphServiceClient;

        public BookingService(IConfiguration configuration, IGraphServiceClient graphServiceClient)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        public async Task<Event> CreateCalendarEventAsync(RequestModel requestModel, string bodyText)
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
                End = requestModel.StartDateTime.Value.AddMinutes(requestModel.MeetingDurationMins).ToDateTimeTimeZone(),
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

            foreach (var emailAddress in requestModel.OptionalAttendees ?? Enumerable.Empty<string>())
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

            if(attendeeList.Count == 0)
            {
                throw new InvalidOperationException("Can't create event with no participants!");
            }

            if (!configuration.IsUsingServicePrincipal())
            {
                return await graphServiceClient.Me.Events.Request()
                    .AddAuthenticationToRequest(configuration.GetAccountEmail(), configuration.GetAccountPassword())
                    .WithMaxRetry(5)
                    .AddAsync(newEvent)
                    .ConfigureAwait(false);
            }
            else
            {
                var organizerEmail = configuration.IsManagedIdentityAuth() ? requestModel.DoctorEmailAddress : configuration.GetAccountEmail();

                if (!String.IsNullOrWhiteSpace(organizerEmail))
                {
                    return await graphServiceClient.Users[organizerEmail].Events.Request()
                        .WithMaxRetry(5)
                        .AddAsync(newEvent)
                        .ConfigureAwait(false);
                }

                return null;
            }
        }

        public async Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel requestModel)
        {
            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = requestModel.StartDateTime,
                EndDateTime = requestModel.StartDateTime.Value.AddMinutes(requestModel.MeetingDurationMins),
                Subject = requestModel.Subject,
            };

            OnlineMeeting meeting;

            if (configuration.IsUsingServicePrincipal())
            {
                onlineMeeting.Participants = new MeetingParticipants
                {
                    Organizer = new MeetingParticipantInfo
                    {
                        Identity = new IdentitySet
                        {
                            User = new Identity
                            {
                                Id = configuration.GetAccountEmail()
                            }
                        }
                    }
                };
                meeting = await graphServiceClient.Communications.OnlineMeetings.Request()
                    .WithMaxRetry(5)
                    .AddAsync(onlineMeeting)
                    .ConfigureAwait(false);
            }
            else
            {
                meeting = await graphServiceClient.Me.OnlineMeetings.Request()
                    .AddAuthenticationToRequest(configuration.GetAccountEmail(), configuration.GetAccountPassword())
                    .WithMaxRetry(5)
                    .AddAsync(onlineMeeting)
                    .ConfigureAwait(false);
            }

            return meeting;
        }
    }
}
