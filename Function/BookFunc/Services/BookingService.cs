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
    class BookingService
    {
        private readonly IConfigurationRoot configuration;
        private readonly GraphServiceClient graphServiceClient;

        bool usingServicePrincipal
        {
            get
            {
                return isClientCredentialAuth || isManagedIdentityAuth;
            }
        }

        string accountEmail => configuration[ConfigConstants.UserEmailCfg];

        private bool isManagedIdentityAuth => string.Equals(configuration.GetConnectionStringOrSetting(ConfigConstants.AuthenticationModeCfg), "managedIdentity", StringComparison.InvariantCultureIgnoreCase);

        private bool isClientCredentialAuth => string.Equals(configuration.GetConnectionStringOrSetting(ConfigConstants.AuthenticationModeCfg), "clientSecret", StringComparison.InvariantCultureIgnoreCase);

        private bool isUsernamePasswordAuth => string.Equals(configuration.GetConnectionStringOrSetting(ConfigConstants.AuthenticationModeCfg), "usernamePassword", StringComparison.InvariantCultureIgnoreCase);

        string accountPassword => configuration.GetConnectionStringOrSetting(ConfigConstants.UserPasswordCfg);

        string clientId => configuration.GetConnectionStringOrSetting(ConfigConstants.ClientIdCfg);

        const string baseUrl = "https://graph.microsoft.com/";

        public BookingService(IConfigurationRoot configuration)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

            // GetAuthenticationProvider already relies on configuration
            var authProvider = GetAuthenticationProvider(baseUrl);

            // need to use beta endpoint if creating an online meeting with service principal
            var versionedBaseUrl = baseUrl + (usingServicePrincipal ? "beta" : "v1.0");

            graphServiceClient = new GraphServiceClient(versionedBaseUrl, authProvider);
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

            if (!usingServicePrincipal)
            {
                return await graphServiceClient.Me.Events.Request()
                    .AddAuthenticationToRequest(accountEmail, accountPassword)
                    .WithMaxRetry(5)
                    .AddAsync(newEvent)
                    .ConfigureAwait(false);
            }
            else
            {
                var organizerEmail = isManagedIdentityAuth ? requestModel.DoctorEmailAddress : accountEmail;

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

        internal async Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel requestModel)
        {
            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = requestModel.StartDateTime,
                EndDateTime = requestModel.EndDateTime,
                Subject = requestModel.Subject,
            };

            OnlineMeeting meeting;

            if (usingServicePrincipal)
            {
                onlineMeeting.Participants = new MeetingParticipants
                {
                    Organizer = new MeetingParticipantInfo
                    {
                        Identity = new IdentitySet
                        {
                            User = new Identity
                            {
                                Id = accountEmail
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
                    .AddAuthenticationToRequest(accountEmail, accountPassword)
                    .WithMaxRetry(5)
                    .AddAsync(onlineMeeting)
                    .ConfigureAwait(false);
            }

            return meeting;
        }

        private IAuthenticationProvider GetAuthenticationProvider(string resourceUrl)
        {
            Uri authority = new Uri($"https://login.microsoftonline.com/{configuration.GetConnectionStringOrSetting(ConfigConstants.TenantIdCfg)}");


            if (isUsernamePasswordAuth)
            {
                var app = PublicClientApplicationBuilder.Create(clientId)
                    .WithAuthority(authority)
                    .Build();
                return new UsernamePasswordProvider(app);
            }

            if (isClientCredentialAuth)
            {
                var app = ConfidentialClientApplicationBuilder.Create(clientId)
                    .WithClientSecret(accountPassword)
                    .WithAuthority(authority)
                    .Build();
                return new ClientCredentialProvider(app);
            }

            if (isManagedIdentityAuth)
            {
                // no Azure Token Service-based provider available in SDK yet
                return new AzureTokenServiceAuthProvider(resourceUrl);
            }

            throw new InvalidOperationException("Unknown AuthenticationMode - use 'usernamePassword' or 'clientCredentials'");
        }
    }
}
