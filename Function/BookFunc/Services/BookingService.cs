using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using TeamsMeetingBookFunc.Models;
using TeamsMeetingBookFunc.Helpers;
using Microsoft.Azure.WebJobs;
using System.IO;
using System.Reflection;
using Microsoft.Build.Framework;

namespace TeamsMeetingBookFunc.Services
{
    public class BookingService : IBookingService
    {
        private readonly IGraphServiceClient graphClient;
        private readonly IConfiguration config;

        public BookingService(IConfiguration _config, IGraphServiceClient _graphClient)
        {
            graphClient = _graphClient ?? throw new ArgumentNullException(nameof(_graphClient));
            config = _config ?? throw new ArgumentNullException(nameof(_config));
        }

        public async Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel requestModel)
        {
            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = requestModel.StartDateTime,
                EndDateTime = requestModel.StartDateTime.Value.AddMinutes(requestModel.MeetingDurationMins),
                Subject = requestModel.Subject
            };

            var meeting = await graphClient.Me.OnlineMeetings.Request()
                .AddAuthenticationToRequest(config.GetValue<string>(ConfigConstants.UserEmailCfg), config.GetValue<string>(ConfigConstants.UserPasswordCfg))
                .WithMaxRetry(5)
                .AddAsync(onlineMeeting).ConfigureAwait(false);
            return meeting;
        }

    }
}
