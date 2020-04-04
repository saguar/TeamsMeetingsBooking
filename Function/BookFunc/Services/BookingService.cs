using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using TeamsMeetingBookFunc.Models;
using TeamsMeetingBookFunc.Helpers;
using Microsoft.Azure.WebJobs;

namespace TeamsMeetingBookFunc.Services
{
    class BookingService
    {
        internal IConfigurationRoot Configuration { get; private set; }

        internal GraphServiceClient GraphServiceClient { get; private set; }

        #region Lazy and thread-safe singleton
        private static readonly Lazy<BookingService> _current = new Lazy<BookingService>(() => new BookingService());

        internal static BookingService Current => _current.Value;

        private BookingService()
        {

        }
        #endregion

        internal void Init(ExecutionContext context){

           Configuration = BuildConfig(context);

            Uri authority = new Uri($"https://login.microsoftonline.com/{Configuration.GetConnectionStringOrSetting(ConfigConstants.TenantIdCfg)}");

            var app = PublicClientApplicationBuilder.Create(Configuration.GetConnectionStringOrSetting(ConfigConstants.ClientIdCfg))
                .WithAuthority(authority)
                .Build();

            GraphServiceClient = new GraphServiceClient(new UsernamePasswordProvider(app));
        }

        private IConfigurationRoot BuildConfig(ExecutionContext context)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(context.FunctionAppDirectory)
                .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();
            return config;
        }

        internal async Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel requestModel)
        {
            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = requestModel.StartDateTime,
                EndDateTime = requestModel.EndDateTime,
                Subject = requestModel.Subject
            };

            var meeting = await GraphServiceClient.Me.OnlineMeetings.Request()
                .AddAuthenticationToRequest(Configuration.GetConnectionStringOrSetting(ConfigConstants.UserEmailCfg), Configuration.GetConnectionStringOrSetting(ConfigConstants.UserPasswordCfg))
                .WithMaxRetry(5)
                .AddAsync(onlineMeeting).ConfigureAwait(false);
            return meeting;
        }

    }
}
