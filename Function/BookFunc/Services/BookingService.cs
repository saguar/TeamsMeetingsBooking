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

namespace TeamsMeetingBookFunc.Services
{
    class BookingService
    {
        private static readonly string binDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static readonly string rootDirectory = Path.GetFullPath(Path.Combine(binDirectory, ".."));
        internal IConfigurationRoot Configuration { get; private set; }

        internal GraphServiceClient GraphServiceClient { get; private set; }
private static readonly string binDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static readonly string rootDirectory = Path.GetFullPath(Path.Combine(binDirectory, ".."));
        #region Lazy and thread-safe singleton
        private static readonly Lazy<BookingService> _current = new Lazy<BookingService>(() => new BookingService());

        internal static BookingService Current => _current.Value;

        private BookingService()
        {
            Configuration = BuildConfig();

            Uri authority = new Uri($"https://login.microsoftonline.com/{Configuration.GetConnectionStringOrSetting(ConfigConstants.TenantIdCfg)}");

            var app = PublicClientApplicationBuilder.Create(Configuration.GetConnectionStringOrSetting(ConfigConstants.ClientIdCfg))
                .WithAuthority(authority)
                .Build();

            GraphServiceClient = new GraphServiceClient(new UsernamePasswordProvider(app));
        }
        #endregion

        private IConfigurationRoot BuildConfig()
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(rootDirectory)
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
