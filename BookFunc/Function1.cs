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
using TeamsMeetingBookFunc;

namespace TeamsMeetingBookingFunction
{
    public static class Function1
    {
        #region config keys
        private const string TenantIdCfg = "TenantID";
        private const string ClientIdCfg = "ClientID";
        private const string UserPasswordCfg = "UserPassword";
        private const string UserEmailCfg = "UserEmail";
        private const string DefaultMeetingNameCfg = "DefaultMeetingName";
        #endregion

        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            var config = BuildConfig(context);

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            DateTime startDate = data?.startDateTime ?? DateTime.Now;
            DateTime endDate = data?.endDateTime ?? DateTime.Now.AddHours(12);

            string meetingName = data?.subject ?? config.GetConnectionStringOrSetting(DefaultMeetingNameCfg);

            try
            {
                var authProvider = GetAuthenticationProvider(config);

                var graphServiceClient = new GraphServiceClient(authProvider);

                var onlineMeeting = new OnlineMeeting
                {
                    StartDateTime = startDate,
                    EndDateTime = endDate,
                    Subject = meetingName
                };

                var meeting = await graphServiceClient.Me.OnlineMeetings.Request()
                    .AddAuthenticationToRequest(config.GetConnectionStringOrSetting(UserEmailCfg), config.GetConnectionStringOrSetting(UserPasswordCfg))
                    .WithMaxRetry(5)
                    .AddAsync(onlineMeeting);

                return new OkObjectResult(meeting.JoinWebUrl);
            }
            catch(ServiceException e)
            {
                log.LogError(e.ToString());
                return new ObjectResult($"Can't perform request now - {e.Message}") { StatusCode = 500 };
            }
        }

        private static IConfigurationRoot BuildConfig(ExecutionContext context)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(context.FunctionAppDirectory)
                .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();
            return config;
        }

        private static UsernamePasswordProvider GetAuthenticationProvider(IConfigurationRoot config)
        {
            string authority = $"https://login.microsoftonline.com/{config.GetConnectionStringOrSetting(TenantIdCfg)}";

            var app = PublicClientApplicationBuilder.Create(config.GetConnectionStringOrSetting(ClientIdCfg))
                .WithAuthority(authority)
                .Build();

            return new UsernamePasswordProvider(app);
        }
    }
}
