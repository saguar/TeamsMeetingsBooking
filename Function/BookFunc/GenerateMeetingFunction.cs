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
using TeamsMeetingBookFunc.Models;
using TeamsMeetingBookFunc.Services;
using TeamsMeetingBookFunc.Helpers;

namespace TeamsMeetingBookingFunction
{
    public static class GenerateMeetingFunction
    {
        [FunctionName("GenerateMeetingFunction")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
            RequestModel requestModel,
            // HttpRequest is still passed on but currently not used
            HttpRequest _,
            ILogger log, ExecutionContext context)
        {
            var config = BuildConfig(context);

            // use defaults if required
            requestModel.StartDateTime ??= DateTime.Now;
            requestModel.EndDateTime ??= DateTime.Now.AddHours(1);
            requestModel.Subject ??= config.GetConnectionStringOrSetting(ConfigConstants.DefaultMeetingNameCfg);

            try
            {
                var service = new BookingService(config);

                var onlineMeeting = await service.CreateTeamsMeetingAsync(requestModel).ConfigureAwait(false);

                var newEvent = await service.CreateCalendarEventAsync(requestModel, onlineMeeting.JoinWebUrl).ConfigureAwait(false);

                var result = new
                {
                    meetingUrl = onlineMeeting.JoinWebUrl,
                    eventId = newEvent.Id,
                    meetingId = onlineMeeting.Id
                };

                return new OkObjectResult(result);
            }
            catch (ServiceException e)
            {
                log.LogError($"Error:\n{e}");
                return new ObjectResult($"\"Can't perform request now - {e.Message}\"") { StatusCode = 500 };
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
    }
}
