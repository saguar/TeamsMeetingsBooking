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
            // HttpRequest is still passed if required - but currently not used
            // HttpRequest httpRequest,
            ILogger log, ExecutionContext context)
        {
            // parameter check
            try
            {
                if (requestModel is null)
                {
                    throw new ArgumentNullException("Please check if POST body contained valid JSON request.");
                }

                if (log is null)
                {
                    // do nothing - we invoke log with ?. operator
                }

                if (context is null)
                {
                    throw new ArgumentNullException("Azure Function runtime error - no execution context provided.");
                }
            }
            catch (ArgumentNullException e)
            {
                return new ObjectResult($"\"{e.Message}\"") { StatusCode = 500 };
            }

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
                    eventId = newEvent?.Id ?? "No event could be created - doctor's account was not specified",
                    meetingId = onlineMeeting.Id
                };

                return new OkObjectResult(result);
            }
            catch (ServiceException e)
            {
                return LogAndReturnErrorResult(log, "Can't perform request now", e);
            }
            catch (InvalidOperationException e)
            {
                return LogAndReturnErrorResult(log, "Invalid request", e);
            }
        }

        private static IActionResult LogAndReturnErrorResult(ILogger log, string message, Exception e)
        {
            log?.LogError($"{message}:\n{e}");
            return new ObjectResult($"\"{message}: - {e.Message}\"") { StatusCode = 500 };
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
