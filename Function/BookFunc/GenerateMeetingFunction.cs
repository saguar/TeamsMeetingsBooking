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
using System.Web.Http;

namespace TeamsMeetingBookingFunction
{
    public class GenerateMeetingFunction
    {
        private readonly IConfiguration config;
        private readonly IBookingService bookingSvc;

        public GenerateMeetingFunction(IConfiguration config, IBookingService bookingSvc)
        {
            this.config = config ?? throw new ArgumentNullException(nameof(config));
            this.bookingSvc = bookingSvc ?? throw new ArgumentNullException(nameof(bookingSvc));
        }

        [FunctionName("GenerateMeetingFunction")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
            RequestModel requestModel,
            ILogger log, ExecutionContext context)
        {
            // parameter check
            try
            {
                if (requestModel is null)
                {
                    throw new ArgumentNullException(nameof(requestModel), "Please check if POST body contains a valid JSON request.");
                }

                if (log is null)
                {
                    // do nothing - we invoke log with ?. operator
                }

                if (context is null)
                {
                    throw new ArgumentNullException(nameof(context), "Azure Function runtime error - no execution context provided.");
                }
            }
            catch (ArgumentNullException e)
            {
                return new ObjectResult($"\"{e.Message}\"") { StatusCode = 500 };
            }

            //StartDateTime is mandatory. Return BadRequest if not passed or if input format is invalid
            if (!requestModel.StartDateTime.HasValue)
            {
                log.LogError($"{nameof(RequestModel.StartDateTime)} is null. Invalid format or parameter not passed. Returning BadRequest");
                return new BadRequestErrorMessageResult($"{nameof(RequestModel.StartDateTime)} not present or invalid. Please use the format YYYY-mm-DDTHH:mm:ss");
            }

            if (requestModel.MeetingDurationMins == 0)
            {
                requestModel.MeetingDurationMins = config.GetValue<int>(ConfigConstants.DefaultMeetingDurationMinsCfg);
            }

            requestModel.Subject ??= config.GetConnectionStringOrSetting(ConfigConstants.DefaultMeetingNameCfg);

            try
            {

                var onlineMeeting = await bookingSvc.CreateTeamsMeetingAsync(requestModel).ConfigureAwait(false);
                var eventId = "No event was requested";
                
                if (requestModel.CreateEvent ?? false)
                {
                    var newEvent = await bookingSvc.CreateCalendarEventAsync(requestModel, onlineMeeting.JoinWebUrl).ConfigureAwait(false);
                    eventId = newEvent?.Id ?? "No event could be created - doctor's account was not specified";
                }

                var result = new
                {
                    meetingUrl = onlineMeeting.JoinWebUrl,
                    eventId = eventId,
                    meetingId = onlineMeeting.Id
                };

                return new OkObjectResult(result);
            }
            catch (ServiceException e)
            {
                return LogAndReturnErrorResult(log, $"An error occurred invoking the Microsoft Graph API using StartDateTime = {requestModel.StartDateTime}" +
                    $", DurationMins = {requestModel.MeetingDurationMins}, Subject = {requestModel.Subject}", e);
            }
            catch (InvalidOperationException e)
            {
                return LogAndReturnErrorResult(log, "Invalid request", e);
            }
        }

        private static IActionResult LogAndReturnErrorResult(ILogger log, string message, Exception e)
        {
            if (e is ServiceException se && se.StatusCode == System.Net.HttpStatusCode.BadRequest)
            {
                return new BadRequestErrorMessageResult(e.Message);
            }

            log?.LogError($"{message}:\n{e}");
            return new ObjectResult($"\"{message}: - {e.Message}\"") { StatusCode = 500 };
        }
    }
}
