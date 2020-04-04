using System;
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
            HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            // use defaults if required
            requestModel.StartDateTime ??= DateTime.Now;
            requestModel.EndDateTime ??= DateTime.Now.AddHours(1);
            requestModel.Subject ??= BookingService.Current.Configuration.GetConnectionStringOrSetting(ConfigConstants.DefaultMeetingNameCfg);

            try
            {

                var onlineMeeting = await BookingService.Current.CreateTeamsMeetingAsync(requestModel).ConfigureAwait(false);

                var result = new
                {
                    meetingUrl = onlineMeeting.JoinWebUrl,
                    meetingId = onlineMeeting.Id,
                    meetingName = onlineMeeting.Subject
                };

                return new OkObjectResult(result);
            }
            catch (ServiceException e)
            {
                log.LogError($"Error:\n{e}");
                return new ObjectResult($"\"Can't perform request now - {e.Message}\"") { StatusCode = 500 };
            }
        }


    }
}
