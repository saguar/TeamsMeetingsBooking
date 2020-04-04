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
using System.Web.Http;

namespace TeamsMeetingBookingFunction
{
	public static class GenerateMeetingFunction
	{
		[FunctionName("GenerateMeetingFunction")]
		public static async Task<IActionResult> Run(
			[HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
			RequestModel requestModel,
			ILogger log)
		{
			//You can't specify only EndDateTime
			if (requestModel.EndDateTime != null && requestModel.StartDateTime == null)
			{
				return new BadRequestErrorMessageResult($"If you specify {nameof(requestModel.EndDateTime)}, you must specify {nameof(requestModel.StartDateTime)} as well");
			}

			requestModel.StartDateTime ??= DateTime.Now;
			requestModel.EndDateTime ??= requestModel.StartDateTime.Value.AddHours(1);
			requestModel.Subject ??= BookingService.Current.Configuration.GetConnectionStringOrSetting(ConfigConstants.DefaultMeetingNameCfg);

			if (requestModel.EndDateTime.Value < requestModel.StartDateTime.Value)
			{
				return new BadRequestErrorMessageResult($"{nameof(requestModel.EndDateTime)} must be after {nameof(requestModel.StartDateTime)}");
			}
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

				log.LogError(e, "An error occurred invoking the Microsoft Graph API using StartDate = {startDate}, EndDate = {endDate}, Subject = {subject}",
					requestModel.StartDateTime, requestModel.EndDateTime, requestModel.Subject);

				if (e.StatusCode == System.Net.HttpStatusCode.BadRequest)
				{
					return new BadRequestErrorMessageResult(e.Message);
				}

				return new InternalServerErrorResult();
			}
		}


	}
}
