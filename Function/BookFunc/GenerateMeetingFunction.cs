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
	public class GenerateMeetingFunction
	{
		private readonly IConfiguration config;
		private readonly IBookingService bookingSvc;

		public GenerateMeetingFunction(IConfiguration _config, IBookingService _bookingSvc)
		{
			config = _config ?? throw new ArgumentNullException(nameof(_config));
			bookingSvc = _bookingSvc ?? throw new ArgumentNullException(nameof(_bookingSvc));
		}
		[FunctionName("GenerateMeetingFunction")]
		public async Task<IActionResult> Run(
			[HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
			RequestModel requestModel,
			ILogger log)
		{
			
			//StartDateTime is mandatory. Return BadRequest if not passed or if input format is invalid
			if(!requestModel.StartDateTime.HasValue)
			{
				log.LogError($"{nameof(RequestModel.StartDateTime)} is null. Invalid format or parameter not passed. Returning BadRequest");
				return new BadRequestErrorMessageResult($"{nameof(RequestModel.StartDateTime)} not present or invalid. Please use the format YYYY-mm-DDTHH:mm:ss");
			}

			if(requestModel.MeetingDurationMins == 0)
			{
				requestModel.MeetingDurationMins = config.GetValue<int>(ConfigConstants.DefaultMeetingDurationMinsCfg);
			}

			requestModel.Subject ??= config.GetValue<string>(ConfigConstants.DefaultMeetingNameCfg);
			
			try
			{
				log.LogInformation("Creating a meeting with following info: StartDateTime = {startDateTime}, DurationMins = {durationMins}, Subject = {subject}",
					requestModel.StartDateTime, requestModel.MeetingDurationMins, requestModel.Subject);
				
				var onlineMeeting = await bookingSvc.CreateTeamsMeetingAsync(requestModel).ConfigureAwait(false);
				
				log.LogInformation("Meeting created. MeetingUrl = {meetingUrl}, MeetingId = {meetingId}", onlineMeeting.JoinWebUrl, onlineMeeting.Id);

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

				log.LogError(e, "An error occurred invoking the Microsoft Graph API using StartDateTime = {startDateTime}, DurationMins = {durationMins}, Subject = {subject}",
					requestModel.StartDateTime, requestModel.MeetingDurationMins, requestModel.Subject);

				if (e.StatusCode == System.Net.HttpStatusCode.BadRequest)
				{
					return new BadRequestErrorMessageResult(e.Message);
				}

				return new InternalServerErrorResult();
			}
		}


	}
}
