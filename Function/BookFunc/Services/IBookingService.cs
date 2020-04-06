using Microsoft.Build.Utilities;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using TeamsMeetingBookFunc.Models;

namespace TeamsMeetingBookFunc.Services
{
	public interface IBookingService
	{
		Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel model);
	}
}
