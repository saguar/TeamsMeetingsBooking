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
        Task<Event> CreateCalendarEventAsync(RequestModel requestModel, string bodyText);
        Task<OnlineMeeting> CreateTeamsMeetingAsync(RequestModel requestModel);
    }
}
