using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace TeamsMeetingBookFunc.Models
{
    [SuppressMessage("Design", "CA1051:Do not declare visible instance fields", Justification = "Data Transfer Object")]
    public class RequestModel
    {
        public DateTime? StartDateTime;
        public int MeetingDurationMins;
        public string Subject;
        public string PatientEmailAddress;
        public string DoctorEmailAddress;
        public List<string> OptionalAttendees;
    }
}
