using System;
using System.Collections.Generic;

namespace TeamsMeetingBookFunc.Models
{
    public class RequestModel
    {
        public DateTime? StartDateTime;
        public DateTime? EndDateTime;
        public string Subject;
        public string PatientEmailAddress;
        public string DoctorEmailAddress;
        public List<string> OptionalAttendees;
    }
}
