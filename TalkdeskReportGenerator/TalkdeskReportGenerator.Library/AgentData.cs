using System.Collections.Generic;

namespace TalkdeskReportGenerator.Library
{
    public class AgentData
    {
        public string AgentName { get; set; }
        public List<Status> Statuses { get; set; } = new List<Status>();
        public List<Call> Calls { get; set; } = new List<Call>();
    }
}
