using System.Collections.Generic;

namespace TalkdeskReportGenerator.Library
{
    public class AgentStartStops
    {
        public string AgentName { get; set; }
        public List<StartStop> StartStopList { get; set; } = new List<StartStop>();
    }
}
