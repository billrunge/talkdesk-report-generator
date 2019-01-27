using System;
using System.Collections.Generic;

namespace TalkdeskReportGenerator.Assemblies
{
    public class AgentStatuses
    {
        public string AgentName { get; set; }
        public List<Status> Statuses { get; set; } = new List<Status>();
    }
}
