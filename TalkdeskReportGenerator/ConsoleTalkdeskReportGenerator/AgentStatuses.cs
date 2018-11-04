using System;
using System.Collections.Generic;

namespace ConsoleTalkdeskReportGenerator
{
    class AgentStatuses
    {
        public string AgentName { get; set; }
        public List<Status> Statuses { get; set; } = new List<Status>();
    }
}
