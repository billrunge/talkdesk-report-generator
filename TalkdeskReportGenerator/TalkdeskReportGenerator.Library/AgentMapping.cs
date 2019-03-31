using System;

namespace TalkdeskReportGenerator.Library
{
    public class AgentMapping
    {
        public string ExcelAgentName { get; set; }
        public string TalkdeskAgentName { get; set; }
        public Guid GUID { get; }

        public AgentMapping()
        {
            GUID = Guid.NewGuid();
        }
    }
}
