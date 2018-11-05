using System.Collections.Generic;
using System.Linq;

namespace ConsoleTalkdeskReportGenerator
{

    internal class ConsolidateAgentStatuses
    {
        public List<AgentStatuses> Consolidate(List<AgentStatuses> agentStatusesList)
        {
            List<AgentStatuses> outputList = new List<AgentStatuses>();

            foreach (AgentStatuses agentStatuses in agentStatusesList)
            {
                List<Status> newStatuses = new List<Status>();

                newStatuses = agentStatuses.Statuses
                    .GroupBy(a => new
                    {
                        a.DayName,
                        a.StatusLabel
                    })
                    .Select(a => new Status
                    {
                        DayName = a.Key.DayName,
                        StatusLabel = a.Key.StatusLabel,
                        StatusTime = a.Sum(ag => ag.StatusTime)

                    }).OrderByDescending(a => a.DayName)
                    .ToList();

                AgentStatuses newAgentStatuses = new AgentStatuses
                {
                    AgentName = agentStatuses.AgentName,
                    Statuses = newStatuses                    
                };

                outputList.Add(newAgentStatuses);

            }
            return outputList;
        }
    }
}
