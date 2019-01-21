using log4net;
using System.Collections.Generic;
using System.Linq;

namespace WpfTalkdeskReportGenerator
{
    interface IConsolidateAgentStatuses
    {
        List<AgentStatuses> Consolidate(List<AgentStatuses> agentStatusesList);
    }
    class ConsolidateAgentStatuses : IConsolidateAgentStatuses
    {
        private readonly ILog _log;

        public ConsolidateAgentStatuses(ILog log)
        {
            _log = log;
        }

        public List<AgentStatuses> Consolidate(List<AgentStatuses> agentStatusesList)
        {
            List<AgentStatuses> outputList = new List<AgentStatuses>();


            _log.Debug("Looping through all AgentStatuses lists in agentStatusesList");
            foreach (AgentStatuses agentStatuses in agentStatusesList)
            {
                List<Status> newStatuses = new List<Status>();

                _log.Debug("Using LINQ to group AgentStatuses by day and status label");

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

                _log.Debug("Creating new AgentStatuses object from output of LINQ query");

                AgentStatuses newAgentStatuses = new AgentStatuses
                {
                    AgentName = agentStatuses.AgentName,
                    Statuses = newStatuses                    
                };

                _log.Debug("Adding AgentStatuses to output list");
                outputList.Add(newAgentStatuses);

            }
            return outputList;
        }
    }
}
