using System;
using System.Collections.Generic;

namespace WpfTalkdeskReportGenerator
{
    internal interface IGetStatusesFromStartStops
    {
        List<AgentStatuses> GetAgentStatusesList(IGetStatuses getStatuses, List<AgentStartStops> agentStartStops, DateTime monday);
    }

    internal class GetStatusesFromStartStops : IGetStatusesFromStartStops
    {
        private readonly TimeZoneInfo _excelTimeZone;

        public GetStatusesFromStartStops()
        {
            _excelTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");

        }

        public List<AgentStatuses> GetAgentStatusesList(IGetStatuses getStatuses, List<AgentStartStops> agentStartStops, DateTime day)
        {
            List<AgentStatuses> agentStatusesList = new List<AgentStatuses>();

            int utcOffset = Math.Abs(_excelTimeZone.GetUtcOffset(day).Hours);

            foreach (AgentStartStops agentStartStop in agentStartStops)
            {
                string userId = getStatuses.GetUserIdFromName(agentStartStop.AgentName);

                AgentStatuses agentStatuses = new AgentStatuses()
                {
                    AgentName = agentStartStop.AgentName
                };

                foreach (StartStop startStop in agentStartStop.StartStopList)
                {
                    DateTime startTime = day.Add(startStop.Start);
                    DateTime stopTime = day.Add(startStop.Stop);

                    List<Status> agentStatus = getStatuses.GetStatusesList(userId, startTime, stopTime, utcOffset);
                    agentStatuses.Statuses.AddRange(agentStatus);
                }
                agentStatusesList.Add(agentStatuses);
            }

            return agentStatusesList;
        }

    }
}
