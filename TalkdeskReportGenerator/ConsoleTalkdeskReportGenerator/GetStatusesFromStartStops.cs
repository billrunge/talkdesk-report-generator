using System;
using System.Collections.Generic;

namespace ConsoleTalkdeskReportGenerator
{
    internal class GetStatusesFromStartStops
    {
        public TimeZoneInfo TimeZoneInfo { get; set; } = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
        private readonly IGetStatuses _getStatuses;

        public GetStatusesFromStartStops(IGetStatuses getStatuses)
        {
            _getStatuses = getStatuses;
        }

        public List<AgentStatuses> GetAgentStatusesList(List<AgentStartStop> agentStartStops, DateTime monday)
        {
            List<AgentStatuses> agentStatusesList = new List<AgentStatuses>();
            monday = monday.Date;
            if (monday.DayOfWeek == DayOfWeek.Monday)
            {         
                 
                int utcOffset = Math.Abs(TimeZoneInfo.GetUtcOffset(monday).Hours);

                foreach (AgentStartStop agentStartStop in agentStartStops)
                {
                    string userId = _getStatuses.GetUserIdFromName(agentStartStop.AgentName);

                    AgentStatuses agentStatuses = new AgentStatuses()
                    {
                        AgentName = agentStartStop.AgentName
                    };

                    foreach (StartStop startStop in agentStartStop.StartStopList)
                    {
                        
                        for (int mondayOffset = 0; mondayOffset < 5; mondayOffset++)
                        {                                                    
                            DateTime day = monday.AddDays(mondayOffset);
                            DateTime startTime = day.Add(startStop.Start);
                            DateTime stopTime = day.Add(startStop.Stop);                    

                            List<Status> agentStatus = _getStatuses.GetStatusesList(userId, startTime, stopTime, utcOffset);
                            agentStatuses.Statuses.AddRange(agentStatus);
                        }
                    }
                    agentStatusesList.Add(agentStatuses);
                }
            }
            else
            {
                throw new ArgumentException($"{monday.ToString()} is not a Monday");
            }
            return agentStatusesList;
        }


    }
}
