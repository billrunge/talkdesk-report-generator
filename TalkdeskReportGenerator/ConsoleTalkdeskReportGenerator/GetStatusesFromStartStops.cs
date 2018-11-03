using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTalkdeskReportGenerator
{
    class GetStatusesFromStartStops
    {
        private readonly IGetStatuses _getStatuses;

        public GetStatusesFromStartStops(IGetStatuses getStatuses)
        {
            _getStatuses = getStatuses;                  
        }

        public List<AgentStatuses> GetAgentStatusesList(List<AgentStartStop> agentStartStops, DateTime monday)
        {
            List<AgentStatuses> agentStatusesList = new List<AgentStatuses>();
            monday = monday.Date;
            if (monday.DayOfWeek == DayOfWeek.Monday) {

                foreach (var agentStartStop in agentStartStops)
                {
                    string userId = _getStatuses.GetUserIdFromName(agentStartStop.AgentName);

                    AgentStatuses agentStatuses = new AgentStatuses()
                    {
                        AgentName = agentStartStop.AgentName
                    };

                    foreach (var startStop in agentStartStop.StartStopList)
                    {
                        for (int mondayOffset = 0; mondayOffset < 5; mondayOffset++)
                        {
                            DateTime day = monday.AddDays(mondayOffset);

                            DateTime startTime = day.

                            List<Status> agentStatus = _getStatuses.GetStatusesList(userId, );
                        }
                    }

                }
            }
            else
            {
                //The day you entered was not a monday
            }

            return agentStatuses;
        }


    }
}
