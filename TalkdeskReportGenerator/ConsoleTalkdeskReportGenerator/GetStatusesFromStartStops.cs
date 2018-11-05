using System;
using System.Collections.Generic;
using System.Configuration;

namespace ConsoleTalkdeskReportGenerator
{
    interface IGetStatusesFromStartStops
    {
        List<AgentStatuses> GetAgentStatusesList(IGetStatuses getStatuses, List<AgentStartStops> agentStartStops, DateTime monday);
    }
    class GetStatusesFromStartStops : IGetStatusesFromStartStops
    {
        private readonly TimeZoneInfo _excelTimeZone;

        public GetStatusesFromStartStops()
        {

            if (ConfigurationManager.AppSettings["ExcelTimeZone"] != null)
            {
                _excelTimeZone = TimeZoneInfo.FindSystemTimeZoneById(ConfigurationManager.AppSettings["ExcelTimeZone"]);
            }
            else
            {
                throw new ConfigurationErrorsException("Unable to retrieve ExcelTimeZone key from App.config file");
            }
        }

        public List<AgentStatuses> GetAgentStatusesList(IGetStatuses getStatuses, List<AgentStartStops> agentStartStops, DateTime monday)
        {
            List<AgentStatuses> agentStatusesList = new List<AgentStatuses>();
            monday = monday.Date;
            if (monday.DayOfWeek == DayOfWeek.Monday)
            {         
                 
                int utcOffset = Math.Abs(_excelTimeZone.GetUtcOffset(monday).Hours);

                foreach (AgentStartStops agentStartStop in agentStartStops)
                {
                    string userId = getStatuses.GetUserIdFromName(agentStartStop.AgentName);

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

                            List<Status> agentStatus = getStatuses.GetStatusesList(userId, startTime, stopTime, utcOffset);
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
