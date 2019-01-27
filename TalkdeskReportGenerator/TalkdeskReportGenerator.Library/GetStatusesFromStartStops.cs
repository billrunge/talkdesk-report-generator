﻿using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TalkdeskReportGenerator.Library
{
    public interface IGetStatusesFromStartStops
    {
        Task<List<AgentStatuses>> GetAgentStatusesListAsync(IGetStatuses getStatuses, List<AgentStartStops> agentStartStops, DateTime day, TimeZoneInfo excelTimeZone);
    }

    public class GetStatusesFromStartStops : IGetStatusesFromStartStops
    {
        public async Task<List<AgentStatuses>> GetAgentStatusesListAsync(IGetStatuses getStatuses, List<AgentStartStops> agentStartStops, DateTime day, TimeZoneInfo excelTimeZone)
        {
            List<AgentStatuses> agentStatusesList = new List<AgentStatuses>();
            int utcOffset = Math.Abs(excelTimeZone.GetUtcOffset(day).Hours);

            List<Task<AgentStatuses>> tasks = new List<Task<AgentStatuses>>();

            foreach (AgentStartStops agentStartStop in agentStartStops)
            {
                tasks.Add(GetAgentStatusesByStartStopAsync(getStatuses, agentStartStop, day, utcOffset));
            }

            var results = await Task.WhenAll(tasks);

            foreach(var agentStatus in results)
            {
                agentStatusesList.Add(agentStatus);
            }
            return agentStatusesList;
        }
                
        private async Task<AgentStatuses> GetAgentStatusesByStartStopAsync(IGetStatuses getStatuses, AgentStartStops agentStartStop, DateTime day, int utcOffset)
        {
            string userId = await getStatuses.GetUserIdFromNameAsync(agentStartStop.AgentName);

            AgentStatuses agentStatuses = new AgentStatuses()
            {
                AgentName = agentStartStop.AgentName
            };

            List<Task<List<Status>>> tasks = new List<Task<List<Status>>>();

            foreach (StartStop startStop in agentStartStop.StartStopList)
            {
                DateTime startTime = day.Add(startStop.Start);
                DateTime stopTime = day.Add(startStop.Stop);

                tasks.Add(getStatuses.GetStatusesListAsync(userId, startTime, stopTime, utcOffset));
            }

             List<Status>[] results = await Task.WhenAll(tasks);

            foreach (List<Status> statuses in results)
            {
                agentStatuses.Statuses.AddRange(statuses);
            }    

            return agentStatuses;
        }
    }
}