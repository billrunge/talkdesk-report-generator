using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TalkdeskReportGenerator.Library
{
    public interface IGetAgentDataFromStartStops
    {
        Task<List<AgentData>> GetAgentDataListAsync(IGetStatuses getStatuses, IGetCalls getCalls, List<AgentStartStops> agentStartStops, DateTime day, TimeZoneInfo excelTimeZone, List<AgentMapping> mappings);
    }

    public class GetAgentDataFromStartStops : IGetAgentDataFromStartStops
    {
        public async Task<List<AgentData>> GetAgentDataListAsync(IGetStatuses getStatuses, IGetCalls getCalls, List<AgentStartStops> agentStartStops, DateTime day, TimeZoneInfo excelTimeZone, List<AgentMapping> mappings)
        {
            List<AgentData> agentDataList = new List<AgentData>();
            int utcOffset = Math.Abs(excelTimeZone.GetUtcOffset(day).Hours);

            List<Task<AgentData>> tasks = new List<Task<AgentData>>();

            foreach (AgentStartStops agentStartStop in agentStartStops)
            {
                tasks.Add(GetAgentDataByStartStopAsync(getStatuses, getCalls, agentStartStop, day, utcOffset, mappings));
            }

            AgentData[] results = await Task.WhenAll(tasks);

            foreach (AgentData agentData in results)
            {
                agentDataList.Add(agentData);
            }
            return agentDataList;
        }

        private async Task<AgentData> GetAgentDataByStartStopAsync(IGetStatuses getStatuses, IGetCalls getCalls, AgentStartStops agentStartStop, DateTime day, int utcOffset, List<AgentMapping> mappings)
        {
            List<Task<List<Call>>> callTasks = new List<Task<List<Call>>>();
            List<Task<List<Status>>> statusTasks = new List<Task<List<Status>>>();

            string agentName = agentStartStop.AgentName;

            foreach (AgentMapping mapping in mappings)
            {
                agentName = (agentName == mapping.ExcelAgentName) ? mapping.TalkdeskAgentName : agentName;
            }

            foreach (StartStop startStop in agentStartStop.StartStopList)
            {
                DateTime startTime = day.Add(startStop.Start);
                DateTime stopTime = day.Add(startStop.Stop);
                callTasks.Add(getCalls.GetCallListAsync(agentName, startTime, stopTime, utcOffset));
                statusTasks.Add(getStatuses.GetStatusesListAsync(await getStatuses.GetUserIdFromNameAsync(agentName), startTime, stopTime, utcOffset));
            }

            AgentData agentData = new AgentData()
            {
                AgentName = agentStartStop.AgentName
            };

            List<Call>[] callResults = await Task.WhenAll(callTasks);

            foreach (List<Call> calls in callResults)
            {
                agentData.Calls.AddRange(calls);
            }

            List<Status>[] statusResults = await Task.WhenAll(statusTasks);

            foreach (List<Status> statuses in statusResults)
            {
                agentData.Statuses.AddRange(statuses);
            }
            return agentData;
        }
    }
}
