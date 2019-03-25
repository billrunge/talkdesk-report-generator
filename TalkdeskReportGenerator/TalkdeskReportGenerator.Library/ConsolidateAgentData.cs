using System.Collections.Generic;
using System.Linq;

namespace TalkdeskReportGenerator.Library
{
    public interface IConsolidateAgentData
    {
        List<AgentData> Consolidate(List<AgentData> agentDataList);
    }

    public class ConsolidateAgentData : IConsolidateAgentData
    {
        public List<AgentData> Consolidate(List<AgentData> agentDataList)
        {
            List<AgentData> outputList = new List<AgentData>();


            foreach (AgentData agentData in agentDataList)
            {
                List<Status> newStatuses = new List<Status>();

                newStatuses = agentData.Statuses
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

                List<Call> newCalls = new List<Call>();


                newCalls = agentData.Calls
                    .GroupBy(a => new
                    {
                        a.Type
                    })
                    .Select(a => new Call
                    {
                        Type = a.Key.Type,
                        Count = a.Sum(ag => ag.Count)

                    }).OrderByDescending(a => a.Type)
                    .ToList();

                AgentData newAgentData = new AgentData
                {
                    AgentName = agentData.AgentName,
                    Statuses = newStatuses,
                    Calls = newCalls
                };

                outputList.Add(newAgentData);

            }
            return outputList;
        }


    }
}
