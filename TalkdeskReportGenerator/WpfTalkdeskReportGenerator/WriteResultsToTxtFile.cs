using System;
using System.Collections.Generic;

namespace WpfTalkdeskReportGenerator
{
    internal interface IWriteResults
    {
        void WriteResults(string folderPath, List<AgentStatuses> consolidatedAgentStatuses, DateTime workbookDate);
    }

    internal class WriteResultsToTxtFile : IWriteResults
    {
        public void WriteResults(string folderPath, List<AgentStatuses> consolidatedAgentStatuses, DateTime workbookDate)
        {

            string date = workbookDate.ToShortDateString().Replace(@"/", "-");

            using (System.IO.StreamWriter file = new System.IO.StreamWriter($"{folderPath}TalkDeskReport-{date}.txt"))
            {
                foreach (AgentStatuses aStatus in consolidatedAgentStatuses)
                {
                    file.WriteLine(aStatus.AgentName);
                    string outputString = "";

                    foreach (Status status in aStatus.Statuses)
                    {
                        outputString += $" - {status.StatusLabel} : {((decimal)status.StatusTime/60).ToString("0.##")} {Environment.NewLine}";
                    }
                    file.WriteLine(outputString);
                }

            }
        }
    }
}
