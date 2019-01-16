using System;
using System.Collections.Generic;

namespace WpfTalkdeskReportGenerator
{
    internal interface IWriteResults
    {
        void WriteResults(string folderPath, List<AgentStatuses> consolidatedAgentStatuses);
    }

    internal class WriteResultsToTxtFile : IWriteResults
    {
        public string OutputFileName { get; set; } = "TalkdeskReport.txt";

        public void WriteResults(string folderPath, List<AgentStatuses> consolidatedAgentStatuses)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(folderPath + OutputFileName))
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
