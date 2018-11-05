using System;
using System.Collections.Generic;

namespace ConsoleTalkdeskReportGenerator
{
    class WriteResultsToTxtFile
    {
        public string OutputFileName { get; set; } = "TalkdeskReport.txt";

        public void WriteResults(string folderPath, List<AgentStatuses> consolidatedAgentStatuses)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(folderPath + OutputFileName))
            {
                foreach (var aStatus in consolidatedAgentStatuses)
                {
                    file.WriteLine(aStatus.AgentName);
                    string mondayString = $"- Monday {Environment.NewLine}";
                    string tuesdayString = $"- Tuesday {Environment.NewLine}";
                    string wednesdayString = $"- Wednesday {Environment.NewLine}";
                    string thursdayString = $"- Thursday {Environment.NewLine}";
                    string fridayString = $"- Friday {Environment.NewLine}";

                    foreach (var status in aStatus.Statuses)
                    {
                        switch (status.DayName)
                        {
                            case "Monday":
                                mondayString += $" - {status.StatusLabel} : {status.StatusTime / 60} {Environment.NewLine}";
                                break;
                            case "Tuesday":
                                tuesdayString += $" - {status.StatusLabel} : {status.StatusTime / 60} {Environment.NewLine}";
                                break;
                            case "Wednesday":
                                wednesdayString += $" - {status.StatusLabel} : {status.StatusTime / 60} {Environment.NewLine}";
                                break;
                            case "Thursday":
                                thursdayString += $" - {status.StatusLabel} : {status.StatusTime / 60} {Environment.NewLine}";
                                break;
                            case "Friday":
                                fridayString += $" - {status.StatusLabel} : {status.StatusTime / 60} {Environment.NewLine}";
                                break;
                            default:
                                break;
                        }

                    }
                    file.WriteLine(mondayString);
                    file.WriteLine(tuesdayString);
                    file.WriteLine(wednesdayString);
                    file.WriteLine(thursdayString);
                    file.WriteLine(fridayString);
                }

            }
        }
    }
}
