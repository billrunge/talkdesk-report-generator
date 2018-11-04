using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ConsoleTalkdeskReportGenerator
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                IDatabase db = new Database();
                IGetStatuses getStatuses = new GetStatuses(db);
                IGetAgentTimes getAgentTimes = new GetAgentTimesFromExcel();

                string filePath = "";

                OpenFileDialog fileDialog = new OpenFileDialog
                {
                    Title = "Open Schedule Excel",
                    Filter = "Excel Files|*.xlsx",
                    InitialDirectory = @"C:\"
                };

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = fileDialog.FileName.ToString();
                }

                List<AgentStartStop> startStopList = getAgentTimes.GetAgentStartStopList(filePath);
                GetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops(getStatuses);

                DateTime monday = getAgentTimes.WorkbookMonday;

                List<AgentStatuses> agentStatuses = getStatusesFromStartStops.GetAgentStatusesList(startStopList, monday);

                ConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();

                List<AgentStatuses> consolidatedAgentStatuses = consolidateStatuses.Consolidate(agentStatuses);

                foreach(var aStatus in consolidatedAgentStatuses)
                {
                    Console.WriteLine(aStatus.AgentName);
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
                    Console.WriteLine(mondayString);
                    Console.WriteLine(tuesdayString);
                    Console.WriteLine(wednesdayString);
                    Console.WriteLine(thursdayString);
                    Console.WriteLine(fridayString);


                }
                Console.ReadLine();

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Console.ReadLine();
            }
        }
    }
}
