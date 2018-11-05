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
                string test = "";

                IDatabase db = new Database();
                IGetStatuses getStatuses = new GetStatuses(db);
                IGetAgentTimes getAgentTimes = new GetAgentTimesFromExcel();

                string filePath = "";

                Console.WriteLine("Welcome!");
                Console.WriteLine("Please select the Weekly Schedule Excel file");

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

                Console.WriteLine("Working...");

                List<AgentStartStops> startStopList = getAgentTimes.GetAgentStartStopList(filePath);
                IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();

                DateTime monday = getAgentTimes.WorkbookMonday;

                List<AgentStatuses> agentStatuses = getStatusesFromStartStops.GetAgentStatusesList(getStatuses, startStopList, monday);

                ConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();
                List<AgentStatuses> consolidatedAgentStatuses = consolidateStatuses.Consolidate(agentStatuses);

                WriteResultsToTxtFile writeResults = new WriteResultsToTxtFile();

                Console.WriteLine("Please select output directory");

                string folderPath = "";
                FolderBrowserDialog folderBrowser = new FolderBrowserDialog()
                {
                    Description = "Select Output Folder",
                    ShowNewFolderButton = true                   

                };
                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    folderPath = folderBrowser.SelectedPath + @"\";
                }

                writeResults.WriteResults(folderPath, consolidatedAgentStatuses);

                Console.WriteLine("Complete");
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
