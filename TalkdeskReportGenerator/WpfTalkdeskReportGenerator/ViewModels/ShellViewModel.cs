using Caliburn.Micro;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;

namespace WpfTalkdeskReportGenerator.ViewModels
{
    public class ShellViewModel : Screen
    {
        private string _inputExcelPath;
        private string _outputPath;
        private string _status;
        private string _selectedTeam;
        private List<string> _teamNames;

        public string InputExcelPath
        {
            get => _inputExcelPath;
            set
            {
                _inputExcelPath = value;
                NotifyOfPropertyChange(() => InputExcelPath);
                NotifyOfPropertyChange(() => CanGetTeamNames);
                NotifyOfPropertyChange(() => CanSetExcelPath);
            }
        }

        public string OutputPath
        {
            get => _outputPath;
            set
            {
                _outputPath = value;
                NotifyOfPropertyChange(() => OutputPath);
                NotifyOfPropertyChange(() => CanSetOutputPath);
                NotifyOfPropertyChange(() => CanGetTeamNames);
            }

        }
        public string TempExcelPath { get; set; }
        public string Status
        {
            get => _status;
            set
            {
                _status = value;
                NotifyOfPropertyChange(() => Status);
            }
        }
        public string SelectedTeam
        {
            get => _selectedTeam;
            set
            {
                _selectedTeam = value;
                NotifyOfPropertyChange(() => SelectedTeam);
                NotifyOfPropertyChange(() => CanGenerateReport);
            }
        }
        public List<string> TeamNames
        {
            get => _teamNames;
            set
            {
                _teamNames = value;
                NotifyOfPropertyChange(() => TeamNames);
                NotifyOfPropertyChange(() => CanSetTeamName);
                NotifyOfPropertyChange(() => CanGetTeamNames);
            }
        }

        public bool CanGetTeamNames => (string.IsNullOrWhiteSpace(InputExcelPath) || string.IsNullOrWhiteSpace(OutputPath) || TeamNames.Count > 0) ? false : true;
        public bool CanSetExcelPath => (string.IsNullOrWhiteSpace(InputExcelPath)) ? true : false;
        public bool CanSetOutputPath => (string.IsNullOrWhiteSpace(OutputPath)) ? true : false;
        public bool CanSetTeamName => (TeamNames.Count > 0) ? true : false;
        public bool CanGenerateReport => (string.IsNullOrWhiteSpace(SelectedTeam)) ? false : true;

        public ShellViewModel()
        {
            TeamNames = new List<string>();
        }

        public void SetExcelPath()
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Title = "Open Schedule Excel",
                Filter = "Excel Files|*.xlsx; *.xlsb",
                InitialDirectory = @"C:\"
            };

            if (fileDialog.ShowDialog() == true)
            {
                InputExcelPath = fileDialog.FileName.ToString();
            }

        }

        public async Task Clear()
        {
            InputExcelPath = null;
            OutputPath = null;
            SelectedTeam = null;
            TeamNames = new List<string>();

            if (!string.IsNullOrWhiteSpace(TempExcelPath))
            {
                ExcelReader excelReader = new ExcelReader();
                await excelReader.DeleteExcelAsync(TempExcelPath);
                TempExcelPath = null;
            }

            Status = "";
        }

        public void SetOutputPath()
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog()
            {
                Description = "Select Output Folder",
                ShowNewFolderButton = true
            };
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputPath = folderBrowser.SelectedPath + @"\";
            }
        }

        public async Task GetTeamNamesAsync()
        {
            ExcelReader excelReader = new ExcelReader();
            Status = "Generating a working copy Excel...";
            TempExcelPath = await excelReader.CreateLightweightExcelAsync(InputExcelPath);
            Status = "Getting team names from Excel...";
            TeamNames = await excelReader.GetTeamNamesAsync(TempExcelPath);
            Status = "Please select team name.";
        }

        public async Task GenerateReportAsync()
        {
            IDatabase db = new Database();
            IGetStatuses getStatuses = new GetStatuses(db);
            ExcelReader excelReader = new ExcelReader();

            Status = "Reading Excel...";
            List<AgentStartStops> startStopList = await excelReader.GetAgentStartStopListAsync(TempExcelPath, SelectedTeam);

            IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();
            DateTime day = excelReader.WorkbookDay;

            Status = "Retrieving agent statuses...";
            List<AgentStatuses> agentStatuses = await getStatusesFromStartStops.GetAgentStatusesListAsync(getStatuses, startStopList, day);
            IConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();

            List<AgentStatuses> consolidatedAgentStatuses = await Task.Run(() => consolidateStatuses.Consolidate(agentStatuses));
            Status = "Writing results to file...";
            IWriteResults writeResults = new WriteResultsToTxtFile();

            await Task.Run(() => writeResults.WriteResults(OutputPath, consolidatedAgentStatuses, SelectedTeam, excelReader.WorkbookDay));

            Status = "Job complete!";

        }

        public void Exit()
        {
            Application.Current.Shutdown();
        }

        public async Task OnClose(CancelEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(TempExcelPath))
            {
                ExcelReader excelReader = new ExcelReader();
                await excelReader.DeleteExcelAsync(TempExcelPath);
            }
        }

        public void About()
        {
            MessageBox.Show("2018 Relativity ODA LLC.");
        }
    }
}
