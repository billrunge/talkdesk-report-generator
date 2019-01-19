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
        private string _excelPath;
        private string _excelPathStatus;
        private string _outputPath;
        private string _outputPathStatus;
        private string _status;
        private string _selectedTeam;
        private List<string> _teamNames;

        public string ExcelPath
        {
            get => _excelPath;
            set
            {
                _excelPath = value;
                SetExcelPathStatus();
                NotifyOfPropertyChange(() => ExcelPath);
                NotifyOfPropertyChange(() => CanGetTeamNames);
                NotifyOfPropertyChange(() => CanSetExcelPath);
            }
        }
        public string ExcelPathStatus
        {
            get => _excelPathStatus;
            set
            {
                _excelPathStatus = value;
                NotifyOfPropertyChange(() => ExcelPathStatus);
            }
        }
        public string OutputPath
        {
            get => _outputPath;
            set
            {
                _outputPath = value;
                SetOutputPathStatus();
                NotifyOfPropertyChange(() => OutputPath);
                NotifyOfPropertyChange(() => CanSetOutputPath);
            }

        }
        public string OutputPathStatus
        {
            get => _outputPathStatus;
            set
            {
                _outputPathStatus = value;
                NotifyOfPropertyChange(() => OutputPathStatus);
                NotifyOfPropertyChange(() => CanGetTeamNames);
            }
        }
        public string FilePath { get; set; }
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
            }
        }

        public bool CanGetTeamNames => (string.IsNullOrWhiteSpace(ExcelPath) || string.IsNullOrWhiteSpace(OutputPath)) ? false : true;
        public bool CanSetExcelPath => (string.IsNullOrWhiteSpace(ExcelPath)) ? true : false;
        public bool CanSetOutputPath => (string.IsNullOrWhiteSpace(OutputPath)) ? true : false;
        public bool CanSetTeamName => (TeamNames.Count > 0) ? true : false;
        public bool CanGenerateReport => (string.IsNullOrWhiteSpace(SelectedTeam)) ? false : true;

        public ShellViewModel()
        {
            TeamNames = new List<string>();
            SetExcelPathStatus();
            SetOutputPathStatus();
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
                ExcelPath = fileDialog.FileName.ToString();
            }

        }

        public void SetExcelPathStatus()
        {
            if (ExcelPath == null)
            {
                ExcelPathStatus = "✖";
            }
            else
            {
                ExcelPathStatus = "✔";
            }
        }

        public void ClearExcelPath()
        {
            ExcelPath = null;
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

        public void SetOutputPathStatus()
        {
            if (OutputPath == null)
            {
                OutputPathStatus = "✖";
            }
            else
            {
                OutputPathStatus = "✔";
            }
        }

        public void ClearOutputPath()
        {
            OutputPath = null;
        }

        public async Task GetTeamNamesAsync()
        {
            ExcelReader excelReader = new ExcelReader();
            Status = "Generating a working copy Excel...";
            FilePath = await Task.Run(() => excelReader.CreateLightweightExcel(ExcelPath));
            Status = "Getting team names from Excel...";
            TeamNames = await Task.Run(() => excelReader.GetTeamNames(FilePath));
            Status = "Please select team name.";
        }

        public async Task GenerateReportAsync()
        {
            IDatabase db = new Database();
            IGetStatuses getStatuses = new GetStatuses(db);
            ExcelReader excelReader = new ExcelReader();

            Status = "Reading Excel...";
            List<AgentStartStops> startStopList = await Task.Run(() => excelReader.GetAgentStartStopList(FilePath, SelectedTeam));

            IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();
            DateTime day = excelReader.WorkbookDay;

            Status = "Retrieving agent statuses...";
            List<AgentStatuses> agentStatuses = await Task.Run(() => getStatusesFromStartStops.GetAgentStatusesList(getStatuses, startStopList, day));
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

        public void OnClose(CancelEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(FilePath))
            {
                ExcelReader excelReader = new ExcelReader();
                excelReader.DeleteExcel(FilePath);
            }
        }

        public void About()
        {
            MessageBox.Show("2018 Relativity ODA LLC.");
        }
    }
}
