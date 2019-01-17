using Caliburn.Micro;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
            get
            {
                return _selectedTeam;
            }
            set
            {
                _selectedTeam = value;
                NotifyOfPropertyChange(() => SelectedTeam);
            }
        }
        public List<string> TeamNames
        {
            get => _teamNames;
            set
            {
                _teamNames = value;
                NotifyOfPropertyChange(() => CanGenerateReport);
                NotifyOfPropertyChange(() => TeamNames);
            }
        }

        public bool CanGetTeamNames => (string.IsNullOrWhiteSpace(ExcelPath) || string.IsNullOrWhiteSpace(OutputPath)) ? false : true;
        public bool CanSetExcelPath => (string.IsNullOrWhiteSpace(ExcelPath)) ? true : false;
        public bool CanSetOutputPath => (string.IsNullOrWhiteSpace(OutputPath)) ? true : false;
        public bool CanGenerateReport => (TeamNames.Count > 0) ? true : false;

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

        public void GetTeamNames()
        {
            ExcelReader excelReader = new ExcelReader();
            FilePath = excelReader.CreateLightweightExcel(ExcelPath);
            TeamNames = excelReader.GetTeamNames(FilePath);
        }

        public void GenerateReport()
        {
            IDatabase db = new Database();
            IGetStatuses getStatuses = new GetStatuses(db);
            ExcelReader excelReader = new ExcelReader();
            List<AgentStartStops> startStopList = excelReader.GetAgentStartStopList(FilePath, SelectedTeam);
            excelReader.DeleteExcel(FilePath);
            IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();

            DateTime day = excelReader.WorkbookDay;

            List<AgentStatuses> agentStatuses = getStatusesFromStartStops.GetAgentStatusesList(getStatuses, startStopList, day);

            IConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();
            List<AgentStatuses> consolidatedAgentStatuses = consolidateStatuses.Consolidate(agentStatuses);

            IWriteResults writeResults = new WriteResultsToTxtFile();        
            


            writeResults.WriteResults(OutputPath, consolidatedAgentStatuses, excelReader.WorkbookDay);

            MessageBox.Show("Job Complete!");

        }

        public void Exit()
        {
            Application.Current.Shutdown();
        }

        public void About()
        {
            MessageBox.Show("2018 Relativity ODA LLC.");
        }
    }
}
