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
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private string _inputExcelPath;
        private string _outputPath;
        private string _status;
        private string _selectedTeam;
        private bool _getTeamNamesRan;
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
        public bool GetTeamNamesRan
        {
            get
            {
                return _getTeamNamesRan;
            }
            set
            {
                _getTeamNamesRan = value;
                NotifyOfPropertyChange(() => CanGetTeamNames);
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



        public bool CanGetTeamNames => (string.IsNullOrWhiteSpace(InputExcelPath) || string.IsNullOrWhiteSpace(OutputPath) || TeamNames.Count > 0 || GetTeamNamesRan) ? false : true;
        public bool CanSetExcelPath => (string.IsNullOrWhiteSpace(InputExcelPath)) ? true : false;
        public bool CanSetOutputPath => (string.IsNullOrWhiteSpace(OutputPath)) ? true : false;
        public bool CanSetTeamName => (TeamNames.Count > 0) ? true : false;
        public bool CanGenerateReport => (string.IsNullOrWhiteSpace(SelectedTeam)) ? false : true;

        public ShellViewModel()
        {
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.ShellViewModelStarting - Application");
            }
            TeamNames = new List<string>();
            GetTeamNamesRan = false;
        }

        public void SetExcelPath()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetExcelPath - Opening file dialog");
            }

            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Title = "Open Schedule Excel",
                Filter = "Excel Files|*.xlsx; *.xlsb",
                InitialDirectory = @"C:\"
            };

            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetExcelPath - Checking to see if file name was actually set");
            }

            if (fileDialog.ShowDialog() == true)
            {
                InputExcelPath = fileDialog.FileName.ToString();
                _log.Info($"ShellViewModel.SetExcelPath - InputExcelPath set to { InputExcelPath }");
            }

        }

        public async Task Clear()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.Clear - Clearing InputExcelPath, OutputPath, SelectedTeam, TeamNames, GetTeamNamesRan, and Status properties");
            }
            InputExcelPath = null;
            OutputPath = null;
            SelectedTeam = null;
            TeamNames = new List<string>();
            GetTeamNamesRan = false;
            Status = "";

            if (!string.IsNullOrWhiteSpace(TempExcelPath))
            {
                if (_log.IsInfoEnabled)
                {
                    _log.Info($"ShellViewModel.Clear - Deleting the temporary file: { TempExcelPath }");
                }
                ExcelReader excelReader = new ExcelReader(_log);
                await excelReader.DeleteExcelAsync(TempExcelPath);
                TempExcelPath = null;
            }
        }

        public void SetOutputPath()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetOutputPath - Opening Folder Browser Dialog");
            }
            System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog()
            {
                Description = "ShellViewModel.SetOutputPath - Select Output Folder",
                ShowNewFolderButton = true
            };
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputPath = folderBrowser.SelectedPath + @"\";
                if (_log.IsInfoEnabled)
                {
                    _log.Info($"ShellViewModel.SetOutputPath - OuputPath set to { OutputPath }");
                }
            }
        }

        public async Task GetTeamNamesAsync()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.GetTeamNamesAsync - Setting GetTeamNamesRan to true");
            }
            await Task.Run(() => GetTeamNamesRan = true);

            Status = "Generating a working copy Excel...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - " + Status);
            }

            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.GetTeamNamesAsync - Generating a new ExcelReader");
            }

            ExcelReader excelReader = await Task.Run(() => new ExcelReader(_log));


            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - Generating temporary, lightweight Excel");
            }
            TempExcelPath = await excelReader.CreateLightweightExcelAsync(InputExcelPath);
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - Generating temporary, lightweight Excel complete");
            }

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ShellViewModel.GetTeamNamesAsync - TempExcelPath = { TempExcelPath }");
            }

            Status = "Getting team names from Excel...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - " + Status);
            }

            TeamNames = await excelReader.GetTeamNamesAsync(TempExcelPath);
            Status = "Please select team name.";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - " + Status);
            }
        }

        public async Task GenerateReportAsync()
        {
            IDatabase db = new Database(_log);
            IGetStatuses getStatuses = new GetStatuses(db, _log);
            ExcelReader excelReader = new ExcelReader(_log);

            Status = "Reading Excel...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }
            List<AgentStartStops> startStopList = await excelReader.GetAgentStartStopListAsync(TempExcelPath, SelectedTeam);

            IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();
            DateTime day = excelReader.WorksheetDay;

            Status = "Retrieving agent statuses...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }
            List<AgentStatuses> agentStatuses = await getStatusesFromStartStops.GetAgentStatusesListAsync(getStatuses, startStopList, day);
            IConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();

            List<AgentStatuses> consolidatedAgentStatuses = await Task.Run(() => consolidateStatuses.Consolidate(agentStatuses));
            Status = "Writing results to file...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }
            IWriteResults writeResults = new WriteResultsToTxtFile();

            await Task.Run(() => writeResults.WriteResults(OutputPath, consolidatedAgentStatuses, SelectedTeam, excelReader.WorksheetDay));

            Status = "Job complete!";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }

        }

        public void Exit()
        {
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.Exit - Exiting application from Exit() function");
                Application.Current.Shutdown();
            }
        }

        public async Task OnClose(CancelEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(TempExcelPath))
            {
                ExcelReader excelReader = new ExcelReader(_log);
                await excelReader.DeleteExcelAsync(TempExcelPath);
            }
        }

        public void About()
        {
            MessageBox.Show("2018 Relativity ODA LLC.");
        }
    }
}
