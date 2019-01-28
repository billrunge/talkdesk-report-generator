using Caliburn.Micro;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using TalkdeskReportGenerator.Library;

namespace TalkdeskReportGenerator.ViewModels
{
    public class ReportsViewModel : Screen
    {
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private string _inputExcelPath;
        private string _outputPath;
        private string _status;
        private string _selectedName;
        private bool _getNamesRan;
        private List<string> _names;

        public string InputExcelPath
        {
            get => _inputExcelPath;
            set
            {
                _inputExcelPath = value;
                NotifyOfPropertyChange(() => InputExcelPath);
                NotifyOfPropertyChange(() => CanGetNames);
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
                NotifyOfPropertyChange(() => CanGetNames);
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
        public string SelectedName
        {
            get => _selectedName;
            set
            {
                _selectedName = value;
                NotifyOfPropertyChange(() => SelectedName);
                NotifyOfPropertyChange(() => CanGenerateReport);
                Status = "Click Generate Report";
            }
        }
        public bool GetNamesRan
        {
            get
            {
                return _getNamesRan;
            }
            set
            {
                _getNamesRan = value;
                NotifyOfPropertyChange(() => CanGetNames);
            }
        }
        public List<string> Names
        {
            get => _names;
            set
            {
                _names = value;
                NotifyOfPropertyChange(() => Names);
                NotifyOfPropertyChange(() => CanSetName);
                NotifyOfPropertyChange(() => CanGetNames);
            }
        }
        public string SelectNameText {
            get
            {
                return _selectedName;
            }
            set
            {
                _selectedName = value;
                NotifyOfPropertyChange(() => SelectNameText);
            }
        }

        public bool CanGetNames => (string.IsNullOrWhiteSpace(InputExcelPath) || string.IsNullOrWhiteSpace(OutputPath) || Names.Count > 0 || GetNamesRan) ? false : true;
        public bool CanSetExcelPath => (string.IsNullOrWhiteSpace(InputExcelPath)) ? true : false;
        public bool CanSetOutputPath => (string.IsNullOrWhiteSpace(OutputPath)) ? true : false;
        public bool CanSetName => (Names.Count > 0) ? true : false;
        public bool CanGenerateReport => (string.IsNullOrWhiteSpace(SelectedName)) ? false : true;

        public ReportsViewModel()
        {
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.ShellViewModelStarting - Application");
            }
            Names = new List<string>();
            GetNamesRan = false;
            Properties.Settings.Default.TemporaryExcelPaths = new List<string>();
        }

        public void SetExcelPath()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetExcelPath - Opening file dialog");
            }

            string initDirectory;

            if (string.IsNullOrEmpty(Properties.Settings.Default.InputDirectory))
            {
                initDirectory = @"C:\";
            }
            else
            {
                initDirectory = Properties.Settings.Default.InputDirectory;
            }
            

            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Title = "Open Schedule Excel",
                Filter = "Excel Files|*.xlsx; *.xlsb",
                InitialDirectory = initDirectory
            };

            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetExcelPath - Checking to see if file name was actually set");
            }

            if (fileDialog.ShowDialog() == true)
            {
                InputExcelPath = fileDialog.FileName.ToString();
                Properties.Settings.Default.InputDirectory = Path.GetDirectoryName(InputExcelPath);
                _log.Info($"ShellViewModel.SetExcelPath - InputExcelPath set to { InputExcelPath }");
                Properties.Settings.Default.Save();
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
            SelectedName = null;
            Names = new List<string>();
            GetNamesRan = false;
            Status = "";
            TempExcelPath = null;

            foreach (string tempExcelPath in Properties.Settings.Default.TemporaryExcelPaths)
            {
                if (!string.IsNullOrWhiteSpace(tempExcelPath))
                {
                    if (_log.IsInfoEnabled)
                    {
                        _log.Info($"ReportsViewModel.Clear - Deleting the temporary file: { tempExcelPath }");
                    }

                    ExcelReader excelReader = new ExcelReader(_log);
                    await excelReader.DeleteExcelAsync(tempExcelPath);
                }
            }

            Properties.Settings.Default.TemporaryExcelPaths = new List<string>();
            Properties.Settings.Default.Save();

        }

        public void SetOutputPath()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetOutputPath - Opening Folder Browser Dialog");
            }
            string rootFolder;

            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.OutputDirectory)){
                rootFolder = @"C:\";
            }
            else
            {
                rootFolder = Properties.Settings.Default.OutputDirectory;
            }

            System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog()
            {
                Description = "Select Output Folder",
                ShowNewFolderButton = true,
                SelectedPath = rootFolder
            };

            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputPath = folderBrowser.SelectedPath + @"\";
                Properties.Settings.Default.OutputDirectory = Path.GetDirectoryName(OutputPath);
                Properties.Settings.Default.Save();
                if (_log.IsInfoEnabled)
                {
                    _log.Info($"ShellViewModel.SetOutputPath - OuputPath set to { OutputPath }");
                }
            }
        }

        public async Task GetNamesAsync()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.GetTeamNamesAsync - Setting GetNamesRan to true");
            }
            await Task.Run(() => GetNamesRan = true);

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

            Properties.Settings.Default.TemporaryExcelPaths.Add(TempExcelPath);

            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - Generating temporary, lightweight Excel complete");
            }

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ShellViewModel.GetTeamNamesAsync - TempExcelPath = { TempExcelPath }");
            }

            Status = "Getting names from Excel...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ReportsViewmodel.GetTeamNamesAsync - " + Status);
            }

            ExcelCell groupByCell = new ExcelCell
            {
                Column = Properties.Settings.Default.GroupByNameColumn,
                Row = Properties.Settings.Default.GroupByNameRow
            };

            string groupByName = await excelReader.GetGroupByNameAsync(TempExcelPath, groupByCell);

            Names = await excelReader.GetNamesAsync(TempExcelPath, groupByCell);

            SelectNameText = $"Select { groupByName }";
            Status = $"Please select { groupByName } name";
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

            ExcelCell groupByNameCell = new ExcelCell()
            {
                Column = Properties.Settings.Default.GroupByNameColumn,
                Row = Properties.Settings.Default.GroupByNameRow
            };

            ExcelCell phoneColorKeyCell = new ExcelCell()
            {
                Column = Properties.Settings.Default.PhoneColorKeyColumn,
                Row = Properties.Settings.Default.PhoneColorKeyRow
            };

            string agentNameColumn = Properties.Settings.Default.AgentNameColumn;
            string twelveAmColumn = Properties.Settings.Default.TwelveAmColumn;

            List<AgentStartStops> startStopList = await excelReader.GetAgentStartStopListAsync(TempExcelPath, SelectedName, agentNameColumn, twelveAmColumn, groupByNameCell, phoneColorKeyCell);

            IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();
            DateTime day = excelReader.WorksheetDay;

            Status = "Retrieving agent statuses...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }

            TimeZoneInfo excelTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Properties.Settings.Default.TimeZoneId);

            List<AgentStatuses> agentStatuses = await getStatusesFromStartStops.GetAgentStatusesListAsync(getStatuses, startStopList, day, excelTimeZone);
            IConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();

            List<AgentStatuses> consolidatedAgentStatuses = await Task.Run(() => consolidateStatuses.Consolidate(agentStatuses));
            Status = "Writing results to file...";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }
            IWriteResults writeResults = new WriteResultsToExcelFile();

            await Task.Run(() => writeResults.WriteResults(OutputPath, consolidatedAgentStatuses, SelectedName, excelReader.WorksheetDay));

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
            }
            Application.Current.Shutdown();
        }

        public void Settings()
        {
            ActivateWindow.ViewSettings();
        }

        public void About()
        {
            MessageBox.Show("2019 Relativity ODA LLC.");
        }
    }
}

