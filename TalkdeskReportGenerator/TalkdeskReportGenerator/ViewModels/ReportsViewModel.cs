using Caliburn.Micro;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using TalkdeskReportGenerator.Library;

namespace TalkdeskReportGenerator.ViewModels
{
    public class ReportsViewModel : Screen
    {
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private List<string> _inputExcelPaths;
        private List<string> _names;
        private string _outputPath;
        private string _status;
        private string _selectedName;
        private bool _getNamesRan;
        private bool _reportRan;

        public List<string> InputExcelPaths
        {
            get => _inputExcelPaths;
            set
            {
                _inputExcelPaths = value;
                NotifyOfPropertyChange(() => InputExcelPaths);
                NotifyOfPropertyChange(() => CanGetNames);
                NotifyOfPropertyChange(() => CanSetExcelPath);
            }
        }
        public List<string> TempExcelPaths { get; set; }
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
        public string SelectNameText
        {
            get => _selectedName;
            set
            {
                _selectedName = value;
                NotifyOfPropertyChange(() => SelectNameText);
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
            get => _getNamesRan;
            set
            {
                _getNamesRan = value;
                NotifyOfPropertyChange(() => CanGetNames);
            }
        }
        public bool ReportRan
        { get
            {
                return _reportRan;
            }
            set
            {
                _reportRan = value;
                NotifyOfPropertyChange(() => CanGenerateReport);
                NotifyOfPropertyChange(() => CanSetName);
            }
        }

        public ReportsViewModel()
        {
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.ShellViewModelStarting - Application");
            }
            Names = new List<string>();
            InputExcelPaths = new List<string>();
            TempExcelPaths = new List<string>();
            GetNamesRan = false;
            ReportRan = false;
            Properties.Settings.Default.TemporaryExcelPaths = new List<string>();

        }

        public bool CanGetNames => (InputExcelPaths.Count < 1 || string.IsNullOrWhiteSpace(OutputPath) || Names.Count > 0 || GetNamesRan) ? false : true;
        public bool CanSetExcelPath => (InputExcelPaths.Count < 1) ? true : false;
        public bool CanSetOutputPath => (string.IsNullOrWhiteSpace(OutputPath)) ? true : false;
        public bool CanSetName => (Names.Count > 0 && !ReportRan) ? true : false;
        public bool CanGenerateReport => (string.IsNullOrWhiteSpace(SelectedName) || ReportRan) ? false : true;


        public void SetExcelPath()
        {
            try
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
                InitialDirectory = initDirectory,
                Multiselect = true
            };

            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetExcelPath - Checking to see if file name was actually set");
            }

                if (fileDialog.ShowDialog() == true)
                {
                    InputExcelPaths = fileDialog.FileNames.ToList();

                    Properties.Settings.Default.InputDirectory = Path.GetDirectoryName(InputExcelPaths[0]);

                    if (_log.IsDebugEnabled)
                    {
                        _log.Info($"ShellViewModel.SetExcelPath - ({ InputExcelPaths.Count.ToString() }) InputExcelPaths");
                        foreach (string inputExcelPath in InputExcelPaths)
                        {
                            _log.Info($"ShellViewModel.SetExcelPath - { inputExcelPath }");
                        }
                    }

                    Properties.Settings.Default.Save();
                }
            }
            catch(Exception e)
            {
                MessageBox.Show("An issue has occurred. Please check log for more details");
                _log.Error($@"An error has occcurred when setting input Excel(s) path {Environment.NewLine}
                                {e.Message} {Environment.NewLine}
                                {e.StackTrace} {Environment.NewLine}");
            }

        }

        public void SetOutputPath()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.SetOutputPath - Opening Folder Browser Dialog");
            }
            string rootFolder;

            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.OutputDirectory))
            {
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

            Status = $"Generating a working copy Excel...";
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

            string tempPath = await excelReader.CreateLightweightExcelAsync(InputExcelPaths[0]);
            TempExcelPaths.Add(tempPath);
            Properties.Settings.Default.TemporaryExcelPaths.Add(tempPath);

            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - Generating temporary, lightweight Excel complete");
            }

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ShellViewModel.GetTeamNamesAsync - TempExcelPaths[0] = { TempExcelPaths[0] }");
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

            string groupByName = await excelReader.GetGroupByNameAsync(TempExcelPaths[0], groupByCell);

            Names = await excelReader.GetNamesAsync(TempExcelPaths[0], groupByCell);
            Names.Add("All");

            SelectNameText = $"Select { groupByName }";
            Status = $"Please select { groupByName } name";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GetTeamNamesAsync - " + Status);
            }
        }

        public async Task GenerateReportAsync()
        {
            ReportRan = true;
            IDatabase db = new Database(_log);
            IGetStatuses getStatuses = new GetStatuses(db, _log);
            ExcelReader excelReader = new ExcelReader(_log);

            int inputExcelPathsCount = InputExcelPaths.Count;

            for (int i = 1; i < inputExcelPathsCount; i++)
            {
                Status = $"Generating working copy Excel { (i + 1).ToString() } of { inputExcelPathsCount.ToString() }...";
                string tempPath = await excelReader.CreateLightweightExcelAsync(InputExcelPaths[i]);
                TempExcelPaths.Add(tempPath);
                Properties.Settings.Default.TemporaryExcelPaths.Add(tempPath);
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

            int currentExcelCount = 1;
            foreach (string tempExcelPath in TempExcelPaths)
            {
                Status = $"Reading Excel { currentExcelCount.ToString() } of {  inputExcelPathsCount.ToString() }...";
                if (_log.IsInfoEnabled)
                {
                    _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
                }

                string agentNameColumn = Properties.Settings.Default.AgentNameColumn;
                string twelveAmColumn = Properties.Settings.Default.TwelveAmColumn;

                List<AgentStartStops> startStopList = await excelReader.GetAgentStartStopListAsync(tempExcelPath, SelectedName, agentNameColumn, twelveAmColumn, groupByNameCell, phoneColorKeyCell);

                IGetStatusesFromStartStops getStatusesFromStartStops = new GetStatusesFromStartStops();
                DateTime day = excelReader.WorksheetDay;

                Status = $"Retrieving agent statuses for Excel  { currentExcelCount.ToString() } of {  inputExcelPathsCount.ToString() }...";
                if (_log.IsInfoEnabled)
                {
                    _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
                }

                TimeZoneInfo excelTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Properties.Settings.Default.TimeZoneId);

                List<AgentStatuses> agentStatuses = await getStatusesFromStartStops.GetAgentStatusesListAsync(getStatuses, startStopList, day, excelTimeZone);
                IConsolidateAgentStatuses consolidateStatuses = new ConsolidateAgentStatuses();

                List<AgentStatuses> consolidatedAgentStatuses = await Task.Run(() => consolidateStatuses.Consolidate(agentStatuses));
                Status = $"Writing results to file for Excel  { currentExcelCount.ToString() } of {  inputExcelPathsCount.ToString() }...";
                if (_log.IsInfoEnabled)
                {
                    _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
                }
                IWriteResults writeResults = new WriteResultsToExcelFile();

                await Task.Run(() => writeResults.WriteResults(OutputPath, consolidatedAgentStatuses, SelectedName, excelReader.WorksheetDay));
                currentExcelCount++;
            }

            Status = "Job complete!";
            if (_log.IsInfoEnabled)
            {
                _log.Info("ShellViewModel.GenerateReportAsync - " + Status);
            }

        }

        public async Task Clear()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("ShellViewModel.Clear - Clearing InputExcelPath, OutputPath, SelectedTeam, TeamNames, GetTeamNamesRan, and Status properties");
            }
            InputExcelPaths = new List<string>();
            OutputPath = null;
            SelectedName = null;
            Names = new List<string>();
            GetNamesRan = false;
            ReportRan = false;
            Status = "";
            TempExcelPaths = new List<string>();

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

