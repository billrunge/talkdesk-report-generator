using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using log4net;

namespace WpfTalkdeskReportGenerator
{
    public interface IExcelReader
    {
        Task<List<AgentStartStops>> GetAgentStartStopListAsync(string filePath, string teamName);
        Task<List<string>> GetTeamNamesAsync(string filePath);
        Task<string> CreateLightweightExcelAsync(string filePath);
        Task DeleteExcelAsync(string filePath);
        DateTime WorksheetDay { get; }
    }

    public class ExcelReader : IExcelReader
    {
        private string _teamName;
        private readonly string _phoneTimeCellFill;
        private readonly int _teamNameColumn;
        private readonly int _agentNameColumn;
        private readonly int _twelveAmColumn;
        private readonly int _elevenPmColumn;
        private readonly ILog _log;

        private Workbook workbook;
        public DateTime WorksheetDay { get; private set; }

        public ExcelReader(ILog log)
        {
            _phoneTimeCellFill = "Solid Color Theme: Accent1, Tint: 0.799981688894314";
            _teamNameColumn = 5;
            _agentNameColumn = 7;
            _twelveAmColumn = 8;
            _elevenPmColumn = _twelveAmColumn + 23;
            _log = log;
        }

        public async Task<List<AgentStartStops>> GetAgentStartStopListAsync(string excelPath, string teamName)
        {
            _teamName = teamName;
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetAgentStartStopListAsync - Setting _teamNAme = { _teamName }");
            }

            List<AgentStartStops> startStopList = new List<AgentStartStops>();

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetAgentStartStopListAsync - Creating Filestream for working excel at {excelPath}");
            }
            using (FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(stream);
                
                int workSheetCount = excel.Worksheets.Count;
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetAgentStartStopListAsync - workSheetCount = { workSheetCount }");
                }

                IXLWorksheet lastWorksheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                List<int> teamRows = await GetTeamRowsAsync(lastWorksheet);

                await SetWorksheetDateAsync(lastWorksheet);
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetAgentStartStopListAsync - WorksheetDay = {WorksheetDay.ToShortDateString()}");
                }

                List<Task<AgentStartStops>> tasks = new List<Task<AgentStartStops>>();

                foreach (int row in teamRows)
                {
                    tasks.Add(GetAgentStartStopFromRowAsync(lastWorksheet, row));
                }

                AgentStartStops[] results = await Task.WhenAll(tasks);

                foreach (AgentStartStops result in results)
                {
                    startStopList.Add(result);
                }

            }
            return startStopList;
        }

        public async Task<List<string>> GetTeamNamesAsync(string excelPath)
        {
            List<string> teamNames = new List<string>();

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetTeamNamesAsync - Creating a new file stream to extract team names from source Excel at { excelPath }");
            }
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetTeamNamesAsync - workSheetCount = { workSheetCount }");
                }

                IXLWorksheet worksheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                IXLRows col = await Task.Run(() => worksheet.RowsUsed());

                foreach (IXLRow row in col)
                {
                    string cellValue = row.Cell(_teamNameColumn).Value.ToString().Trim();
                    if (!(string.IsNullOrEmpty(cellValue) || cellValue == "Team"))
                    {
                        if (_log.IsDebugEnabled)
                        {
                            _log.Debug($"ExcelReader.GetTeamNamesAsync - Adding { cellValue } to team list");
                        }
                        teamNames.Add(cellValue);
                    }
                }
            }
            return await Task.Run(() => teamNames.Distinct().ToList());
        }

        public async Task<string> CreateLightweightExcelAsync(string excelPath)
        {
            try
            {
                if (_log.IsDebugEnabled)
                {
                    _log.Debug("ExcelReader.CreateLightweightExcelAsync - Generating new instance of Excel application");
                }
                Microsoft.Office.Interop.Excel.Application excelApplication = await Task.Run(() => new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false,
                    AskToUpdateLinks = false
                });
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.CreateLightweightExcelAsync - Opening the excel at { excelPath }");
                }
                workbook = await Task.Run(() => excelApplication.Workbooks.Open(excelPath, XlUpdateLinks.xlUpdateLinksNever, true,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                excelPath = $"{Path.GetDirectoryName(excelPath)}\\{Guid.NewGuid().ToString()}.xlsx";
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.CreateLightweightExcelAsync - excelPath changed to { excelPath }");
                }
                    /* Tried running this as a task list. Interop.Excel.Worksheet.Name did not like it. */
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (!(Regex.IsMatch(worksheet.Name, "[0-9]{1,2}[.][0-9]{1,2}[.][0-9]{2}")))
                    {
                        await Task.Run(() => worksheet.Delete());
                    }
                }

                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.CreateLightweightExcelAsync - Saving stripped down Excel to { excelPath }");
                }
                await Task.Run(() => workbook.SaveAs(excelPath, XlFileFormat.xlWorkbookDefault,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.CreateLightweightExcelAsync - Closing open Excel");
                }
                await Task.Run(() => workbook.Close(false, Type.Missing, Type.Missing));

                _log.Debug($"Shutting down Excel application instance");
                await Task.Run(() => excelApplication.Quit());
                
            }
            catch (Exception e)
            {
                _log.Error($@"ExcelReader.CreateLightweightExcelAsync - An error has occurred creating a working copy of the source Excel
                              { e.Message } {Environment.NewLine}
                              {e.StackTrace}");
                MessageBox.Show($@"{ e.Message } {Environment.NewLine}
                                   {e.StackTrace}");
            }
            return excelPath;
        }

        private async Task<List<int>> GetTeamRowsAsync(IXLWorksheet worksheet)
        {
            List<int> teamRows = new List<int>();
            IXLRows col = await Task.Run(() => worksheet.RowsUsed());

            foreach (IXLRow row in col)
            {

                if (!int.TryParse(await Task.Run(() => Regex.Replace(row.Cell(_teamNameColumn).Address.ToString(), "[^0-9.]", "")), out int currentRowAddress))
                {
                    throw new InvalidCastException("Unable to parse row int from cell address resturned from Excel");
                }


                if (row.Cell(_teamNameColumn).Value.ToString().Trim() == _teamName)
                {
                    if (_log.IsDebugEnabled)
                    {
                        _log.Debug($"ExcelReader.GetTeamRowsAsync - Adding { currentRowAddress } to the teamRows list");
                    }
                    teamRows.Add(currentRowAddress);
                }

            }
            return teamRows;
        }

        private async Task SetWorksheetDateAsync(IXLWorksheet worksheet)
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.SetWorksheetDateAsync - Trying to parse month, day, and year from '{ worksheet.Name }'");
            }
            string dateString = worksheet.Name;

            if (!int.TryParse(dateString.Split('.')[0], out int month))
            {
                throw new FormatException($"Unable to parse {dateString.Split('.')[0]} to month int");
            }

            if (!int.TryParse(dateString.Split('.')[1], out int day))
            {
                throw new FormatException($"Unable to parse {dateString.Split('.')[1]} to day int");
            }

            if (!int.TryParse($"20{dateString.Split('.')[2]}", out int year))
            {
                throw new FormatException($"Unable to parse {dateString.Split('.')[2]} to year int");
            }

            WorksheetDay = new DateTime(year, month, day);
        }

        private async Task<AgentStartStops> GetAgentStartStopFromRowAsync(IXLWorksheet worksheet, int rowNumber)
        {
            AgentStartStops agentStartStop = new AgentStartStops();
            List<int> phoneTimeColumns = new List<int>();

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetAgentStartStopFromRowAsync - Creating row object from worksheet and rowNumber { rowNumber }");
            }
            IXLRow row = await Task.Run(() => worksheet.Row(rowNumber));

            agentStartStop.AgentName = row.Cell(_agentNameColumn).Value.ToString();
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetAgentStartStopFromRowAsync - Setting AgentName = { agentStartStop.AgentName }");
            }

            for (int i = _twelveAmColumn; i <= _elevenPmColumn; i++)
            {
                if (row.Cell(i).Style.Fill.ToString() == _phoneTimeCellFill)
                {
                    if (_log.IsDebugEnabled)
                    {
                        _log.Debug($"ExcelReader.GetAgentStartStopFromRowAsync - Adding {i} to phoneTimeColumns List<int>");
                    }
                    phoneTimeColumns.Add(i);
                }
            }

            List<Task<StartStop>> tasks = new List<Task<StartStop>>();

            foreach (int column in phoneTimeColumns)
            {
                tasks.Add(GetStartStopByCellPositionAsync(column - _twelveAmColumn));
            }

            StartStop[] results = await Task.WhenAll(tasks);

            foreach (StartStop startStop in results)
            {
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetAgentStartStopFromRowAsync - Adding start:{ startStop.Start } and stop: { startStop.Stop } to agentStartStop.StartStopList");
                }
                agentStartStop.StartStopList.Add(startStop);
            }

            return agentStartStop;
        }

        private async Task<StartStop> GetStartStopByCellPositionAsync(int position)
        {
            if (position > -1 && position < 24)
            {
                return new StartStop
                {
                    Start = new TimeSpan(position, 0, 0),
                    Stop = new TimeSpan(position, 59, 59)
                };

            }
            else
            {
                throw new ArgumentOutOfRangeException($"{position} is an invalid positive offset from midnight");
            }
        }

        public async Task DeleteExcelAsync(string excelPath)
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.DeleteExcelAsync - Deleting file at { excelPath }");
            }

            await Task.Run(() => File.Delete(excelPath));
        }

    }
}

