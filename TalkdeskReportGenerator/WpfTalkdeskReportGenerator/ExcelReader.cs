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
        DateTime WorkbookDay { get; }
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
        public DateTime WorkbookDay { get; private set; }

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
            _log.Debug($"_teamName = { _teamName }");

            List<AgentStartStops> startStopList = new List<AgentStartStops>();

            _log.Debug($"Creating Filestream for working excel");
            using (FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                _log.Debug("Generting XLWorkbook object from Filestream");
                XLWorkbook excel = new XLWorkbook(stream);
                
                int workSheetCount = excel.Worksheets.Count;
                _log.Debug($"workSheetCount = { workSheetCount }");

                _log.Debug($"Using workSheetCount ({workSheetCount}) to return the last worksheet");
                IXLWorksheet lastWorkSheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                _log.Debug($"Getting a list of row numbers that represent the members of the selected team");
                List<int> teamRows = await GetTeamRowsAsync(lastWorkSheet);

                await SetWorksheetDateAsync(lastWorkSheet);
                _log.Debug($"WorksheetDay = {WorkbookDay.ToShortDateString()}");

                List<Task<AgentStartStops>> tasks = new List<Task<AgentStartStops>>();

                _log.Debug($"Creating list of tasks to retrieve lists of AgentStartStops");
                foreach (int row in teamRows)
                {
                    tasks.Add(GetAgentStartStopFromRowAsync(lastWorkSheet, row));
                }

                _log.Debug($"Awaiting list of Task<AgentStartStops> to complete");
                AgentStartStops[] results = await Task.WhenAll(tasks);

                _log.Debug($"Adding AgentStartStops to output list");
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

            _log.Debug($"Creating a new file stream to extract team names from source Excel at { excelPath }");
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;
                _log.Debug($"workSheetCount = { workSheetCount }");

                _log.Debug($"Getting last worksheet from Excel using workSheetCount as index");
                IXLWorksheet worksheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                _log.Debug($"Finding used rows in Excel");
                IXLRows col = await Task.Run(() => worksheet.RowsUsed());

                _log.Debug($"Checking each row in Excel to see if the value in _teamNameColumn matches up with selected team name");
                foreach (IXLRow row in col)
                {
                    string cellValue = row.Cell(_teamNameColumn).Value.ToString().Trim();
                    _log.Debug($"Checking if {cellValue} is not empty or does not equal 'Team'");
                    if (!(string.IsNullOrEmpty(cellValue) || cellValue == "Team"))
                    {
                        _log.Debug($"Adding { cellValue } to team list");
                        teamNames.Add(cellValue);
                    }
                }
            }
            _log.Debug("Finding distinct team names in list and returning consolidated team name list");
            return await Task.Run(() => teamNames.Distinct().ToList());
        }

        public async Task<string> CreateLightweightExcelAsync(string excelPath)
        {
            try
            {
                _log.Debug("Generating new instance of Excel application");
                Microsoft.Office.Interop.Excel.Application excelApplication = await Task.Run(() => new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false,
                    AskToUpdateLinks = false
                });

                _log.Debug($"Opening the excel at { excelPath }");
                workbook = await Task.Run(() => excelApplication.Workbooks.Open(excelPath, XlUpdateLinks.xlUpdateLinksNever, true,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                excelPath = $"{Path.GetDirectoryName(excelPath)}\\{Guid.NewGuid().ToString()}.xlsx";
                _log.Debug($"excelPath changed to { excelPath }");


                _log.Debug("Looping through each worksheet and deleting those without a format like 'Mm.Dd.YY'");
                /* Tried running this as a task list. Interop.Excel.Worksheet.Name did not like it. */
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (!(Regex.IsMatch(worksheet.Name, "[0-9]{1,2}[.][0-9]{1,2}[.][0-9]{2}")))
                    {
                        await Task.Run(() => worksheet.Delete());
                    }
                }

                _log.Debug($"Saving stripped down Excel to { excelPath }");
                await Task.Run(() => workbook.SaveAs(excelPath, XlFileFormat.xlWorkbookDefault,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                _log.Debug($"Closing excel in memory");
                await Task.Run(() => workbook.Close(false, Type.Missing, Type.Missing));

                _log.Debug($"Shutting down Excel application instance");
                await Task.Run(() => excelApplication.Quit());
                
            }
            catch (Exception e)
            {
                _log.Error($@"An error has occurred creating a working copy of the source Excel
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
            _log.Debug($"Finding used rows in Excel");
            IXLRows col = await Task.Run(() => worksheet.RowsUsed());

            _log.Debug("Looping through every row in the worksheet");
            foreach (IXLRow row in col)
            {
                _log.Debug("Parsing the current row address from cell address");
                if (!int.TryParse(await Task.Run(() => Regex.Replace(row.Cell(_teamNameColumn).Address.ToString(), "[^0-9.]", "")), out int currentRowAddress))
                {
                    throw new InvalidCastException("Unable to parse row int from cell address resturned from Excel");
                }

                _log.Debug($"If the value of the cell == { _teamName } add its row address to the row address list");
                if (row.Cell(_teamNameColumn).Value.ToString().Trim() == _teamName)
                {
                    teamRows.Add(currentRowAddress);
                }

            }
            return teamRows;
        }

        private async Task SetWorksheetDateAsync(IXLWorksheet worksheet)
        {

            _log.Debug($"Trying to parse month, day, and year from { worksheet.Name }");
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

            _log.Debug($"Returning parsed month, day, and year as a DateTime object");
            WorkbookDay = new DateTime(year, month, day);
        }

        private async Task<AgentStartStops> GetAgentStartStopFromRowAsync(IXLWorksheet worksheet, int rowNumber)
        {
            AgentStartStops agentStartStop = new AgentStartStops();
            List<int> phoneTimeColumns = new List<int>();

            _log.Debug($"Creating row object from worksheet and rowNumber { rowNumber }");
            IXLRow row = await Task.Run(() => worksheet.Row(rowNumber));

            agentStartStop.AgentName = row.Cell(_agentNameColumn).Value.ToString();
            _log.Debug($"Setting AgentName = { agentStartStop.AgentName }");


            _log.Debug("Looping through all columns from _twelveAmColumn to _elevenPmColumn");
            for (int i = _twelveAmColumn; i <= _elevenPmColumn; i++)
            {
                _log.Debug("Checking to see if cell value's fill matches the configured _phoneTimeCellFill value");
                if (row.Cell(i).Style.Fill.ToString() == _phoneTimeCellFill)
                {
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
                agentStartStop.StartStopList.Add(startStop);
            }

            return agentStartStop;
        }

        private async Task<StartStop> GetStartStopByCellPositionAsync(int position)
        {
            _log.Debug("Converting position offset int to TimeSpan offset from midnight");
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
            _log.Debug($"Deleting file at { excelPath }");
            await Task.Run(() => File.Delete(excelPath));
        }

    }
}

