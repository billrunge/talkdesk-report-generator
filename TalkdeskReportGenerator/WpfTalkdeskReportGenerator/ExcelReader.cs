using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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
        private Workbook workbook;
        public DateTime WorkbookDay { get; private set; }

        public ExcelReader()
        {
            _phoneTimeCellFill = "Solid Color Theme: Accent1, Tint: 0.799981688894314";
            _teamNameColumn = 5;
            _agentNameColumn = 7;
            _twelveAmColumn = 8;
            _elevenPmColumn = _twelveAmColumn + 23;
        }

        public async Task<List<AgentStartStops>> GetAgentStartStopListAsync(string filePath, string teamName)
        {
            _teamName = teamName;

            List<AgentStartStops> startStopList = new List<AgentStartStops>();

            //Using a Filestream so the Excel can be open while operation is occurring
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;

                //Use the worksheet count to return the last worksheet
                IXLWorksheet lastWorkSheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                //Get the relevant rows for the team in question

                List<int> teamRows = await GetTeamRowsAsync(lastWorkSheet);

                //Extract date from Worksheet name
                await SetWorksheetDateAsync(lastWorkSheet);

                List<Task<AgentStartStops>> tasks = new List<Task<AgentStartStops>>();

                foreach (int row in teamRows)
                {
                    tasks.Add(GetAgentStartStopFromRowAsync(lastWorkSheet, row));
                }

                AgentStartStops[] results = await Task.WhenAll(tasks);

                foreach (AgentStartStops result in results)
                {
                    startStopList.Add(result);
                }

            }
            return startStopList;
        }

        public async Task<List<string>> GetTeamNamesAsync(string filePath)
        {
            List<string> teamNames = new List<string>();

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;

                //Use the worksheet count to return the last worksheet
                IXLWorksheet worksheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                IXLRows col = await Task.Run(() => worksheet.RowsUsed());

                foreach (IXLRow row in col)
                {
                    string cellValue = row.Cell(_teamNameColumn).Value.ToString().Trim();
                    if (!(string.IsNullOrEmpty(cellValue) || cellValue == "Team"))
                    {
                        teamNames.Add(cellValue);
                    }
                }
            }
            return await Task.Run(() => teamNames.Distinct().ToList());
        }

        public async Task<string> CreateLightweightExcelAsync(string filePath)
        {
            filePath = await Task.Run(() => filePath.ToLower());

            Microsoft.Office.Interop.Excel.Application excelApplication = await Task.Run(() => new Microsoft.Office.Interop.Excel.Application
            {
                DisplayAlerts = false,
                AskToUpdateLinks = false
            });

            workbook = await Task.Run(() => excelApplication.Workbooks.Open(filePath, XlUpdateLinks.xlUpdateLinksNever, true,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

            filePath = $"{Path.GetDirectoryName(filePath)}\\{Guid.NewGuid().ToString()}.xlsx";

            List<Task> tasks = new List<Task>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (!(Regex.IsMatch(sheet.Name, "[0-9]{1,2}[.][0-9]{1,2}[.][0-9]{2}")))
                {
                    tasks.Add(Task.Run(() => sheet.Delete()));
                }
            }

            await Task.WhenAll(tasks);

            await Task.Run(() => workbook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

            await Task.Run(() => workbook.Close(false, Type.Missing, Type.Missing));
            await Task.Run(() => excelApplication.Quit());
            return filePath;
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
                    teamRows.Add(currentRowAddress);
                }

            }
            return teamRows;
        }

        private async Task SetWorksheetDateAsync(IXLWorksheet worksheet)
        {
            //Extract workbook date so we can determine Monday later
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

            WorkbookDay = new DateTime(year, month, day);
        }

        private async Task<AgentStartStops> GetAgentStartStopFromRowAsync(IXLWorksheet worksheet, int rowNumber)
        {
            AgentStartStops agentStartStop = new AgentStartStops();
            List<int> phoneTimeColumns = new List<int>();

            IXLRow row = await Task.Run(() => worksheet.Row(rowNumber));

            agentStartStop.AgentName = row.Cell(_agentNameColumn).Value.ToString();

            for (int i = _twelveAmColumn; i <= _elevenPmColumn; i++)
            {
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


        /* This class will give you a timespan representing the start and stop midnight offset
         * based off of how many cells away it is from the midnight column */
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

        public async Task DeleteExcelAsync(string filePath)
        {
            await Task.Run(() => File.Delete(filePath));
        }

    }
}

