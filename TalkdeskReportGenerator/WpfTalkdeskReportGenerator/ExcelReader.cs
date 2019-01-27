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
        Task<List<AgentStartStops>> GetAgentStartStopListAsync(string filePath, string teamName, string agentNameColumn, string twelveAmColumn, ExcelCell groupByNameCell, ExcelCell phoneColorKeyCell);
        Task<List<string>> GetNamesAsync(string filePath, ExcelCell groupByNameCell);
        Task<string> CreateLightweightExcelAsync(string filePath);
        Task DeleteExcelAsync(string filePath);
        DateTime WorksheetDay { get; }
    }

    public class ExcelReader : IExcelReader
    {
        private readonly ILog _log;
        private Workbook workbook;
        public DateTime WorksheetDay { get; private set; }

        public ExcelReader(ILog log)
        {
            _log = log;
        }

        public async Task<List<AgentStartStops>> GetAgentStartStopListAsync(string excelPath, string columnName, string agentNameColumn, string twelveAmColumn, ExcelCell groupByNameCell, ExcelCell phoneColorKeyCell)
        {
            if (_log.IsDebugEnabled)    
            {
                _log.Debug($"ExcelReader.GetAgentStartStopListAsync - Setting _teamName = { columnName }");
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

                List<int> teamRows = await GetTeamRowsAsync(lastWorksheet, columnName, groupByNameCell);

                await SetWorksheetDateAsync(lastWorksheet);
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetAgentStartStopListAsync - WorksheetDay = {WorksheetDay.ToShortDateString()}");
                }

                List<Task<AgentStartStops>> tasks = new List<Task<AgentStartStops>>();

                foreach (int row in teamRows)
                {
                    tasks.Add(GetAgentStartStopFromRowAsync(lastWorksheet, row, agentNameColumn, twelveAmColumn, phoneColorKeyCell));
                }

                AgentStartStops[] results = await Task.WhenAll(tasks);

                foreach (AgentStartStops result in results)
                {
                    startStopList.Add(result);
                }

            }
            return startStopList;
        }  

        public async Task<List<string>> GetNamesAsync(string excelPath, ExcelCell groupByNameCell)
        {
            List<string> managerNames = new List<string>();

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetManagerNamesAsync - Creating a new file stream to extract  names from source Excel at { excelPath }");
            }
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetManagerNamesAsync - workSheetCount = { workSheetCount }");
                }

                IXLWorksheet worksheet = await Task.Run(() => excel.Worksheet(workSheetCount));

                string nameColumnHeader = worksheet.Row(groupByNameCell.Row)
                    .Cell(XLHelper.GetColumnNumberFromLetter(groupByNameCell.Column)).Value.ToString();
                 

                IXLRows rows = await Task.Run(() => worksheet.RowsUsed());

                foreach (IXLRow row in rows)
                {
                    string cellValue = row.Cell(XLHelper.GetColumnNumberFromLetter(groupByNameCell.Column)).Value.ToString().Trim();
                    if (!(string.IsNullOrEmpty(cellValue) || cellValue == nameColumnHeader))
                    {
                        if (_log.IsDebugEnabled)
                        {
                            _log.Debug($"ExcelReader.GetManagerNamesAsync - Adding { cellValue } to manager list");
                        }
                        managerNames.Add(cellValue);
                    }
                }
            }
            return await Task.Run(() => managerNames.Distinct().ToList());
        }

        public async Task<string> GetGroupByNameAsync(string excelPath, ExcelCell groupByNameCell)
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetGroupByNameAsync - Creating a new file stream to extract  names from source Excel at { excelPath }");
            }

            string groupName;

            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;
                if (_log.IsDebugEnabled)
                {
                    _log.Debug($"ExcelReader.GetGroupByNameAsync - workSheetCount = { workSheetCount }");
                }
                IXLWorksheet worksheet = excel.Worksheet(workSheetCount);
                groupName = worksheet.Row(groupByNameCell.Row).Cell(XLHelper.GetColumnNumberFromLetter(groupByNameCell.Column)).Value.ToString(); 
            }
            return groupName;
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
                    if (!Regex.IsMatch(worksheet.Name, "[0-9]{1,2}[.][0-9]{1,2}[.][0-9]{2}") || worksheet.Visible == XlSheetVisibility.xlSheetHidden)
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

        private async Task<List<int>> GetTeamRowsAsync(IXLWorksheet worksheet,string columnName, ExcelCell groupByNameCell)
        {
            List<int> teamRows = new List<int>();
            IXLRows col = await Task.Run(() => worksheet.RowsUsed());

            foreach (IXLRow row in col)
            {

                int columnNumber = XLHelper.GetColumnNumberFromLetter(groupByNameCell.Column);

                if (!int.TryParse(await Task.Run(() => Regex.Replace(row.Cell(columnNumber).Address.ToString(), "[^0-9.]", "")), out int currentRowAddress))
                {
                    throw new InvalidCastException("Unable to parse row int from cell address resturned from Excel");
                }


                if (row.Cell(columnNumber).Value.ToString().Trim() == columnName)
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

        private async Task<AgentStartStops> GetAgentStartStopFromRowAsync(IXLWorksheet worksheet, int rowNumber, string agentNameColumn, string twelveAmColumn, ExcelCell phoneColorKeyCell)
        {
            AgentStartStops agentStartStop = new AgentStartStops();
            List<int> phoneTimeColumns = new List<int>();

            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetAgentStartStopFromRowAsync - Creating row object from worksheet and rowNumber { rowNumber }");
            }
            IXLRow row = await Task.Run(() => worksheet.Row(rowNumber));

            agentStartStop.AgentName = row.Cell(XLHelper.GetColumnNumberFromLetter(agentNameColumn)).Value.ToString();
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"ExcelReader.GetAgentStartStopFromRowAsync - Setting AgentName = { agentStartStop.AgentName }");
            }
            int twelveAmColumnInt = XLHelper.GetColumnNumberFromLetter(twelveAmColumn);

            for (int i = twelveAmColumnInt; i <= twelveAmColumnInt + 23; i++)
            {
                if (row.Cell(i).Style.Fill.ToString() == await GetPhoneTimeCellFill(worksheet, phoneColorKeyCell))
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
                tasks.Add(GetStartStopByCellPositionAsync(column - twelveAmColumnInt));
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


        private async Task<string> GetPhoneTimeCellFill(IXLWorksheet worksheet, ExcelCell excelCell)
        {
            IXLRow sheetRow = await Task.Run(() => worksheet.Row(excelCell.Row));
            IXLCell cell = await Task.Run(() => sheetRow.Cell(XLHelper.GetColumnNumberFromLetter(excelCell.Column)));
            return cell.Style.Fill.ToString();
        }

    }
}

