using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace WpfTalkdeskReportGenerator
{
    public interface IExcelReader
    {
        List<AgentStartStops> GetAgentStartStopList(string filePath);
        DateTime WorkbookDay { get; }
    }

    public class ExcelReader : IExcelReader
    {
        private readonly string _teamName = "RelativityOne";
        private readonly string _phoneTimeCellFill;
        private readonly int _teamNameColumn;
        private readonly int _agentNameColumn;
        private readonly int _twelveAmColumn;
        private readonly int _elevenPmColumn;
        public DateTime WorkbookDay { get; private set; }

        public ExcelReader()
        {
            _phoneTimeCellFill = "Solid Color Theme: Accent1, Tint: 0.799981688894314";
            _teamName = "RelativityOne";
            _teamNameColumn = 5;
            _agentNameColumn = 7;
            _twelveAmColumn = 8;
            _elevenPmColumn = _twelveAmColumn + 23;
        }

        public List<AgentStartStops> GetAgentStartStopList(string filePath)
        {
            List<AgentStartStops> startStopList = new List<AgentStartStops>();

            filePath = filePath.ToLower();

            if (filePath.Contains(".xlsb"))
            {
                Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false,
                    AskToUpdateLinks = false
                };

                Workbook workbook = excelApplication.Workbooks.Open(filePath, XlUpdateLinks.xlUpdateLinksNever, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                filePath = filePath.Replace(".xlsb", ".xlsx");

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (!(Regex.IsMatch(sheet.Name, "[0-9]{1,2}[.][0-9]{1,2}[.][0-9]{2}")))
                    {
                        sheet.Delete();
                    }

                }
                workbook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook.Close(false, Type.Missing, Type.Missing);
                excelApplication.Quit();
            }

            //Using a Filestream so the Excel can be open while operation is occurring
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;

                //Use the worksheet count to return the last worksheet
                IXLWorksheet lastWorkSheet = excel.Worksheet(workSheetCount);

                //Get the range of relevant rows for the team in question
                ExcelRowRange range = GetRowRange(lastWorkSheet);
                
                //Extract date from Worksheet name
                SetWorksheetDate(lastWorkSheet);

                for (int i = range.FirstValue; i <= range.SecondValue; i++)
                {
                    AgentStartStops agentStartStop = GetAgentStartStopFromRow(lastWorkSheet, i);
                    startStopList.Add(agentStartStop);
                }
                return startStopList;
            }
        }

        private ExcelRowRange GetRowRange(IXLWorksheet worksheet)
        {
            ExcelRowRange excelRowRange = new ExcelRowRange();
            IXLRows col = worksheet.RowsUsed();

            int firstValue = 2147483647;
            int secondValue = -2147483648;

            foreach (IXLRow row in col)
            {
                bool isTeamRow = (row.Cell(_teamNameColumn).Value.ToString().Trim() == _teamName);

                if (!int.TryParse(Regex.Replace(row.Cell(_teamNameColumn).Address.ToString(), "[^0-9.]", ""), out int currentRowAddress))
                {
                    throw new InvalidCastException("Unable to parse row int from cell address resturned from Excel");
                }

                if (isTeamRow && currentRowAddress < firstValue)
                {
                    firstValue = currentRowAddress;
                }
                else if (isTeamRow && currentRowAddress > secondValue)
                {
                    secondValue = currentRowAddress;
                }
            }

            excelRowRange.FirstValue = firstValue;
            excelRowRange.SecondValue = secondValue;
            return excelRowRange;
        }

        private void SetWorksheetDate(IXLWorksheet worksheet)
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

        private AgentStartStops GetAgentStartStopFromRow(IXLWorksheet worksheet, int rowNumber)
        {
            AgentStartStops agentStartStop = new AgentStartStops();
            List<int> phoneTimeColumns = new List<int>();

            IXLRow row = worksheet.Row(rowNumber);

            agentStartStop.AgentName = row.Cell(_agentNameColumn).Value.ToString();

            for (int i = _twelveAmColumn; i <= _elevenPmColumn; i++)
            {
                if (row.Cell(i).Style.Fill.ToString() == _phoneTimeCellFill)
                {
                    phoneTimeColumns.Add(i);
                }
            }

            foreach (int column in phoneTimeColumns)
            {
                agentStartStop.StartStopList.Add(GetStartStopByCellPosition(column - _twelveAmColumn));
            }

            return agentStartStop;
        }

        /* This class will give you a timespan representing the start and stop midnight offset
         * based off of how many cells away it is from the midnight column */
        private StartStop GetStartStopByCellPosition(int position)
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

    }
}




