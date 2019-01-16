using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using WpfTalkdeskReportGenerator;
using System.Windows;

namespace WpfTalkdeskReportGenerator
{
    public interface IExcelReader
    {
        List<AgentStartStops> GetAgentStartStopList(string filePath);
        DateTime WorkbookMonday { get; }
    }

    public class ExcelReader : IExcelReader
    {
        private readonly string _teamName = "RelativityOne";
        private readonly string _phoneTimeCellFill;
        private readonly string _rowRangeRegEx;
        private readonly int _teamNameColumn;
        private readonly int _agentNameColumn;
        private readonly int _twelveAmColumn;
        private readonly int _elevenPmColumn;
        public DateTime WorkbookMonday { get; private set; }

        public ExcelReader()
        {
            _phoneTimeCellFill = "Solid Color Theme: Accent1, Tint: 0.799981688894314";
            _teamName = "RelativityOne";
            _teamNameColumn = 2;
            _agentNameColumn = 7;
            _twelveAmColumn = 8;
            _elevenPmColumn = _twelveAmColumn + 23;
            _rowRangeRegEx = "[0-9][0-9][.][0-9][0-9][.][0-9][0-9][!][aA-zZ]+[0-9]+[:][aA-zZ]+[0-9]+";

        }

        public List<AgentStartStops> GetAgentStartStopList(string filePath)
        {
            MessageBox.Show("Start GetAGentStartStopList");
            List<AgentStartStops> startStopList = new List<AgentStartStops>();

            filePath = filePath.ToLower();

            if (filePath.Contains(".xlsb"))
            {
                var excelApplication = new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false,
                    AskToUpdateLinks = false
                };
                Workbook workbook = excelApplication.Workbooks.Open(filePath, XlUpdateLinks.xlUpdateLinksNever, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                filePath = filePath.Replace(".xlsb", ".xlsx");

                Array links = (Array)(object)workbook.LinkSources(XlLink.xlExcelLinks);

                if (links.Length > 0)
                {
                    for (int i = 1; i <= links.Length; i++)
                    {
                        workbook.BreakLink((string)links.GetValue(i), XlLinkType.xlLinkTypeExcelLinks);
                    }
                }

                foreach (WorkbookConnection connection in workbook.Connections)
                {
                    connection.Delete();                                
                }

                foreach(Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name != "1.14.19")
                    {
                        sheet.Delete();
                    }                    
                }


                workbook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook.Close(false, Type.Missing, Type.Missing);
                excelApplication.Quit();
            }
            MessageBox.Show("XLSB conversion complete");

            //Using a Filestream so the Excel can be open while operation is occurring
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                MessageBox.Show("Filestream starting");

                XLWorkbook excel = new XLWorkbook(fs);

                MessageBox.Show("Getting worksheet count");
                int workSheetCount = excel.Worksheets.Count;

                MessageBox.Show("Getting last worksheet");
                //Use the worksheet count to return the last worksheet
                IXLWorksheet lastWorkSheet = excel.Worksheet(workSheetCount);

                //Get the range of relevant rows for the team in question

                MessageBox.Show("Get row range");
                ExcelRowRange range = GetRowRange(lastWorkSheet);

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

            foreach (IXLRow row in col)
            {
                if (row.Cell(_teamNameColumn).Value.ToString().Trim() == _teamName)
                {
                    string rowRangeString = row.Cell(_teamNameColumn).MergedRange().ToString();

                    /* 
                     * Value returned formatted like:
                     * <workbookName>!<columnLetter><rowNumber>:<columnLetter><rowNumber>
                     */

                    if (Regex.IsMatch(rowRangeString, _rowRangeRegEx))
                    {

                        //Extract workbook date so we can determine Monday later
                        string dateString = rowRangeString.Split('!')[0];

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

                        DateTime workbookDay = new DateTime(year, month, day);

                        switch (workbookDay.DayOfWeek)
                        {
                            case DayOfWeek.Monday:
                                WorkbookMonday = workbookDay;
                                break;
                            case DayOfWeek.Tuesday:
                                WorkbookMonday = workbookDay.AddDays(-1);
                                break;
                            case DayOfWeek.Wednesday:
                                WorkbookMonday = workbookDay.AddDays(-2);
                                break;
                            case DayOfWeek.Thursday:
                                WorkbookMonday = workbookDay.AddDays(-3);
                                break;
                            case DayOfWeek.Friday:
                                WorkbookMonday = workbookDay.AddDays(-4);
                                break;
                            default:
                                throw new ArgumentOutOfRangeException($"{workbookDay.DayOfWeek.ToString()} is not a valid weekday");
                        }

                        rowRangeString = rowRangeString.Substring(rowRangeString.IndexOf("!") + 1);

                        if (!int.TryParse(Regex.Replace(rowRangeString.Split(':')[0], "[^0-9]", ""), out int firstValue))
                        {
                            throw new FormatException($"Unable to parse {rowRangeString.Split(':')[0]} to firstValue int");
                        }

                        if (!int.TryParse(Regex.Replace(rowRangeString.Split(':')[1], "[^0-9]", ""), out int secondValue))
                        {
                            throw new FormatException($"Unable to parse {rowRangeString.Split(':')[1]} to secondValue int");
                        }

                        excelRowRange.FirstValue = firstValue;
                        excelRowRange.SecondValue = secondValue;
                    }
                    else
                    {
                        throw new FormatException($"The row range string retrieved from the Excel was invalid. String received: {rowRangeString}");
                    }
                }
            }
            return excelRowRange;
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




