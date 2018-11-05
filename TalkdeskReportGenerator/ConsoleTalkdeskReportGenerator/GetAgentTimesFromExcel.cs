using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ConsoleTalkdeskReportGenerator
{
    internal interface IGetAgentTimes
    {
        List<AgentStartStops> GetAgentStartStopList(string filePath);
        DateTime WorkbookMonday { get; }
    }

    internal class GetAgentTimesFromExcel : IGetAgentTimes
    {

        public string TeamName { get; set; } = "RelativityOne";
        public string PhoneTimeCellFill { get; set; } = "Solid Color Theme: Accent1, Tint: 0.799981688894314";
        public int TeamNameColumn { get; set; } = 2;
        public int AgentNameColumn { get; set; } = 7;
        public int TwelveAmColumn { get; set; } = 8;
        public int ElevenPmColumn { get; }
        public DateTime WorkbookMonday { get; private set; }

        public GetAgentTimesFromExcel()
        {
            ElevenPmColumn = TwelveAmColumn + 23;
        }

        public List<AgentStartStops> GetAgentStartStopList(string filePath)
        {
            List<AgentStartStops> startStopList = new List<AgentStartStops>();

            //Using a Filestream so the Excel can be open while operation is occurring
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                XLWorkbook excel = new XLWorkbook(fs);
                int workSheetCount = excel.Worksheets.Count;

                //Use the worksheet count to return the last worksheet
                IXLWorksheet lastWorkSheet = excel.Worksheet(workSheetCount);

                //Get the range of relevant rows for the team in question
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
                if (row.Cell(TeamNameColumn).Value.ToString().Trim() == TeamName)
                {
                    string rowRangeString = row.Cell(TeamNameColumn).MergedRange().ToString();

                    /* 
                     * Value returned formatted like:
                     * <workbookName>!<columnLetter><rowNumber>:<columnLetter><rowNumber>
                     */

                    if (Regex.IsMatch(rowRangeString, "[0-9][0-9][.][0-9][0-9][.][0-9][0-9][!][aA-zZ]+[0-9]+[:][aA-zZ]+[0-9]+")) {

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

            agentStartStop.AgentName = row.Cell(AgentNameColumn).Value.ToString();

            for (int i = TwelveAmColumn; i <= ElevenPmColumn; i++)
            {
                if (row.Cell(i).Style.Fill.ToString() == PhoneTimeCellFill)
                {
                    phoneTimeColumns.Add(i);
                }
            }

            foreach (int column in phoneTimeColumns)
            {
                agentStartStop.StartStopList.Add(GetStartStopByCellPosition(column - TwelveAmColumn));
            }

            return agentStartStop;
        }

        /* This class will give you a timespan representing the start and stop midnight offset
         * based off of how many cells away it is from the midnight column */
        private StartStop GetStartStopByCellPosition(int position)
        {
            StartStop startStop = new StartStop();

            switch (position)
            {
                case 0:
                    startStop.Start = new TimeSpan(0, 0, 0);
                    startStop.Stop = new TimeSpan(0, 59, 59);
                    break;
                case 1:
                    startStop.Start = new TimeSpan(1, 0, 0);
                    startStop.Stop = new TimeSpan(1, 59, 59);
                    break;
                case 2:
                    startStop.Start = new TimeSpan(2, 0, 0);
                    startStop.Stop = new TimeSpan(2, 59, 59);
                    break;
                case 3:
                    startStop.Start = new TimeSpan(3, 0, 0);
                    startStop.Stop = new TimeSpan(3, 59, 59);
                    break;
                case 4:
                    startStop.Start = new TimeSpan(4, 0, 0);
                    startStop.Stop = new TimeSpan(4, 59, 59);
                    break;
                case 5:
                    startStop.Start = new TimeSpan(5, 0, 0);
                    startStop.Stop = new TimeSpan(5, 59, 59);
                    break;
                case 6:
                    startStop.Start = new TimeSpan(6, 0, 0);
                    startStop.Stop = new TimeSpan(6, 59, 59);
                    break;
                case 7:
                    startStop.Start = new TimeSpan(7, 0, 0);
                    startStop.Stop = new TimeSpan(7, 59, 59);
                    break;
                case 8:
                    startStop.Start = new TimeSpan(8, 0, 0);
                    startStop.Stop = new TimeSpan(8, 59, 59);
                    break;
                case 9:
                    startStop.Start = new TimeSpan(9, 0, 0);
                    startStop.Stop = new TimeSpan(9, 59, 59);
                    break;
                case 10:
                    startStop.Start = new TimeSpan(10, 0, 0);
                    startStop.Stop = new TimeSpan(10, 59, 59);
                    break;
                case 11:
                    startStop.Start = new TimeSpan(11, 0, 0);
                    startStop.Stop = new TimeSpan(11, 59, 59);
                    break;
                case 12:
                    startStop.Start = new TimeSpan(12, 0, 0);
                    startStop.Stop = new TimeSpan(12, 59, 59);
                    break;
                case 13:
                    startStop.Start = new TimeSpan(13, 0, 0);
                    startStop.Stop = new TimeSpan(13, 59, 59);
                    break;
                case 14:
                    startStop.Start = new TimeSpan(14, 0, 0);
                    startStop.Stop = new TimeSpan(14, 59, 59);
                    break;
                case 15:
                    startStop.Start = new TimeSpan(15, 0, 0);
                    startStop.Stop = new TimeSpan(15, 59, 59);
                    break;
                case 16:
                    startStop.Start = new TimeSpan(16, 0, 0);
                    startStop.Stop = new TimeSpan(16, 59, 59);
                    break;
                case 17:
                    startStop.Start = new TimeSpan(17, 0, 0);
                    startStop.Stop = new TimeSpan(17, 59, 59);
                    break;
                case 18:
                    startStop.Start = new TimeSpan(18, 0, 0);
                    startStop.Stop = new TimeSpan(18, 59, 59);
                    break;
                case 19:
                    startStop.Start = new TimeSpan(19, 0, 0);
                    startStop.Stop = new TimeSpan(19, 59, 59);
                    break;
                case 20:
                    startStop.Start = new TimeSpan(20, 0, 0);
                    startStop.Stop = new TimeSpan(20, 59, 59);
                    break;
                case 21:
                    startStop.Start = new TimeSpan(21, 0, 0);
                    startStop.Stop = new TimeSpan(21, 59, 59);
                    break;
                case 22:
                    startStop.Start = new TimeSpan(22, 0, 0);
                    startStop.Stop = new TimeSpan(22, 59, 59);
                    break;
                case 23:
                    startStop.Start = new TimeSpan(23, 0, 0);
                    startStop.Stop = new TimeSpan(23, 59, 59);
                    break;
                default:
                    throw new ArgumentOutOfRangeException($"{position} is an invalid offset from midnight");
            }
            return startStop;
        }

    }
}




