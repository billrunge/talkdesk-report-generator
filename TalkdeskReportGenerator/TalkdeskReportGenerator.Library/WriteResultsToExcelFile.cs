using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace TalkdeskReportGenerator.Library
{
    public interface IWriteResults
    {
        void WriteResults(string folderPath, List<AgentData> consolidatedAgentData, string teamName, DateTime workbookDate);
    }

    public class WriteResultsToExcelFile : IWriteResults
    {
        public void WriteResults(string folderPath, List<AgentData> consolidatedAgentData, string teamName, DateTime workbookDate)
        {
            string date = workbookDate.ToShortDateString().Replace(@"/", "-");
            string filePath = $"{folderPath}TalkDesk - {teamName} - {date}.xlsx";
            int afterCallWorkSeconds = 120;
            int timeDivide = 3600;
            string timeDivideName = "Hours";

            
            using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet sheet = wb.Worksheets.Add(date);
                sheet.Cell(1, 1).Value = "Agent Name";
                sheet.Cell(1, 2).Value = "Status";
                sheet.Cell(1, 3).Value = $"{ timeDivideName } in Status";
                sheet.Cell(1, 4).Value = "Compliance Percentage";
                sheet.Cell(1, 5).Value = "Time in Compliance";
                sheet.Cell(1, 6).Value = "Total Scheduled Time";
                sheet.Cell(1, 7).Value = "Inbound Calls";
                sheet.Cell(1, 8).Value = "Outbound Calls";
                sheet.Cell(1, 9).Value = "Missed Calls";
                sheet.Cell(1, 10).Value = "Abandoned Calls";
                sheet.Cell(1, 11).Value = "Short Abandoned Calls";

                int currentRow = 2;


                for (int i = 0; i < consolidatedAgentData.Count; i++)
                {
                    sheet.Cell(currentRow, 1).Value = consolidatedAgentData[i].AgentName;

                    int inboundCalls = 0;
                    int outboundCalls = 0;
                    int missedCalls = 0;
                    int abandonedCalls = 0;
                    int shortAbandonedCalls = 0;


                    foreach (Call call in consolidatedAgentData[i].Calls)
                    {
                        switch (call.Type)
                        {
                            case CallType.inbound:
                                inboundCalls += call.Count;
                                break;
                            case CallType.outbound:
                                outboundCalls += call.Count;
                                break;
                            case CallType.missed:
                                missedCalls += call.Count;
                                break;
                            case CallType.abandoned:
                                abandonedCalls += call.Count;
                                break;
                            case CallType.short_abandoned:
                                shortAbandonedCalls += call.Count;
                                break;
                            default:
                                break;
                        }
                    }

                    int goodStatusTime = 0;
                    int totalStatusTime = 0;
                    int afterCallWork = 0;

                    foreach (Status status in consolidatedAgentData[i].Statuses)
                    {
                        switch (status.StatusLabel)
                        {
                            case "Available":
                                goodStatusTime += status.StatusTime;
                                totalStatusTime += status.StatusTime;
                                break;
                            case "After Call Work":
                                afterCallWork += status.StatusTime;
                                totalStatusTime += status.StatusTime;
                                break;
                            case "On a Call":
                                goodStatusTime += status.StatusTime;
                                totalStatusTime += status.StatusTime;
                                break;
                            default:
                                totalStatusTime += status.StatusTime;
                                break;
                        }
                    }

                    goodStatusTime += (afterCallWork < ((inboundCalls + outboundCalls) * afterCallWorkSeconds)) ? afterCallWork : ((inboundCalls + outboundCalls) * afterCallWorkSeconds);

                    if (totalStatusTime > 0)
                    {
                        sheet.Cell(currentRow, 4).Value = $"{ (((decimal)goodStatusTime / (decimal)totalStatusTime) * 100).ToString("0.##") }%";
                        sheet.Cell(currentRow, 5).Value = ((decimal)goodStatusTime / timeDivide).ToString("0.##"); 
                        sheet.Cell(currentRow, 6).Value = ((decimal)totalStatusTime / timeDivide).ToString("0.##");
                        sheet.Cell(currentRow, 7).Value = inboundCalls;
                        sheet.Cell(currentRow, 8).Value = outboundCalls;
                        sheet.Cell(currentRow, 9).Value = missedCalls;
                        sheet.Cell(currentRow, 10).Value = abandonedCalls;
                        sheet.Cell(currentRow, 11).Value = shortAbandonedCalls;

                    } else
                    {
                        sheet.Cell(currentRow, 4).Value = "0%";
                    }

                    currentRow += 1;

                    foreach (Status status in consolidatedAgentData[i].Statuses)
                    {
                        sheet.Cell(currentRow, 2).Value = status.StatusLabel;
                        sheet.Cell(currentRow, 3).Value = ((decimal)status.StatusTime / timeDivide).ToString("0.##");
                        currentRow += 1;
                    }               

                }
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(filePath);
            }
        }
    }
}