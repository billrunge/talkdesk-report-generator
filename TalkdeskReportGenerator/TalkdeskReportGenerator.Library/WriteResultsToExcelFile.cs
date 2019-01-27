using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace TalkdeskReportGenerator.Library
{
    public class WriteResultsToExcelFile : IWriteResults
    {
        public void WriteResults(string folderPath, List<AgentStatuses> consolidatedAgentStatuses, string teamName, DateTime workbookDate)
        {
            string date = workbookDate.ToShortDateString().Replace(@"/", "-");
            string filePath = $"{folderPath}TalkDesk - {teamName} - {date}.xlsx";


            using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet sheet = wb.Worksheets.Add(date);
                sheet.Cell(1, 1).Value = "Agent Name";
                sheet.Cell(1, 2).Value = "Status";
                sheet.Cell(1, 3).Value = "Minutes in status";
                sheet.Cell(1, 4).Value = "Compliance Percentage";
                int currentRow = 2;


                for (int i = 0; i < consolidatedAgentStatuses.Count; i++)
                {
                    sheet.Cell(currentRow, 1).Value = consolidatedAgentStatuses[i].AgentName;
                    int goodStatusTime = 0;
                    int totalStatusTime = 0;

                    foreach (Status status in consolidatedAgentStatuses[i].Statuses)
                    {
                        switch (status.StatusLabel)
                        {
                            case "Available":
                                goodStatusTime += status.StatusTime;
                                totalStatusTime += status.StatusTime;
                                break;
                            case "After Call Work":
                                goodStatusTime += status.StatusTime;
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
                    if (totalStatusTime > 0)
                    {
                        sheet.Cell(currentRow, 4).Value = $"{ (((decimal)goodStatusTime / (decimal)totalStatusTime) * 100).ToString("0.##") }%";
                    } else
                    {
                        sheet.Cell(currentRow, 4).Value = "0%";
                    }

                    currentRow += 1;

                    foreach (Status status in consolidatedAgentStatuses[i].Statuses)
                    {
                        sheet.Cell(currentRow, 2).Value = status.StatusLabel;
                        sheet.Cell(currentRow, 3).Value = ((decimal)status.StatusTime / 60).ToString("0.##");
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