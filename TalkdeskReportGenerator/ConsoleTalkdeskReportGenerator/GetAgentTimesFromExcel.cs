using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConsoleTalkdeskReportGenerator
{
    class GetAgentTimesFromExcel
    {
        public void GetTimes(string filePath)
        {
            GetLatestWorkbookNumber(filePath);
        }

        private void GetLatestWorkbookNumber(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                ZipArchive excelZip = new ZipArchive(stream);

                foreach (var e in excelZip.Entries)
                {
                    int topSheetNum = 0;
                    

                    if (Regex.IsMatch(e.Name, "sheet[0-9]\\.xml$"))
                    {
                        string currentSheetNum = Regex.Match(e.Name, @"\d+").Value;

                        Console.WriteLine(currentSheetNum);
                    }
                }
            }

        }
    }

}
