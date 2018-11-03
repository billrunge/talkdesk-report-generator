using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTalkdeskReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                IDatabase db = new Database();

                string filePath = @"C:\TalkdeskProject\Schedule.xlsx";

                IGetAgentTimes getAgentTimes = new GetAgentTimesFromExcel();

                List<AgentStartStop> startStopList = getAgentTimes.GetAgentStartStopList(filePath);
                foreach (var startStop in startStopList)
                {
                    Console.WriteLine(startStop.AgentName);
                    foreach (var s in startStop.StartStopList)
                    {
                        Console.WriteLine($"Start {s.Start}, End {s.Stop}");
                    }
                }
                Console.ReadLine();

            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Console.ReadLine();
            }
        }
    }
}
