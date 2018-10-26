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
                //GetAgentStatuses getAgentStatuses = new GetAgentStatuses(db);
                //string userId = "58507b7e324f849110000017";
                //DateTime statusStart = new DateTime(2018, 10, 20);
                //DateTime statusEnd = new DateTime(2018, 10, 25);


                string filePath = @"C:\TalkdeskProject\Schedule.xlsx";

                IGetDataSet getDataSet = new GetDataSetFromExcel();
                GetAgentStartStops agentStartStops = new GetAgentStartStops(getDataSet, filePath);
                agentStartStops.GetAgentStartStopList();
                Console.ReadLine();








                //List<AgentStatus> agentStatuses = getAgentStatuses.GetAgentStatusesList(userId, statusStart, statusEnd);

                //foreach (var status in agentStatuses)
                //{
                //    Console.WriteLine($"{status.StatusLabel} : {status.StatusTime  / 60}");
                //}

                //Console.ReadLine();


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
