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
            Database db = new Database();
            GetAgents ga = new GetAgents(db);
            ga.GetAgentsList();
        }
    }
}
