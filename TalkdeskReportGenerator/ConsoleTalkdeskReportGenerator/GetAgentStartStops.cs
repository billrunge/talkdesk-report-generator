using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTalkdeskReportGenerator
{
    class GetAgentStartStops
    {
        private IGetDataSet _getDataSet;
        private string _filePath;

        public GetAgentStartStops(IGetDataSet getDataSet, string filePath)
        {
            _getDataSet = getDataSet;
            _filePath = filePath;
        }

        public void GetAgentStartStopList()
        {
            DataSet dataSet = _getDataSet.GetDataSet(_filePath);

            DataTable dataTable = dataSet.Tables[5];

            Console.WriteLine(dataTable.Rows[70][15]);
  
        }


    }
}
