using ExcelDataReader;
using System.Data;
using System.IO;

namespace ConsoleTalkdeskReportGenerator
{
    public interface IGetDataSet
    {
        DataSet GetDataSet(string filePath);
    }

    class GetDataSetFromExcel : IGetDataSet
    {
        public DataSet GetDataSet(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet();
                }
            }
        }
    }
}