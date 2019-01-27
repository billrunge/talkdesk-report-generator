using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalkdeskReportGenerator.Models
{
    public class SettingsModel
    {
        public ColumnRow PhoneColorKeyCell { get; set; }
        public ColumnRow GroupByNameCell { get; set; }
        public string AgentNameColumn { get; set; }
        public string TwelveAmColumn { get; set; }
        public TimeZoneInfo ExcelTimeZoneInfo { get; set; }

    }

    public class ColumnRow
    {
       public string Column { get; set; }
       public int Row { get; set; }
    }



    


}
