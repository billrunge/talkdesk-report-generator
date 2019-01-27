using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalkdeskReportGenerator.Models
{
    public class ExcelStatusModel
    {
        public enum Status
        {
            Waiting,
            Loaded,
            Error
        };

        public Status CurrentStatus { get; set; }

    }
}
