using System;

namespace TalkdeskReportGenerator.Library
{
    public class UserMapping
    {
        public string ExcelUser { get; set; }
        public string TalkdeskUser { get; set; }
        public Guid GUID { get; }

        public UserMapping()
        {
            GUID = Guid.NewGuid();
        }
    }
}
