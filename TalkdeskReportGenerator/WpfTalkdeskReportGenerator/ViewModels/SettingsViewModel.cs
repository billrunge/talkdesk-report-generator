using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfTalkdeskReportGenerator.ViewModels
{
    public class SettingsViewModel : Screen
    {
        public List<string> TimeZones { get; set; }



        public SettingsViewModel()
        {
            TimeZones = GetTimeZoneList();
        }


        private List<string> GetTimeZoneList()
        {
            List<string> timeZoneNames = new List<string>();
            IReadOnlyCollection<TimeZoneInfo> timeZones = TimeZoneInfo.GetSystemTimeZones();

            foreach(TimeZoneInfo tz in timeZones)
            {
                timeZoneNames.Add(tz.DisplayName);
            }
            return timeZoneNames;
        }


        public void Back()
        {
            ActivateWindow.ViewReports();
        }
    }
}
