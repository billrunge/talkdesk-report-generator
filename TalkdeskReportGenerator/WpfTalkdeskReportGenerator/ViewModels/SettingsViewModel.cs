using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WpfTalkdeskReportGenerator.Models;

namespace WpfTalkdeskReportGenerator.ViewModels
{
    public class SettingsViewModel : Screen
    {
        private SettingsModel _settings;
        public SettingsModel Settings {
            get
            {
                return _settings;
            }
            set
            {
                _settings = value;
                NotifyOfPropertyChange(() => Settings);
            }
        }

        public List<string> TimeZones { get; set; }

        public SettingsViewModel()
        {
            TimeZones = GetTimeZoneList();
            Settings = new SettingsModel
            {
                ExcelTimeZone = Properties.Settings.Default.TimeZoneName,
                PhoneColorKeyCell = new ColumnRow
                {
                    Column = Properties.Settings.Default.PhoneColorKeyColumn,
                    Row = Properties.Settings.Default.PhoneColorKeyRow
                },
                GroupByNameCell = new ColumnRow
                {
                    Column = Properties.Settings.Default.GroupByNameColumn,
                    Row = Properties.Settings.Default.GroupByNameRow
                },
                AgentNameColumn = Properties.Settings.Default.AgentNameColumn,
                TwelveAmColumn = Properties.Settings.Default.TwelveAmColumn
            };



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
        public void Save()
        {
            Properties.Settings.Default.TimeZoneName = Settings.ExcelTimeZone;
            Properties.Settings.Default.PhoneColorKeyColumn = Settings.PhoneColorKeyCell.Column;
            Properties.Settings.Default.PhoneColorKeyRow = Settings.PhoneColorKeyCell.Row;
            Properties.Settings.Default.GroupByNameColumn = Settings.GroupByNameCell.Column;
            Properties.Settings.Default.GroupByNameRow = Settings.GroupByNameCell.Row;
            Properties.Settings.Default.AgentNameColumn = Settings.AgentNameColumn;
            Properties.Settings.Default.TwelveAmColumn = Settings.TwelveAmColumn;
            Properties.Settings.Default.Save();
        }
    }
}
