using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TalkdeskReportGenerator.Models;

namespace TalkdeskReportGenerator.ViewModels
{
    public class SettingsViewModel : Screen
    { 
        private TimeZoneInfo _selectedTimeZone;
        private int _phoneColorKeyRow;
        private int _groupByNameRow;
        private string _phoneColorKeyColumn;
        private string _groupByNameColumn;
        private string _agentNameColumn;
        private string _twelveAmColumn;

        public TimeZoneInfo SelectedTimeZone
        {
            get
            {
                return _selectedTimeZone;
            }
            set
            {
                _selectedTimeZone = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => SelectedTimeZone);
            }
        }
        public int PhoneColorKeyRow
        {
            get
            {
                return _phoneColorKeyRow;
            }
            set
            {
                _phoneColorKeyRow = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => PhoneColorKeyRow);
            }
        }
        public int GroupByNameRow
        {
            get
            {
                return _groupByNameRow;
            }
            set
            {
                _groupByNameRow = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => GroupByNameRow);
            }
        }
        public string PhoneColorKeyColumn
        {
            get
            {
                return _phoneColorKeyColumn;
            }
            set
            {
                _phoneColorKeyColumn = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => PhoneColorKeyColumn);
            }
        }
        public string GroupByNameColumn
        {
            get
            {
                return _groupByNameColumn;
            }
            set
            {
                _groupByNameColumn = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => GroupByNameColumn);
            }
        }
        public string AgentNameColumn
        {
            get
            {
                return _agentNameColumn;
            }
            set
            {
                _agentNameColumn = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => AgentNameColumn);

            }
        }
        public string TwelveAmColumn
        {
            get
            {
                return _twelveAmColumn;
            }
            set
            {
                _twelveAmColumn = value;
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => TwelveAmColumn);
            }
        }

        public BindableCollection<TimeZoneInfo> TimeZoneInfos { get; set; }

        public bool CanSave
        {
            get
            {
                return HaveSettingsChanged();
            }
        }
        public bool CanCancel
        {
            get
            {
                return HaveSettingsChanged();
            }
        }


        private bool HaveSettingsChanged()
        {
            bool hasChanges = false;
            if (Properties.Settings.Default.TimeZoneId != SelectedTimeZone.Id ||
            Properties.Settings.Default.PhoneColorKeyColumn != PhoneColorKeyColumn ||
            Properties.Settings.Default.PhoneColorKeyRow != PhoneColorKeyRow ||
            Properties.Settings.Default.GroupByNameColumn != GroupByNameColumn ||
            Properties.Settings.Default.GroupByNameRow != GroupByNameRow ||
            Properties.Settings.Default.AgentNameColumn != AgentNameColumn ||
            Properties.Settings.Default.TwelveAmColumn != TwelveAmColumn)
            {
                hasChanges = true;
            }
            return hasChanges;
        }


        public SettingsViewModel()
        {
            TimeZoneInfos = GetTimeZoneInfoList();
            SelectedTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Properties.Settings.Default.TimeZoneId);
            PhoneColorKeyRow = Properties.Settings.Default.PhoneColorKeyRow;
            PhoneColorKeyColumn = Properties.Settings.Default.PhoneColorKeyColumn;
            GroupByNameRow = Properties.Settings.Default.GroupByNameRow;
            GroupByNameColumn = Properties.Settings.Default.GroupByNameColumn;
            AgentNameColumn = Properties.Settings.Default.AgentNameColumn;
            TwelveAmColumn = Properties.Settings.Default.TwelveAmColumn;
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

        private BindableCollection<TimeZoneInfo> GetTimeZoneInfoList()
        {
            BindableCollection<TimeZoneInfo> timeZoneInfos = new BindableCollection<TimeZoneInfo>();

            foreach (TimeZoneInfo tzi in TimeZoneInfo.GetSystemTimeZones().ToList())
            {
                timeZoneInfos.Add(tzi);
            }
            return timeZoneInfos;
        }

        public void Cancel()
        {
            SelectedTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Properties.Settings.Default.TimeZoneId);
            PhoneColorKeyRow = Properties.Settings.Default.PhoneColorKeyRow;
            PhoneColorKeyColumn = Properties.Settings.Default.PhoneColorKeyColumn;
            GroupByNameRow = Properties.Settings.Default.GroupByNameRow;
            GroupByNameColumn = Properties.Settings.Default.GroupByNameColumn;
            AgentNameColumn = Properties.Settings.Default.AgentNameColumn;
            TwelveAmColumn = Properties.Settings.Default.TwelveAmColumn;
            Properties.Settings.Default.Save();
            NotifyOfPropertyChange(() => CanSave);
            NotifyOfPropertyChange(() => CanCancel);
        }




        public void Back()
        {
            ActivateWindow.ViewReports();
        }
        public void Save()
        {
            Properties.Settings.Default.TimeZoneId = SelectedTimeZone.Id;
            Properties.Settings.Default.PhoneColorKeyColumn = PhoneColorKeyColumn;
            Properties.Settings.Default.PhoneColorKeyRow = PhoneColorKeyRow;
            Properties.Settings.Default.GroupByNameColumn = GroupByNameColumn;
            Properties.Settings.Default.GroupByNameRow = GroupByNameRow;
            Properties.Settings.Default.AgentNameColumn = AgentNameColumn;
            Properties.Settings.Default.TwelveAmColumn = TwelveAmColumn;
            Properties.Settings.Default.Save();
            NotifyOfPropertyChange(() => CanSave);
            NotifyOfPropertyChange(() => CanCancel);
        }
    }
}
