using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TalkdeskReportGenerator.Library;

namespace TalkdeskReportGenerator.ViewModels
{
    public class EditUserMappingViewModel : Screen
    {
        private string _excelName;
        private string _talkdeskName;
        private bool _newMapping;
        private UserMapping _editMapping;

        public string ExcelName
        {
            get
            {
                return _excelName;
            }
            set
            {
                _excelName = value;
                NotifyOfPropertyChange(() => ExcelName);
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => CanBack);
            }
        }
        public string TalkdeskName
        {
            get
            {
                return _talkdeskName;
            }
            set
            {
                _talkdeskName = value;
                NotifyOfPropertyChange(() => TalkdeskName);
                NotifyOfPropertyChange(() => CanSave);
                NotifyOfPropertyChange(() => CanCancel);
                NotifyOfPropertyChange(() => CanBack);
            }
        }

        public bool CanSave
        {
            get
            {
                return (ExcelName == _editMapping.ExcelUser && TalkdeskName == _editMapping.TalkdeskUser) ? false : true;
            }
        }

        public bool CanCancel
        {
            get
            {
                return (ExcelName == _editMapping.ExcelUser && TalkdeskName == _editMapping.TalkdeskUser) ? false : true;
            }
        }

        public bool CanBack
        {
            get
            {
                return (ExcelName == _editMapping.ExcelUser && TalkdeskName == _editMapping.TalkdeskUser) ? true : false;
            }
        }


        public EditUserMappingViewModel()
        {
            _newMapping = true;
            _editMapping = new UserMapping();
        }

        public EditUserMappingViewModel(UserMapping mapping)
        {
            _editMapping = mapping;
            _newMapping = false;
            ExcelName = mapping.ExcelUser;
            TalkdeskName = mapping.TalkdeskUser;
        }


        public void Back()
        {
            ActivateWindow.ViewUserMapping();
        }

        public void Save()
        {
            if (string.IsNullOrWhiteSpace(ExcelName))
            {
                MessageBox.Show("Please enter an Excel Name");
            } else if (string.IsNullOrWhiteSpace(TalkdeskName))
            {
                MessageBox.Show("Please enter a Talkdesk Name");
            }
            else
            {
                List<UserMapping> userMappings = Properties.Settings.Default.UserMappings ?? new List<UserMapping>();
                if (!_newMapping)
                {
                    userMappings.Remove(_editMapping);
                }

                UserMapping userMapping = new UserMapping()
                {
                    ExcelUser = ExcelName,
                    TalkdeskUser = TalkdeskName
                };

                userMappings.Add(userMapping);
                Properties.Settings.Default.UserMappings = userMappings;
                ActivateWindow.ViewUserMapping();
            }            
        }

        public void Cancel()
        {
            ExcelName = _editMapping.ExcelUser;
            TalkdeskName = _editMapping.TalkdeskUser;
        }


    }
}
