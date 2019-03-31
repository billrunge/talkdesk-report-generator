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
        private AgentMapping _editMapping;

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
                return (ExcelName == _editMapping.ExcelAgentName && TalkdeskName == _editMapping.TalkdeskAgentName) ? false : true;
            }
        }

        public bool CanCancel
        {
            get
            {
                return (ExcelName == _editMapping.ExcelAgentName && TalkdeskName == _editMapping.TalkdeskAgentName) ? false : true;
            }
        }

        public bool CanBack
        {
            get
            {
                return (ExcelName == _editMapping.ExcelAgentName && TalkdeskName == _editMapping.TalkdeskAgentName) ? true : false;
            }
        }


        public EditUserMappingViewModel()
        {
            _newMapping = true;
            _editMapping = new AgentMapping();
        }

        public EditUserMappingViewModel(AgentMapping mapping)
        {
            _editMapping = mapping;
            _newMapping = false;
            ExcelName = mapping.ExcelAgentName;
            TalkdeskName = mapping.TalkdeskAgentName;
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
                List<AgentMapping> userMappings = Properties.Settings.Default.UserMappings ?? new List<AgentMapping>();
                if (!_newMapping)
                {
                    userMappings.Remove(_editMapping);
                }

                AgentMapping userMapping = new AgentMapping()
                {
                    ExcelAgentName = ExcelName,
                    TalkdeskAgentName = TalkdeskName
                };

                userMappings.Add(userMapping);
                Properties.Settings.Default.UserMappings = userMappings;
                _editMapping = userMapping;
                ActivateWindow.ViewUserMapping();
            }            
        }

        public void Cancel()
        {
            ExcelName = _editMapping.ExcelAgentName;
            TalkdeskName = _editMapping.TalkdeskAgentName;
        }


    }
}
