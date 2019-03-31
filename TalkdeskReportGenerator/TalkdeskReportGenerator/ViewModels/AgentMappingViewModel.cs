using Caliburn.Micro;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TalkdeskReportGenerator.Library;

namespace TalkdeskReportGenerator.ViewModels
{
    public class AgentMappingViewModel : Screen
    {
        private BindableCollection<AgentMapping> _mappings = new BindableCollection<AgentMapping>();



        public BindableCollection<AgentMapping> Mappings
        {
            get => _mappings;
            set
            {
                _mappings = value;
                NotifyOfPropertyChange(() => Mappings);
            }
        }



        public AgentMappingViewModel()
        {
            List<AgentMapping> userMappingList = new List<AgentMapping>();

            if (Properties.Settings.Default.UserMappings == null)
            {
                Properties.Settings.Default.UserMappings = userMappingList;
            }
            else
            {
                userMappingList = Properties.Settings.Default.UserMappings;
            }

            Mappings = new BindableCollection<AgentMapping>(userMappingList);

            Properties.Settings.Default.UserMappings = Mappings.ToList();
        }


        public void Delete(AgentMapping mapping)
        {
            MessageBoxResult result = MessageBox.Show("Are you sure?", "Delete Mapping", MessageBoxButton.YesNo);
            if (result == MessageBoxResult.Yes)
            {
                Mappings.Remove(mapping);
                Properties.Settings.Default.UserMappings = Mappings.ToList();
            }
        }

        public void Back()
        {
            ActivateWindow.ViewSettings();
        }

        public void Add()
        {
            ActivateWindow.AddAgentMapping();
        }

        public void Edit(AgentMapping mapping)
        {
            ActivateWindow.EditAgentMapping(mapping);
        }
    }
}
