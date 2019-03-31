using Caliburn.Micro;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TalkdeskReportGenerator.Library;

namespace TalkdeskReportGenerator.ViewModels
{
    public class UserMappingViewModel : Screen
    {
        private BindableCollection<UserMapping> _mappings = new BindableCollection<UserMapping>();



        public BindableCollection<UserMapping> Mappings
        {
            get => _mappings;
            set
            {
                _mappings = value;
                NotifyOfPropertyChange(() => Mappings);
            }
        }



        public UserMappingViewModel()
        {
            List<UserMapping> userMappingList = new List<UserMapping>();

            if (Properties.Settings.Default.UserMappings == null)
            {
                Properties.Settings.Default.UserMappings = userMappingList;
            }
            else
            {
                userMappingList = Properties.Settings.Default.UserMappings;
            }

            Mappings = new BindableCollection<UserMapping>(userMappingList);

            Properties.Settings.Default.UserMappings = Mappings.ToList();
        }


        public void Delete(UserMapping mapping)
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
            ActivateWindow.AddUserMapping();
        }

        public void Edit(UserMapping mapping)
        {
            ActivateWindow.EditUserMapping(mapping);
        }
    }
}
