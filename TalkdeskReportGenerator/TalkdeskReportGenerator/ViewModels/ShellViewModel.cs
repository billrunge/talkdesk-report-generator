using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfTalkdeskReportGenerator.ViewModels
{
    public class ShellViewModel : Conductor<Object>
    {
                     
        public ShellViewModel()
        {
            ActivateWindow.ShellView = this;
            ActivateWindow.ViewReports();
        }

    }

    public static class ActivateWindow
    {
        public static ShellViewModel ShellView;
        public static ReportsViewModel ReportsView = new ReportsViewModel();
        public static SettingsViewModel SettingsView = new SettingsViewModel();

        public static void OpenItem(IScreen screen)
        {
            ShellView.ActivateItem(screen);
        }

        public static void ViewReports()
        {
            OpenItem(ReportsView);
        }

        public static void ViewSettings()
        {
            OpenItem(SettingsView);
        }
    }
}
