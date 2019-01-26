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
            ActivateWindow.Parent = this;
            ActivateWindow.OpenItem(new ReportsViewModel());
        }
    }

    public static class ActivateWindow
    {
        public static ShellViewModel Parent;

        public static void OpenItem(IScreen t)
        {
            Parent.ActivateItem(t);
        }
    }
}
