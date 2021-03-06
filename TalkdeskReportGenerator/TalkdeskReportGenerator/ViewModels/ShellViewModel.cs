﻿using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using TalkdeskReportGenerator.Library;

namespace TalkdeskReportGenerator.ViewModels
{
    public class ShellViewModel : Conductor<Object>
    {
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public ShellViewModel()
        {
            ActivateWindow.ShellView = this;
            ActivateWindow.ViewReports();
        }

        public async Task OnClose(CancelEventArgs args)
        {
            foreach (string tempExcelPath in Properties.Settings.Default.TemporaryExcelPaths)
            {
                if (!string.IsNullOrWhiteSpace(tempExcelPath))
                {
                    _log.Info($"ShellViewModel.OnClose - Deleting the temporary file: { tempExcelPath }");
                    ExcelReader excelReader = new ExcelReader(_log);
                    await excelReader.DeleteExcelAsync(tempExcelPath);
                }
            }
            Properties.Settings.Default.TemporaryExcelPaths = new List<string>();
            Properties.Settings.Default.Save();
        }
    }

    public static class ActivateWindow
    {
        public static ShellViewModel ShellView;
        public static ReportsViewModel ReportsView = new ReportsViewModel();
        public static SettingsViewModel SettingsView = new SettingsViewModel();

        public static EditAgentMappingViewModel EditUserMappingView = new EditAgentMappingViewModel();

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

        public static void ViewAgentMapping()
        {
        AgentMappingViewModel userMappingView = new AgentMappingViewModel();
        OpenItem(userMappingView);
        }

        public static void AddAgentMapping()
        {
            EditAgentMappingViewModel editUserMappingView = new EditAgentMappingViewModel();
            OpenItem(editUserMappingView);
        }

        public static void EditAgentMapping(AgentMapping mapping)
        {
            EditAgentMappingViewModel editUserMappingView = new EditAgentMappingViewModel(mapping);
            OpenItem(editUserMappingView);
        }
    }
}
