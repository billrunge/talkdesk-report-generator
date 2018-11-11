using Caliburn.Micro;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfTalkdeskReportGenerator.Models;

namespace WpfTalkdeskReportGenerator.ViewModels
{
    public class ShellViewModel : Screen
    {
        private string _excelPath;
        private string _excelPathStatus;
        private string _outputPath;
        private string _outputPathStatus;

        public string ExcelPath
        {
            get { return _excelPath; }
            set
            {
                _excelPath = value;
                SetExcelPathStatus();
                NotifyOfPropertyChange(() => ExcelPath);
                NotifyOfPropertyChange(() => CanBegin);
            }
        }
        public string ExcelPathStatus
        {
            get { return _excelPathStatus; }
            set
            {
                _excelPathStatus = value;
                NotifyOfPropertyChange(() => ExcelPathStatus);
            }
        }
        public string OutputPath
        {
            get { return _outputPath; }
            set
            {
                _outputPath = value;
                SetOutputPathStatus();
                NotifyOfPropertyChange(() => OutputPath);
            }

        }
        public string OutputPathStatus
        {
            get { return _outputPathStatus;  }
            set
            {
                _outputPathStatus = value;
                NotifyOfPropertyChange(() => OutputPathStatus);
                NotifyOfPropertyChange(() => CanBegin);
            }
        }
        public bool CanBegin
        {
            get
            {
                return (string.IsNullOrWhiteSpace(ExcelPath) || string.IsNullOrWhiteSpace(OutputPath)) ? false : true;
            }
        }

        public ShellViewModel()
        {
            SetExcelPathStatus();
            SetOutputPathStatus();                       
        }

        public void SetExcelPath()
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Title = "Open Schedule Excel",
                Filter = "Excel Files|*.xlsx",
                InitialDirectory = @"C:\"
            };

            if (fileDialog.ShowDialog() == true)
            {
                ExcelPath = fileDialog.FileName.ToString();
            }
        }

        public void SetExcelPathStatus()
        {
            if (ExcelPath == null)
            {
                ExcelPathStatus = "✖";
            }
            else
            {
                ExcelPathStatus = "✔";
            }
        }

        public void SetOutputPath()
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog()
            {
                Description = "Select Output Folder",
                ShowNewFolderButton = true
            };
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputPath = folderBrowser.SelectedPath + @"\";
            }
        }

        public void SetOutputPathStatus()
        {
            if (OutputPath == null)
            {
                OutputPathStatus = "✖";
            }
            else
            {
                OutputPathStatus = "✔";
            }
        }

        public void Begin()
        {
            IDatabase db = new Database();
            IGetStatuses getStatuses = new GetStatuses(db);
            IGetAgentTimes getAgentTimes = new GetAgentTimesFromExcel();



        }
    }
}
