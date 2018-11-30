using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using Core;
using System.Windows.Input;
using ReportGenerator.Properties;

namespace ReportGenerator
{
    public class MainWindowViewModel
    {
        public ICommand GenerateReportCommand { get; set; }

        public string Url
        {
            get => Settings.Default.TeamUrl;
            set
            {
                Settings.Default.TeamUrl = value;
                Settings.Default.Save();
            }
        }

        public string PersonalToken
        {
            get => Settings.Default.PersonalToken;
            set
            {
                Settings.Default.PersonalToken = value;
                Settings.Default.Save();
            }
        }

        public string InputPath
        {
            get => Settings.Default.InputPath;
            set
            {
                Settings.Default.InputPath = value;
                Settings.Default.Save();
            }
        }

        public string ResultPath
        {
            get => Settings.Default.OutputPath;
            set
            {
                Settings.Default.OutputPath = value;
                Settings.Default.Save();
            }
        }

        public bool IsBusy;

        public MainWindowViewModel()
        {
            GenerateReportCommand = new RelayCommand(x => !IsBusy, GenerateReport);
        }

        public void GenerateReport(object input)
        {
            try
            {
                var reportGenerator = new ReportCreator(Url, PersonalToken, InputPath, ResultPath);
                //reportGnerator.CreateEstimateForExistingWorkItems();
                var result = reportGenerator.GetEffortsForMonth(Int32.Parse(ResultPath),2018);
                reportGenerator.CreateReportForMonth(result, $@"D:\PerformanceSheet\Report_{ResultPath}_2018.xlsx");
                using (var fileStream = File.Create($@"D:\PerformanceSheet\Log_{ResultPath}_2018.txt"))
                {
                    var info = new UTF8Encoding(true).GetBytes(reportGenerator.Logs.ToString());
                    fileStream.Write(info,0,info.Length);
                }
                MessageBox.Show("Report Creation finished", "Report Creation", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Report Creation failed",MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
