using System;
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
                var reportGnerator = new ReportCreator(Url, PersonalToken, InputPath, ResultPath);
                reportGnerator.CreateEstimateForExistingWorkItems();
                MessageBox.Show("Report Creation finished","Report Creation", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Report Creation failed",MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
