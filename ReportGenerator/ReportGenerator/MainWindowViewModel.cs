using Core;
using System.Windows.Input;

namespace ReportGenerator
{
    public class MainWindowViewModel
    {
        public ICommand GenerateReportCommand { get; set; }

        public bool IsBusy;

        public MainWindowViewModel()
        {
            GenerateReportCommand = new RelayCommand((x) => !IsBusy, GenerateReport);
        }

        public void GenerateReport(object input)
        {
            var reportGnerator = new ReportCreator();
            reportGnerator.CreateEstimateForExistingWorkItems();
        }
    }
}
