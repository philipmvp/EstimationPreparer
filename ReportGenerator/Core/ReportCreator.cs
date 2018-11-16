using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using ClosedXML.Excel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.TeamFoundation.Build.WebApi;

namespace Core
{
    public class ReportCreator
    {
        private readonly string _uri;
        private readonly string _personalAccessToken;
        private readonly string _inputPath;
        private readonly string _outputPath;
        private readonly string _project = "India";
        private readonly Regex _backlogNuberPattern = new Regex(@"#+([0-9]+)", RegexOptions.Compiled);
        public StringBuilder Logs { get; internal set; }

        public ReportCreator(string url, string personalToken, string inputPath, string outputPath)
        {
            _uri = url;
            _personalAccessToken = personalToken;
            _inputPath = inputPath;
            _outputPath = outputPath;
            Logs = new StringBuilder();
        }

        public void CreateEstimateForExistingWorkItems()
        {
            var workItems = GetAllLiveWorkItemsFromTeamServer();
            var actualTimeFromWorksheets = GetActualTime(workItems.Keys.Select(x => x.Id), _inputPath);
            CreateEstimationExcel(workItems, actualTimeFromWorksheets);
        }

        public List<DeveloperEffortDetails> GetEffortsForMonth(int month, int year, string perfRecordsPath)
        {
            var workItemEfforts = new List<DeveloperEffortDetails>();
            try
            {
                Logs.AppendLine($"{DateTime.Now} : Starting reading of performance sheet");
                var expectedWorkingDays = GetExpectedWorkingDaysForMonth(month, year);
                var expectedWorkingHoursForMonth = expectedWorkingDays.Count * 8;
                foreach (var perfRecordPath in Directory.GetFiles(perfRecordsPath, "*.xlsx"))
                {
                    using (var excel = new XLWorkbook(perfRecordPath))
                    {
                        var fileName = Path.GetFileNameWithoutExtension(perfRecordPath);
                        Logs.AppendLine($"Reading file {fileName}");
                        var effortDetails = new DeveloperEffortDetails() { DeveloperName = Path.GetFileNameWithoutExtension(perfRecordPath).Replace("Work performance record ", string.Empty) };
                        var sheetName = $"{month}_{year.ToString().Remove(0, 2)}";
                        var currentMonthWorkSheet = excel.Worksheets.FirstOrDefault(x => x.Name == sheetName);

                        if (currentMonthWorkSheet == null)
                        {
                            Logs.AppendLine($" Sheet {sheetName} not found in {fileName}. Excluding this file");
                            continue;
                        }

                        if (!currentMonthWorkSheet.FirstRow().Cell(1).TryGetValue(out int actualWorkingHours) || actualWorkingHours != expectedWorkingHoursForMonth)
                            Logs.AppendLine($"Total working hours in sheet {sheetName} of {fileName} is not correct. Expected value is {expectedWorkingHoursForMonth}");

                        var hoursInDayDictionary = new Dictionary<DateTime, float>();
                        foreach (var row in currentMonthWorkSheet.RowsUsed())
                        {
                            if (row.Cell(1).TryGetValue(out DateTime date) == false)
                            {
                                Logs.AppendLine($"First column of row {row.RowNumber()} with value {row.Cell(1).Value} is not in Date format. Excluding this row from effort calculation");
                                continue;
                            }

                            var value = row.CellsUsed();
                            if (value == null || value.Count() < 3)
                            {
                                Logs.AppendLine($"No hours added in row {row.RowNumber()} ");
                                continue;
                            }

                            var effortValue = value.ElementAt(2).GetValue<float>();

                            if (hoursInDayDictionary.ContainsKey(date) == false)
                                hoursInDayDictionary.Add(date, effortValue);
                            else
                                hoursInDayDictionary[date] += effortValue;

                            var intValueMatches = _backlogNuberPattern.Matches(row.Cell(2).GetValue<string>());
                            string usedWorkitemId = null;
                            foreach (Match match in intValueMatches)
                            {
                                usedWorkitemId = match.Groups[1].Value;
                            }

                            if (usedWorkitemId != null)
                            {
                                if (effortDetails.BacklogEfforts.ContainsKey(usedWorkitemId) == false)
                                    effortDetails.BacklogEfforts.Add(usedWorkitemId, effortValue);
                                else
                                    effortDetails.BacklogEfforts[usedWorkitemId] += effortValue;
                            }
                        }

                        ValidateHoursInDays(hoursInDayDictionary);
                        workItemEfforts.Add(effortDetails);
                    }
                }

                Logs.AppendLine($"{DateTime.Now} : Finished reading of performance sheet");
            }
            catch (Exception e)
            {
                Logs.AppendLine($"{DateTime.Now} : Finished reading of performance sheet with exception:");
                Logs.AppendLine(e.Message);
                Logs.AppendLine(e.StackTrace);
                throw;
            }

            return workItemEfforts;
        }

        private void ValidateHoursInDays(Dictionary<DateTime, float> hoursInDayDictionary)
        {
            foreach (var dayEffort in hoursInDayDictionary.Where(x => x.Value != 8))
            {
                Logs.AppendLine($"Total effort for day {dayEffort.Key} is {dayEffort.Value} and not 8. Please update.");
            }
        }

        private List<DateTime> GetExpectedWorkingDaysForMonth(int month, int year)
        {
            var workingDayList = new List<DateTime>();

            for (int i = 1; i <= DateTime.DaysInMonth(year, month); i++)
            {
                var currentDate = new DateTime(year, month, i);
                if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
                    workingDayList.Add(currentDate);
            }

            return workingDayList;
        }

        private void CreateEstimationExcel(Dictionary<WorkItem, double> workItems, Dictionary<int?, List<float>> actualTimeFromWorksheets)
        {
            using (var excel = new XLWorkbook())
            {
                using (var worksheet = excel.Worksheets.Add("Estimation"))
                {
                    worksheet.Row(1).Cell(1).Value = "WorkItemId";
                    worksheet.Row(1).Cell(2).Value = "Estimated Effort";
                    worksheet.Row(1).Cell(3).Value = "Actual Effort";
                    worksheet.Row(1).Cell(4).Value = "Remaining Effort";

                    var count = 2;
                    foreach (var workItem in workItems)
                    {
                        var estimatedEffort = workItem.Key.Fields.TryGetValue("Microsoft.VSTS.Scheduling.Effort", out double effort) ? effort : 0;
                        var actualEffort = actualTimeFromWorksheets.TryGetValue(workItem.Key.Id, out List<float> times) ? times.Sum() : 0;
                        if (estimatedEffort == 0 && actualEffort == 0)
                            continue;

                        worksheet.Row(count).Cell(1).Value = workItem.Key.Id + " (" + workItem.Key.Fields["System.Title"] + ")";
                        worksheet.Row(count).Cell(2).Value = estimatedEffort;
                        worksheet.Row(count).Cell(3).Value = actualEffort;
                        worksheet.Row(count).Cell(4).Value = workItem.Value;
                        count++;
                    }

                    excel.SaveAs(_outputPath);
                }
            }
        }

        private Dictionary<int?, List<float>> GetActualTime(IEnumerable<int?> workItemIdList, string perfRecordsPath)
        {
            var workItemEfforts = new Dictionary<int?, List<float>>();
            foreach (var perfRecordPath in Directory.GetFiles(perfRecordsPath, "*.xlsx"))
            {
                using (var excel = new XLWorkbook(perfRecordPath))
                {
                    foreach (var worksheet in excel.Worksheets)
                    {
                        foreach (var cell in worksheet.Column(2).CellsUsed())
                        {
                            var intValueMatches = Regex.Matches(cell.GetValue<string>(), "[0-9]+");
                            int? usedWorkitemId = null;
                            foreach (Match match in intValueMatches)
                            {
                                if (Int32.TryParse(match.Value, out int value))
                                    usedWorkitemId = workItemIdList.FirstOrDefault(x => x == value);
                            }

                            if (usedWorkitemId != null)
                            {
                                var value = cell.WorksheetRow().CellsUsed();
                                if (value == null || value.Count() < 3)
                                    continue;

                                if (workItemEfforts.ContainsKey(usedWorkitemId) == false)
                                    workItemEfforts.Add(usedWorkitemId, new List<float>() { value.ElementAt(2).GetValue<float>() });
                                else
                                    workItemEfforts[usedWorkitemId].Add(value.ElementAt(2).GetValue<float>());
                            }
                        }
                    }
                }
            }

            return workItemEfforts;
        }

        /// <summary>
        /// Execute a WIQL query to return a list of bugs using the .NET client library
        /// </summary>
        /// <returns>List of Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.WorkItem</returns>
        private Dictionary<WorkItem, double> GetAllLiveWorkItemsFromTeamServer()
        {
            Uri uri = new Uri(_uri);
            string personalAccessToken = _personalAccessToken;
            string project = _project;

            VssBasicCredential credentials = new VssBasicCredential("", _personalAccessToken);

            var workItemDetails = new Dictionary<WorkItem, double>();

            //create a wiql object and build our query
            Wiql wiql = new Wiql()
            {
                Query = "Select [State], [Title] ,[Effort], [Remaining Work]" +
                        "From WorkItems " +
                        "Where ([Work Item Type] = 'Bug' " +
                        "Or [Work Item Type] = 'Product Backlog Item')" +
                        "And [System.TeamProject] = '" + project + "' " +
                        "And ([State] = 'New' " +
                        "Or [State] = 'Approved'" +
                        "Or [State] = 'Committed')" +
                        "Order By [Priority] Asc, [Changed Date] Desc"
            };

            //create instance of work item tracking http client
            using (WorkItemTrackingHttpClient workItemTrackingHttpClient = new WorkItemTrackingHttpClient(uri, credentials))
            {
                //execute the query to get the list of work items in the results
                WorkItemQueryResult workItemQueryResult = workItemTrackingHttpClient.QueryByWiqlAsync(wiql).Result;

                //some error handling                
                if (workItemQueryResult.WorkItems.Count() != 0)
                {
                    //need to get the list of our work item ids and put them into an array
                    List<int> list = new List<int>();
                    foreach (var item in workItemQueryResult.WorkItems)
                    {
                        list.Add(item.Id);
                    }
                    int[] arr = list.ToArray();

                    //build a list of the fields we want to see
                    string[] fields = new string[5];
                    fields[0] = "System.Id";
                    fields[1] = "System.Title";
                    fields[2] = "System.State";
                    fields[3] = "Microsoft.VSTS.Scheduling.Effort";
                    fields[4] = "Microsoft.VSTS.Scheduling.RemainingWork";

                    //get work items for the ids found in query
                    var workItems = workItemTrackingHttpClient.GetWorkItemsAsync(arr, null, workItemQueryResult.AsOf, WorkItemExpand.All).Result;

                    Console.WriteLine("Query Results: {0} items found", workItems.Count);
                    list.Clear();

                    //loop though work items and write to console
                    foreach (var workItem in workItems)
                    {
                        var totalRemainingWork = 0.0;
                        if (workItem.Relations?.Any(x => x.Rel == "System.LinkTypes.Hierarchy-Forward") == true)
                        {
                            list.Clear();
                            foreach (var relation in workItem.Relations)
                            {
                                //get the child links
                                if (relation.Rel == "System.LinkTypes.Hierarchy-Forward")
                                {
                                    var lastIndex = relation.Url.LastIndexOf("/");
                                    var itemId = relation.Url.Substring(lastIndex + 1);
                                    list.Add(Convert.ToInt32(itemId));
                                };
                            }

                            int[] taskItemIds = list.ToArray();

                            var taskItems = workItemTrackingHttpClient.GetWorkItemsAsync(taskItemIds, new[] { "Microsoft.VSTS.Scheduling.RemainingWork" }).Result;

                            Console.WriteLine("Getting full work item for each child...");


                            foreach (var item in taskItems)
                            {
                                if (item.Fields.TryGetValue("Microsoft.VSTS.Scheduling.RemainingWork", out double remaningWork))
                                    totalRemainingWork += remaningWork;
                            }
                        }

                        workItemDetails.Add(workItem, totalRemainingWork);
                    }
                }

                return workItemDetails;
            }
        }

        public void CreateReportForMonth(int month, int year, List<DeveloperEffortDetails> effortDetails)
        {
            var combinedBacklogEfforts = new Dictionary<string, float>();
            foreach (var effortDetail in effortDetails.SelectMany(x => x.BacklogEfforts).GroupBy(x => x.Key))
            {
                combinedBacklogEfforts.Add(effortDetail.Key, effortDetail.Sum(x => x.Value));
            }

            var arr = combinedBacklogEfforts.Keys.Select(int.Parse).ToArray();

            Uri uri = new Uri(_uri);

            VssBasicCredential credentials = new VssBasicCredential("", _personalAccessToken);

            using (WorkItemTrackingHttpClient workItemTrackingHttpClient = new WorkItemTrackingHttpClient(uri, credentials))
            {
                //get work items for the ids found in query
                var workItems = workItemTrackingHttpClient.GetWorkItemsAsync(arr, null, null, WorkItemExpand.All).Result;
                using (var excel = new XLWorkbook())
                {
                    using (var worksheet = excel.Worksheets.Add("Estimation"))
                    {
                        worksheet.Row(1).Cell(1).Value = "WorkItemId";
                        worksheet.Row(1).Cell(2).Value = "State";
                        worksheet.Row(1).Cell(3).Value = "Estimated Effort";
                        worksheet.Row(1).Cell(4).Value = "Actual Effort";

                        var count = 2;
                        foreach (var workItem in workItems.OrderBy(x=> x.Fields["System.Title"].ToString()))
                        {
                            var workItemId = workItem.Id.ToString();
                            var title = workItem.Fields["System.Title"].ToString();
                            var state = workItem.Fields["System.State"].ToString();
                            var estimatedEffort = workItem.Fields.TryGetValue("Microsoft.VSTS.Scheduling.Effort", out double effort) ? effort : 0;
                            var actualEffort = combinedBacklogEfforts[workItemId];

                            worksheet.Row(count).Cell(1).Value = $"{workItemId} - {title}";
                            worksheet.Row(count).Cell(2).Value = state;
                            worksheet.Row(count).Cell(3).Value = estimatedEffort;
                            worksheet.Row(count).Cell(4).Value = actualEffort;
                            count++;
                        }
                    }
                    excel.SaveAs(_outputPath);
                }
            }
        }
    }

    public class DeveloperEffortDetails
    {
        public string DeveloperName { get; set; }
        public Dictionary<string, float> BacklogEfforts { get; set; } = new Dictionary<string, float>();
    }
}
