using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using ClosedXML.Excel;
using System.IO;

namespace Core
{
    public class ReportCreator
    {
        private string _uri;
        private string _personalAccessToken;
        private string _project;

        public ReportCreator()
        {
            _uri = "https://medapro.visualstudio.com";
            _personalAccessToken = "23m3msqzrvvma7kzeqsozdiyqubtghpm7uz5wrzzm63bbzq4s5fa";
            _project = "India";
        }

        public void CreateEstimateForExistingWorkItems()
        {
            var workItems = GetAllLiveWorkItemsFromTeamServer();
            var actualTimeFromWorksheets = GetActualTime(workItems.Keys.Select(x => x.Id), @"G:\PerformanceRecords");
            CreateEstimationExcel(workItems, actualTimeFromWorksheets);
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
                    foreach(var workItem in workItems)
                    {
                        worksheet.Row(count).Cell(1).Value = workItem.Key.Id;
                        worksheet.Row(count).Cell(2).Value = workItem.Key.Fields.TryGetValue("Microsoft.VSTS.Scheduling.Effort", out double effort)?effort:0;
                        worksheet.Row(count).Cell(3).Value = actualTimeFromWorksheets.TryGetValue(workItem.Key.Id, out List<float> times) ? times.Sum(): 0;
                        worksheet.Row(count).Cell(4).Value = workItem.Value;
                        count++;
                    }

                    excel.SaveAs(@"G:\Estimation\SprintEstimation.xlsx");
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
                    foreach(var worksheet in excel.Worksheets)
                    {
                        foreach(var cell in worksheet.Column(2).CellsUsed())
                        {
                            var usedWorkitemId = workItemIdList.FirstOrDefault(x => cell.GetValue<string>().Contains(x.ToString()));
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
        private Dictionary<WorkItem,double> GetAllLiveWorkItemsFromTeamServer()
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

            //private List<WorkItem> GetChildWorkItemsByParent()
            //{
            //    int id = Convert.ToInt32(Context.GetValue<WorkItem>("$newWorkItem1").Id);

            //    VssConnection connection = Context.Connection;
            //    WorkItemTrackingHttpClient workItemTrackingClient = connection.GetClient<WorkItemTrackingHttpClient>();

            //    WorkItem workitem = workItemTrackingClient.GetWorkItemAsync(id, expand: WorkItemExpand.Relations).Result;
            //    List<WorkItem> workitems = null;

            //    if (workitem.Relations == null)
            //    {
            //        Console.WriteLine("  No relations found for this work item");
            //    }
            //    else
            //    {
            //        List<int> list = new List<int>();

            //        Console.WriteLine("Getting child work items from parent...");

            //        foreach (var relation in workitem.Relations)
            //        {
            //            //get the child links
            //            if (relation.Rel == "System.LinkTypes.Hierarchy-Forward")
            //            {
            //                var lastIndex = relation.Url.LastIndexOf("/");
            //                var itemId = relation.Url.Substring(lastIndex + 1);
            //                list.Add(Convert.ToInt32(itemId));

            //                Console.WriteLine("  {0} ", itemId);
            //            };
            //        }

            //        int[] workitemIds = list.ToArray();

            //        workitems = workItemTrackingClient.GetWorkItemsAsync(workitemIds).Result;

            //        Console.WriteLine("Getting full work item for each child...");

            //        foreach (var item in workitems)
            //        {
            //            Console.WriteLine(" {0}: {1} : {2}", item.Fields["System.WorkItemType"], item.Id, item.Fields["System.Title"]);
            //        }
            //    }

            //    return workitems;
            //}
        }
    }
}
