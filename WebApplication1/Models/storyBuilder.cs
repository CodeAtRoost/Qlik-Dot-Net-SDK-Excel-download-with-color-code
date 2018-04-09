using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Qlik.Engine;
using Qlik.Engine.Communication;
using Qlik.Sense.Client;
using Qlik.Sense.Client.Snapshot;
using Qlik.Sense.Client.Storytelling;
using Qlik.Sense.Client.Visualizations;
using Qlik.Sense.Client.Visualizations.Components;

namespace QlikService.Models
{
    public class storyBuilder
    {
        public static string  buildStory()
        {
            var location = Qlik.Engine.Location.FromUri(new Uri("ws://localhost:4848"));
            try
            {
                // Connect to desktop
                
                location.AsDirectConnectionToPersonalEdition();
            }
            catch(SystemException e)
            {
                Console.WriteLine("Could not open app! " + e.ToString());
                return "Connection to Qlik engine failed";

            }
          //  location.AsNtlmUserViaProxy(proxyUsesSsl: false);
            try
            {
                // Open the app with name "Beginner's tutorial"
                var appIdentifier = location.AppWithNameOrDefault(@"Consumer Sales", noVersionCheck: true);
                using (var app = location.App(appIdentifier, noVersionCheck: true))
                {
                    // Clear all selections to set the app in an known state
                    app.ClearAll();

                    // Get the sheet with the title "Dashboard"
                    var sheet = GetSheetWithTitle(app, "KPI Dashboard");

                    //Take snapshots
                    ISnapshot marginAmountOverTime;

                    TakeSnapshots(sheet, app, out marginAmountOverTime);

                    // Create a new story
                    var storyProps = new StoryProperties();
                    // Enter the title
                    storyProps.MetaDef.Title = "SDK Creted - Margin Over Time";
                    var story = app.CreateStory("marginovertime", storyProps);
                    // Create a slide
                    var slideProps = new SlideProperties();
                    var slide1 = story.CreateSlide("Margin Trend", slideProps);
                    
                    // Add a title
                    var titleProp = slide1.CreateTextSlideItemProperties("margintrend", Slide.TextType.Title, text: "Margin Trend");
                    titleProp.Position = new SlidePosition
                    {
                        Height = "20%",
                        Left = "5%",
                        Top = "0.1%",
                        Width = "40%",
                        ZIndex = 1
                    };
                    slide1.CreateSlideItem(null, titleProp);

                    // Add the Sales per Region snapshot to the slide
                    // Resize (The .NET SDK sets the size)
                    AddSnapshotToSlide(slide1, "marginAmountOverTime", "SDK_MarginOverTime", "Usa", "1%");


                    // Create slides 2-4

                    // Clear all selections.
                    app.ClearAll();
                    // Save the new story and the snaphots
                    app.DoSave();                    
                    Console.WriteLine(@"A new story by the name 'SDK Created - Margin Over Time' has been created. Open 'Beginner's tutorial' and verify your new story.");
                    return "Story create successfully";
                }
            }
            catch (SystemException e)
            {
                Console.WriteLine("Could not open app! " + e.ToString());
                return "Could not open the app";
            }
           /* catch (TimeoutException e)
            {
                Console.WriteLine("Timeout : " + e.Message);
            }*/
            Console.ReadLine();
        }

        private static void TakeSnapshots(ISheet sheet, IApp app, out ISnapshot marginAmountOverTime)
        {
            // Get the Sales per Region on the sheet
            var marginovertime = GetVisualisationWithTitle(sheet, "Margin Amount Over Time");
            // Take a snapshot of SalesPerRegion
            marginAmountOverTime = null;
            if (marginovertime != null)
                marginAmountOverTime = app.CreateSnapshot("SDK_MarginOverTime", sheet.Id, marginovertime.Id);


        }

        private static void CreateRegionItemsOnSlide(ISlide slide, string title, ISnapshot top5Customers, ISnapshot quarterlyTrend)
        {
            var titleProp = slide.CreateTextSlideItemProperties(null, Slide.TextType.Title, text: title);
            titleProp.Position = new SlidePosition { Height = "20%", Left = "5%", Top = "0.1%", Width = "40%", ZIndex = 1 };
            slide.CreateSlideItem(null, titleProp);
            AddSnapshotToSlide(slide, "Top5Slide", top5Customers.Id, null, "1%");
            AddSnapshotToSlide(slide, "QuarterlySlite", quarterlyTrend.Id, null, "51%");
        }

        private static void AddSnapshotToSlide(ISlide slide, string name, string snapshotId, string region, string left)
        {
            int selectedIndex = 0;
            var prop = slide.CreateSnapshotSlideItemProperties(name, snapshotId);
            prop.Position = new SlidePosition
            {
                Height = "33%",
                Left = left,
                Top = "25%",
                Width = "33%",
                ZIndex = 1
            };
            var slideItem = slide.CreateSlideItem(name, prop);
            slideItem.EmbedSnapshotObject(snapshotId);
            var mySnapShotedItem = slideItem.GetSnapshotObject();

            var data = mySnapShotedItem.GetProperties();
            if (region != null)
            {
                var hypercube = data.Get<HyperCube>("qHyperCube");
                foreach (var nxDataPage in hypercube.DataPages)
                {
                    int index = -1;
                    foreach (var cellRows in nxDataPage.Matrix)
                    {
                        index++;
                        foreach (var row in cellRows)
                        {
                            if (row.Text == region)
                                selectedIndex = index;
                        }
                    }
                }
            }
            using (mySnapShotedItem.SuspendedLayout)
            {
                JObject originalModelSettings = new JObject();
                JObject datapoint = new JObject();
                JObject legend = new JObject();
                datapoint.Add("auto", false);
                datapoint.Add("labelmode", "share");
                legend.Add("show", false);
                legend.Add("dock", "auto");
                legend.Add("showTitle", true);
                originalModelSettings.Add("dataPoint", datapoint);
                originalModelSettings.Add("legend", legend);
                originalModelSettings.Add("dimensionTitle", true);
                data.Set("originalModelSettings", originalModelSettings);
                if (region != null)
                {
                    data.Set("effectPath", "/effects/highlight-value");
                    JObject effectProperties = new JObject();
                    effectProperties.Add("selectedIndex", selectedIndex);

                    data.Set("effectProperties", effectProperties);
                }
            }
        }

        private static int GetSelectAndCreateSnapshots(IApp app, ISheet sheet, string region, out ISnapshot top5CustomersSnapshot, out ISnapshot quarterlyTrendSnapshot)
        {
            // Select region
            int selectedIndex = 0;
            IField field = null;
            foreach (var item in app.GetFieldList().Items)
            {
                if (item.Name == "Region")
                    field = app.GetField(item.Name);
            }
            var res = field != null && field.Select(region);
            var extendedSel = app.GetExtendedCurrentSelection();
            foreach (var nxDataPage in extendedSel.GetField("Region").DataPages)
            {
                foreach (var rows in nxDataPage.Matrix)
                {
                    var cell = rows.FirstOrDefault();
                    if (cell != null)
                    {
                        if (cell.State == StateEnumType.SELECTED)
                            selectedIndex = cell.ElemNumber;
                    }
                }
            }

            // Get the Top 5 Customers on the sheet
            var top5Customers = GetVisualisationWithTitle(sheet, "Top 5 Customers");

            // Take a snapshot of Top 5 Customers
            top5CustomersSnapshot = app.CreateSnapshot(region + "Top5Customers", sheet.Id, top5Customers.Id);

            // Get the Quarterly Trend on the sheet
            var quarterlyTrend = GetVisualisationWithTitle(sheet, "Quarterly Trend");

            // Take a snapshot of Quarterly Trend
            quarterlyTrendSnapshot = app.CreateSnapshot(region + "QuarterlyTrend", sheet.Id, quarterlyTrend.Id);
            return selectedIndex;
        }

        private static ISheet GetSheetWithTitle(IApp app, string title)
        {
            return app.GetSheets().FirstOrDefault(sheet => sheet.MetaAttributes.Title.ToLower() == title.ToLower());
        }

        private static IGenericObject GetVisualisationWithTitle(ISheet sheet, string title)
        {
            foreach (var cell in sheet.Cells)
            {
                var child = sheet.GetChild(cell.Name);
                if (child.GetLayout().As<VisualizationBaseLayout>().Title == title)
                    return child;
            }
            return null;
        }
    }
}
