
namespace TaskReorder
{
    using AppExtensions;
    using Microsoft.Office.Interop.OneNote;
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml.Linq;

    class Program
    {
        private static string UniqueApplicationId = "4C2DE1C4-DF3E-439D-8ECF-0B258A679A59";

        static void Main(string[] args)
        {
            CurrentApplication.EnsureSingleInstance(UniqueApplicationId);
            CurrentApplication.MakeProgramAutoRun();
            CurrentApplication.EnsureBackgroundWorker();

            // Have this app running always periodically waking up
            BlindlyRun(ReorderTasks, TimeSpan.FromSeconds(10));

            Thread.Sleep(Timeout.Infinite);
        }

        public static async Task BlindlyRun(Action action, TimeSpan period, CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                await Task.Delay(period, cancellationToken);
                
                try
                {
                    action();
                }
                catch
                {
                    // method is blind because it doesnt care about exceptions. sigh.
                }
            }
        }

        public static Task BlindlyRun(Action action, TimeSpan period)
        {
            return BlindlyRun(action, period, CancellationToken.None);
        }
                
        private static void ReorderTasks()
        {
            Application onenoteApp = new Application();
            
            XDocument page = onenoteApp.GetPageContents(
                "Ajay-Preethi Shared OneNotes", 
                "Our Notes", 
                "Shopping List");

            XNamespace ns = page.Root.Name.Namespace;

            XElement oeElementsParent = page
                .Element(ns + "Page")
                .Element(ns + "Outline")
                .Element(ns + "OEChildren");

            List<XElement> completedElements = new List<XElement>();
            List<XElement> newCuratedElementList = new List<XElement>();

            bool foundCompletedItem = false;
            bool foundNotCompletedItem = false;
            bool pageNeedsRefresh = false;

            foreach (XElement oeElement in oeElementsParent.Elements(ns + "OE"))
            {
                XElement tag = oeElement.Element(ns + "Tag");
                if (tag != null)
                {
                    if (tag.Attribute("completed").Value == "true")
                    {
                        foundCompletedItem = true;
                        completedElements.Add(oeElement);
                        continue;
                    }
                    else
                    {
                        foundNotCompletedItem = true;
                        if (foundCompletedItem == true)
                        {
                            // we are encountering an item that is not yet completed
                            // if we have already seen items that have been completed, 
                            // then the page needs to be refreshed.

                            pageNeedsRefresh = true;
                        }
                    }
                }

                newCuratedElementList.Add(oeElement);
            }

            if (!pageNeedsRefresh)
            {
                return;
            }

            foreach (XElement completedElement in completedElements)
            {
                newCuratedElementList.Add(completedElement);
            }

            oeElementsParent.RemoveAll();

            foreach (XElement curatedElement in newCuratedElementList)
            {
                oeElementsParent.Add(curatedElement);
            }

            onenoteApp.UpdatePageContent(page.ToString(), DateTime.MinValue);
        }
    }
}
