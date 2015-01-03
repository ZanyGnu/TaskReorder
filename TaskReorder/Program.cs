
namespace TaskReorder
{
    using Microsoft.Office.Interop.OneNote;
    using System;
    using System.Collections.Generic;
    using System.Xml.Linq;

    class Program
    {
        static void Main(string[] args)
        {
            Application onenoteApp = new Application();
            
            XDocument page = onenoteApp.GetPageContents("TestOneNote", "TestSection", "Shopping List");

            XNamespace ns = page.Root.Name.Namespace;

            XElement oeElementsParent = page
                .Element(ns + "Page")
                .Element(ns + "Outline")
                .Element(ns + "OEChildren");

            List<XElement> completedElements = new List<XElement>();
            List<XElement> newCuratedElementList = new List<XElement>();

            foreach (XElement oeElement in oeElementsParent.Elements(ns + "OE"))
            {
                XElement tag = oeElement.Element(ns + "Tag");
                if (tag != null)
                {
                    if (tag.Attribute("completed").Value == "true")
                    {
                        completedElements.Add(oeElement);
                        continue;
                    }
                }

                newCuratedElementList.Add(oeElement);
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
