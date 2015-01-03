
namespace TaskReorder
{
    using Microsoft.Office.Interop.OneNote;
    using System.Linq;
    using System.Xml.Linq;

    public static class ApplicationExtensions
    {
        public static XDocument GetPageContents(this Application onenoteApp, string notebookName, string sectionName, string pageName)
        {
            string notebookXml;

            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            var pageNode = doc.Descendants(ns + "Page").Where(
                n => n.Attribute("name").Value == pageName
                && n.Parent.Attribute("name").Value == sectionName
                && n.Parent.Parent.Attribute("name").Value == notebookName).FirstOrDefault();

            if (pageNode != null)
            {
                string pageXml;
                onenoteApp.GetPageContent(pageNode.Attribute("ID").Value, out pageXml);
                return XDocument.Parse(pageXml);
            }

            return null;
        }
    }
}
