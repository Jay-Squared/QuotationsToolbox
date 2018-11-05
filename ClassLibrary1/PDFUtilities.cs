using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;

using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Controls.WindowsForms;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using pdftron.PDF;

namespace QuotationsToolbox
{
    public static class PDFUtilities
    {
        public static int GetStartPage(this IEnumerable<Quad> quads)
        {
            var pageNumbers = quads
                              .Select(q => q.PageIndex)
                              .Distinct()
                              .OrderBy(p => p);
            return pageNumbers.First();
        }

        public static List<Location> GetPDFLocations(this List<KnowledgeItem> knowledgeItems)
        {
            var locations = new List<Location>();

            foreach (var knowledgeItem in knowledgeItems)
            {
                foreach (var entityLink in knowledgeItem.EntityLinks)
                {
                    if (entityLink.Indication.Equals(EntityLink.PdfKnowledgeItemIndication, StringComparison.OrdinalIgnoreCase) && entityLink.Target is Annotation)
                    {
                        var location = ((Annotation)entityLink.Target).Location;

                        if (location == null) continue;
                        if (location.LocationType != LocationType.ElectronicAddress) continue;
                        if (location.Address.Resolve().LocalPath.EndsWith(".pdf") == false) continue;
                        if (locations.Contains(location)) continue;

                        locations.Add(location);
                    }
                }
            }
            return locations;
        }

        public static List<Location> GetPDFLocations(this Reference reference)
        {
            List<Location> locations = reference.Locations.Where(l => l.LocationType == LocationType.ElectronicAddress && l.Address.Resolve().LocalPath.EndsWith(".pdf")).ToList();
            return locations;
        }
    }
}
