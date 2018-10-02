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
        public static Document GetDocument(this PreviewControl previewControl)
        {
            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return null;

            PdfViewControl control = propertyInfo.GetValue(previewControl) as PdfViewControl;

            return control.Document;
        }

        public static int GetStartPage(this IEnumerable<Quad> quads)
        {
            var pageNumbers = quads
                              .Select(q => q.PageIndex)
                              .Distinct()
                              .OrderBy(p => p);
            return pageNumbers.First();
        }

        public static List<Location> GetPDFLocations(this Reference reference)
        {
            List<Location> locations = reference.Locations.Where(l => l.LocationType == LocationType.ElectronicAddress && l.Address.Resolve().LocalPath.EndsWith(".pdf")).ToList();
            return locations;
        }

        public static Location GetPDFLocationOfDocument(this Document document)
        {
            List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.Project.References.ToList();

            string foundPDFs = document.GetFileName();

            List<Reference> referencesWithLocation = new List<Reference>();

            foreach (Reference reference in references)
            {
                Location[] locations = reference.Locations.ToArray();

                foreach (Location location in locations)
                {
                    if (location.Address.LinkedResourceType == LinkedResourceType.AttachmentFile)
                    {
                        string referenceFile = location.Address.Resolve().LocalPath;
                        if (foundPDFs.Contains(referenceFile)) referencesWithLocation.Add(reference);
                    }
                }
            }

            Reference referenceWithLocation = referencesWithLocation.FirstOrDefault();

            Location result = referenceWithLocation.Locations.Where(l => l.LocationType == LocationType.ElectronicAddress
                && l.Address.Resolve().LocalPath.EndsWith(".pdf") && l.Address.Resolve().LocalPath == document.GetFileName()).FirstOrDefault();
            if (result == null) return null;

            return result;
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
    }
}
