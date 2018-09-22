using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

namespace QuotationsToolbox
{
    class SummaryCreator
    {
        public static void CreatesummaryOnQuotations(List<KnowledgeItem> quotations)
        {
            Reference reference = Program.ActiveProjectShell.PrimaryMainForm.ActiveReference;
            if (reference == null) return;

            Project project = reference.Project;
            if (project == null) return;

            Location location = reference.Locations.ToList().FirstOrDefault();

            Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
            SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;

            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return;

            PdfViewControl pdfViewControl = propertyInfo.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl) as PdfViewControl;

            List<PageRange> pageRanges = new List<PageRange>();
            List<Quad> quads = new List<Quad>();

            Annotation newAnnotation = new Annotation(location);

            KnowledgeItem summary = new KnowledgeItem(reference, QuotationType.Summary);

            foreach (KnowledgeItem quotation in quotations)
            {
                pageRanges.Add(quotation.PageRange);

                Annotation quotationAnnotation = quotation.EntityLinks.Where(link => link.Target is Annotation).FirstOrDefault().Target as Annotation;
                if (quotationAnnotation == null) continue;

                quads.AddRange(quotationAnnotation.Quads);
            }

            newAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
            newAnnotation.Quads = quads;
            newAnnotation.Visible = false;

            location.Annotations.Add(newAnnotation);

            summary.PageRange = PageRangeMerger.PageRangeListToString(pageRanges);
            summary.CoreStatementUpdateType = UpdateType.Automatic;
            reference.Quotations.Add(summary);

            EntityLink summaryAnnotationLink = new EntityLink(project);
            summaryAnnotationLink.Source = summary;
            summaryAnnotationLink.Target = newAnnotation;
            summaryAnnotationLink.Indication = EntityLink.PdfKnowledgeItemIndication;
            project.EntityLinks.Add(summaryAnnotationLink);

            quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate(summary, true);
            pdfViewControl.GoToAnnotation(newAnnotation);
        }
    }
}
