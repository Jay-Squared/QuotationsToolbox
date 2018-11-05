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

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            Document document = pdfViewControl.Document;
            if (document == null) return;

            Location location = reference.Locations.Where
                (
                    l => 
                    l.LocationType == LocationType.ElectronicAddress
                    && l.Address.Resolve().LocalPath.EndsWith(".pdf")
                    && l.Address.Resolve().LocalPath == document.GetFileName()
                )
            .FirstOrDefault();
            if (location == null) return;

            Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
            SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

            List<PageRange> pageRanges = new List<PageRange>();
            List<Quad> quads = new List<Quad>();

            Annotation newAnnotation = new Annotation(location);

            KnowledgeItem summary = new KnowledgeItem(reference, QuotationType.Summary);
            reference.Quotations.Add(summary);

            foreach (KnowledgeItem quotation in quotations)
            {
                pageRanges.Add(quotation.PageRange);

                EntityLink summaryQuotationLink = new EntityLink(project);
                summaryQuotationLink.Source = summary;
                summaryQuotationLink.Target = quotation;
                summaryQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                project.EntityLinks.Add(summaryQuotationLink);

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
