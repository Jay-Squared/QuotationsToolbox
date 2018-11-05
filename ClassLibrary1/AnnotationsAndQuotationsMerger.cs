using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Controls.WindowsForms;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;


namespace QuotationsToolbox
{
    class AnnotationsAndQuotationsMerger
    {
        public static void MergeAnnotations()
        {
            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            Document document = pdfViewControl.Document;
            if (document == null) return;

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return;

            Location location = previewControl.ActiveLocation;
            if (location == null) return;

            Reference reference = location.Reference;
            if (reference == null) return;

            List<Annotation> annotations = pdfViewControl.GetSelectedAnnotations().ToList();

            List<KnowledgeItem> quotations = annotations
                .Where
                (
                    a =>
                    a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0
                )
                .Select
                (
                    a =>
                    (KnowledgeItem)a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Source
                )
                .ToList();

            List<QuotationType> quotationTypes = quotations.Select(q => q.QuotationType).Distinct().ToList();
            var itemToRemove = quotationTypes.SingleOrDefault(q => q == QuotationType.Highlight);
            quotationTypes.Remove(itemToRemove);

            if (quotations.Count() > 0)
            {
                if (quotationTypes.Count > 1)
                {
                    MessageBox.Show("Can't merge quotations, more than one type of quotation is selected.");
                    return;
                }
            }

            // Dynamic Variables

            // The Magic

            KnowledgeItem newQuotation = CombineQuotations(quotations);            
            Annotation newAnnotation = CombineAnnotations(annotations);
            
            if (quotations.Count() > 1)
            {
                EntityLink newEntityLink = new EntityLink(project);
                newEntityLink.Source = newQuotation;
                newEntityLink.Target = newAnnotation;
                newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                project.EntityLinks.Add(newEntityLink);
                newAnnotation.Visible = false;
                reference.Quotations.RemoveRange(quotations);
            }
            else if (quotations.Count == 1)
            {
                EntityLink newEntityLink = new EntityLink(project);
                newEntityLink.Source = quotations.FirstOrDefault();
                newEntityLink.Target = newAnnotation;
                newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                project.EntityLinks.Add(newEntityLink);
                newAnnotation.Visible = false;
            }

            if (quotations.Count > 1 && quotationTypes.Count > 0) Program.ActiveProjectShell.ShowKnowledgeItemFormForExistingItem(Program.ActiveProjectShell.PrimaryMainForm, newQuotation);
        }

        public static void MergeQuotations(List<KnowledgeItem> quotations)
        {
            // Static Variables

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return;

            Reference reference = quotations.FirstOrDefault().Reference;
            if (reference == null) return;

            List<Annotation> annotations = quotations.Where
            (
                q =>
                q.EntityLinks.Any() &&
                q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication && e.Target is Annotation).Count() > 0
            )
            .Select
            (
                q => ((Annotation)
                q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication && e.Target is Annotation).FirstOrDefault().Target)
            )
            .ToList();

            // The Magic

            List<QuotationType> quotationTypes = quotations.Select(q => q.QuotationType).Distinct().ToList();
            var itemToRemove = quotationTypes.SingleOrDefault(q => q == QuotationType.Highlight);
            quotationTypes.Remove(itemToRemove);

            if (quotationTypes.Count > 1)
            {
                MessageBox.Show("Can't merge quotations, more than one type of quotation is selected.");
                return;
            }

            KnowledgeItem newQuotation = CombineQuotations(quotations);

            Annotation newAnnotation = CombineAnnotations(annotations);

            EntityLink newEntityLink = new EntityLink(project);
            newEntityLink.Source = newQuotation;
            newEntityLink.Target = newAnnotation;
            newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
            project.EntityLinks.Add(newEntityLink);

            reference.Quotations.RemoveRange(quotations);

            Program.ActiveProjectShell.ShowKnowledgeItemFormForExistingItem(Program.ActiveProjectShell.PrimaryMainForm, newQuotation);
        }

        public static Annotation CombineAnnotations(List<Annotation> annotations)
        {
            if (annotations.Count() == 1) return annotations.FirstOrDefault();

            // Static Variables

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            List<Location> locations = annotations.Select(a => a.Location).Distinct().ToList();
            if (locations.Count != 1) return null;
            Location location = locations.FirstOrDefault();

            List<Annotation> originalAnnotations = annotations
                .Where
                (
                    a =>
                    a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() == 0
                )
                .ToList();

            List<Annotation> redundantAnnotations = annotations
                .Where
                (
                    a =>
                    a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0
                )
                .ToList();

            List<Annotation> soureAnnotations = annotations
            .Where
            (
                a =>
                a.EntityLinks.Where
                (
                    e =>
                    e.Indication == EntityLink.SourceAnnotLinkIndication &&
                    (Annotation)e.Source != null &&
                    (Annotation)e.Source != a
                )
                .Count() > 0
            )
            .Select
            (
                a =>
                (Annotation)
                a.EntityLinks.Where
                (
                    e =>
                    e.Indication == EntityLink.SourceAnnotLinkIndication &&
                    (Annotation)e.Source != null &&
                    (Annotation)e.Source != a
                )
                .FirstOrDefault().Source
            ).ToList();

            // Dynamic Variables

            Annotation newAnnotation = new Annotation(location);

            List<Quad> quads = new List<Quad>();

            // The Magic

            foreach (Annotation annotation in annotations)
            {
                quads.AddRange(annotation.Quads);
            }

            newAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
            newAnnotation.Quads = quads;
            newAnnotation.Visible = true;
            location.Annotations.Add(newAnnotation);

            foreach (Annotation originalAnnotation in originalAnnotations)
            {
                originalAnnotation.Visible = false;
                AnnotationsImporter.LinkWithKnowledgeItemIndicationAnnotation(originalAnnotation, newAnnotation);
            }

            foreach (Annotation redundantAnnotation in redundantAnnotations)
            {
                location.Annotations.Remove(redundantAnnotation);
            }

            foreach (Annotation soureAnnotation in soureAnnotations)
            {
                soureAnnotation.Visible = false;
                AnnotationsImporter.LinkWithKnowledgeItemIndicationAnnotation(soureAnnotation, newAnnotation);
            }



            return newAnnotation;
        }

        public static KnowledgeItem CombineQuotations(List<KnowledgeItem> quotations)
        {
            if (quotations.Count() == 1) return quotations.FirstOrDefault();
            if (quotations.Count == 0) return null;

            // Static Variables

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return null;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return null;

            Document document = pdfViewControl.Document;
            if (document == null) return null;

            Reference reference = quotations.FirstOrDefault().Reference;
            if (reference == null) return null;

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            List<Location> locations = quotations.GetPDFLocations().Distinct().ToList();
            if (locations.Count != 1) return null;
            Location location = locations.FirstOrDefault();

            List<QuotationType> quotationTypes = quotations.Select(q => q.QuotationType).Distinct().ToList();
            var itemToRemove = quotationTypes.SingleOrDefault(q => q == QuotationType.Highlight);
            quotationTypes.Remove(itemToRemove);

            QuotationType quotationType = quotationTypes.FirstOrDefault();

            if (quotationTypes.Count == 0)
            {
                quotationType = QuotationType.Highlight;
            }

            // Dynamic Variables

            KnowledgeItem newQuotation = new KnowledgeItem(reference, quotationType);

            string text = string.Empty;

            List<Quad> quads = new List<Quad>();
            List<PageRange> pageRangesList = new List<PageRange>();
            string pageRangeText = string.Empty;

            List<PageWidth> store = new List<PageWidth>();

            // The Magic

            if (document != null)
            {
                for (int i = 1; i <= document.GetPageCount(); i++)
                {
                    pdftron.PDF.Page page = document.GetPage(i);
                    if (page.IsValid())
                    {
                        var re = page.GetCropBox();
                        store.Add(new PageWidth(location, i, re.Width()));
                    }
                    else
                    {
                        store.Add(new PageWidth(location, i, 0.00));
                    }
                }
            }

            quotations.Sort(new KnowledgeItemComparer(store));

            foreach (KnowledgeItem quotation in quotations)
            {
                if (!string.IsNullOrEmpty(quotation.Text)) text = MergeRTF(text, quotation.TextRtf);
                if (!string.IsNullOrEmpty(quotation.PageRange.OriginalString)) pageRangesList.Add(quotation.PageRange);
            }

            pageRangesList = PageRangeMerger.MergeAdjacent(pageRangesList);
            pageRangeText = PageRangeMerger.PageRangeListToString(pageRangesList);

            newQuotation.TextRtf = text;
            newQuotation.PageRange = pageRangeText;
            newQuotation.PageRange = newQuotation.PageRange.Update(quotations[0].PageRange.NumberingType);
            newQuotation.PageRange = newQuotation.PageRange.Update(quotations[0].PageRange.NumeralSystem);

            reference.Quotations.Add(newQuotation);
            project.AllKnowledgeItems.Add(newQuotation);

            return newQuotation;
        }

        private static string MergeRTF(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s2)) return s1;

            RichTextBox rtbTemp = new RichTextBox();
            RichTextBox rtbMerged = new RichTextBox();

            if (!string.IsNullOrEmpty(s1))
            {
                rtbTemp.Rtf = s1;
                rtbTemp.SelectAll();
                rtbTemp.Cut();
                rtbMerged.Paste();
                rtbMerged.Select(1, 1);
                Font font = rtbMerged.SelectionFont;
                rtbMerged.AppendText(" ");
                rtbMerged.Select(rtbMerged.Text.Length, 0);
            }

            rtbTemp.Rtf = s2;
            rtbTemp.SelectAll();
            rtbTemp.Cut();
            rtbMerged.Paste();

            return rtbMerged.Rtf;
        }
    }
}
