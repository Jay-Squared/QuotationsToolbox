using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
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
    class QuotationMerger
    {
        public static void MergeQuotations(List<KnowledgeItem> quotations)
        {
            if (quotations.Count <= 1) return;

            Reference reference = quotations.FirstOrDefault().Reference;
            if (reference == null) return;

            Project project = reference.Project;
            if (project == null) return;

            List<Location> locations = quotations.GetPDFLocations();

            if (locations.Count != 1) return;

            Location location = locations.FirstOrDefault();

            List<Annotation> existingCitaviAnnotations = new List<Annotation>();
            List<Annotation> originalAnnotations = new List<Annotation>();

            List<Quad> quads = new List<Quad>();
            List<PageRange> pageRangesList = new List<PageRange>();
            string pageRangeText = string.Empty;

            string text = string.Empty;
            List<PageWidth> store = new List<PageWidth>();

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            Document document = previewControl.GetDocument();

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

            var address = location.Address.Resolve().LocalPath;
            document = new Document(address);

            quotations.Sort(new KnowledgeItemComparer(store));

            foreach (KnowledgeItem quotation in quotations)
            {
                if (quotation.EntityLinks.Any() && quotation.EntityLinks.Where(link => link.Target is Annotation).FirstOrDefault() != null)
                {
                    Annotation annotation = (Annotation)quotation.EntityLinks.Where(e => e.Target is Annotation && e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Target;
                    {
                        if (annotation.Location != null && ((location != null && annotation.Location == location) || location == null))
                        {
                            location = annotation.Location;
                            existingCitaviAnnotations.Add(annotation);

                            foreach (EntityLink entityLink in annotation.EntityLinks.Where(e => e.Indication == EntityLink.SourceAnnotLinkIndication))
                            {
                                if ((Annotation)entityLink.Source != null) originalAnnotations.Add((Annotation)entityLink.Source);
                            }

                            if (!string.IsNullOrEmpty(quotation.PageRange.OriginalString)) pageRangesList.Add(quotation.PageRange);
                            quads.AddRange(annotation.Quads);
                        }
                    }
                }
                text = MergeRTF(text, quotation.TextRtf);
            }


            pageRangeText = PageRangeMerger.PageRangeListToString(pageRangesList);

            pageRangesList = PageRangeMerger.MergeAdjacent(pageRangesList);

            Annotation newAnnotation = new Annotation(location);
            newAnnotation.Quads = quads;
            newAnnotation.Visible = false;

            location.Annotations.Add(newAnnotation);

            KnowledgeItem newQuotation = new KnowledgeItem(reference, QuotationType.DirectQuotation);
            newQuotation.TextRtf = text;
            newQuotation.PageRange = pageRangeText;
            newQuotation.PageRange = newQuotation.PageRange.Update(quotations[0].PageRange.NumberingType);
            newQuotation.PageRange = newQuotation.PageRange.Update(quotations[0].PageRange.NumeralSystem);
            reference.Quotations.Add(newQuotation);

            EntityLink newEntityLink = new EntityLink(project);
            newEntityLink.Source = newQuotation;
            newEntityLink.Target = newAnnotation;
            newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
            project.EntityLinks.Add(newEntityLink);

            foreach (KnowledgeItem quotation in quotations)
            {
                reference.Quotations.Remove(quotation);
            }

            foreach (Annotation annotation in existingCitaviAnnotations)
            {
                location.Annotations.Remove(annotation);
            }

            foreach (Annotation originalAnnotation in originalAnnotations)
            {
                originalAnnotation.Visible = false;
                EntityLink sourceAnnotLink = new EntityLink(project);
                sourceAnnotLink.Indication = EntityLink.SourceAnnotLinkIndication;
                sourceAnnotLink.Source = originalAnnotation;
                sourceAnnotLink.Target = newAnnotation;
                project.EntityLinks.Add(sourceAnnotLink);
            }
        }
        private static string MergeRTF(string s1, string s2)
        {
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
