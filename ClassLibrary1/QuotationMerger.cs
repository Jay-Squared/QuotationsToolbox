using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using SwissAcademic;
using SwissAcademic.Citavi;
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

            Location location = reference.Locations.ToList().FirstOrDefault();

            List<Annotation> existingCitaviAnnotations = new List<Annotation>();
            List<Annotation> originalAnnotations = new List<Annotation>();

            List<Quad> quads = new List<Quad>();
            List<PageRange> pageRangesList = new List<PageRange>();
            string pageRangeText = string.Empty;

            string text = string.Empty;

            var pdfLocations = reference.GetPDFLocations();

            List<PageWidth> store = new List<PageWidth>();
            
            Document document = null;

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

            pageRangesList = MergeAdjacent(pageRangesList);

            int j = 0;
            foreach (PageRange pageRange in pageRangesList)
            {
                if (j > 0) pageRangeText += ", ";
                pageRangeText += pageRange.StartPage;
                if (pageRange.EndPage != null && !string.IsNullOrEmpty(pageRange.EndPage.OriginalString) && pageRange.EndPage != pageRange.StartPage) pageRangeText += "-" + pageRange.EndPage;
                j = j + 1;
            }

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
                rtbMerged.AppendText(" […] ");
                rtbMerged.Select(rtbMerged.Text.Length - 5, 5);
                rtbMerged.SelectionFont = new Font(font, System.Drawing.FontStyle.Bold);
                rtbMerged.Select(rtbMerged.Text.Length, 0);
            }

            rtbTemp.Rtf = s2;
            rtbTemp.SelectAll();
            rtbTemp.Cut();
            rtbMerged.Paste();

            return rtbMerged.Rtf;
        }
        private static List<PageRange> MergeAdjacent(List<PageRange> pageRangesList)
        {
            int i = pageRangesList.Count;

            if (pageRangesList.Count < 1) return pageRangesList;

            List<PageRange> newList = new List<PageRange>();

            PageNumber minPage = new PageNumber();
            PageNumber maxPage = new PageNumber();

            int minPageAsInteger = 0;
            int maxPageAsInteger = 0;

            PageRange first = pageRangesList.FirstOrDefault();
            PageRange last = pageRangesList.Last();

            NumberingType previousPageRangeNumberingType = new NumberingType();
            bool PreviousPageRangeWasNumber = false;

            foreach (PageRange pageRange in pageRangesList)
            {
                PageNumber currentStartPage = pageRange.StartPage;
                PageNumber currentEndPage = pageRange.EndPage;
                if (currentEndPage == null) currentEndPage = currentStartPage;
                if (string.IsNullOrEmpty(currentEndPage.OriginalString)) currentEndPage = currentStartPage;

                int currentStartPageAsInteger = 0;
                int currentEndPageAsInteger = 0;

                bool startPageIsNumber = Int32.TryParse(currentStartPage.OriginalString.Replace(".", ""), out currentStartPageAsInteger);
                bool endPageIsNumber = Int32.TryParse(currentEndPage.OriginalString.Replace(".", ""), out currentEndPageAsInteger);

                bool IsDiscreteRange = false;

                bool SameNumberingType = (previousPageRangeNumberingType == pageRange.NumberingType);

                if (pageRangesList.Count == 1)
                {

                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(pageRange.NumberingType);

                    Double result;
                    if (Double.TryParse(pageRange.StartPage.ToString(), out result) == false)
                    {
                        newList.Add(dummyPageRange.Update(pageRange.OriginalString));
                    }
                    else
                    {
                        newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", currentStartPage.OriginalString, currentEndPage.OriginalString)));
                    }
                    continue;
                }
                else if (pageRange == first)
                {
                    minPage = currentStartPage;
                    maxPage = currentEndPage;
                    minPageAsInteger = currentStartPageAsInteger;
                    maxPageAsInteger = currentEndPageAsInteger;
                }
                else if (!PreviousPageRangeWasNumber || !SameNumberingType || !startPageIsNumber || !endPageIsNumber || pageRange.StartPage.Number == null)
                {
                    IsDiscreteRange = true;
                }
                else if (currentStartPageAsInteger.CompareTo(minPageAsInteger) >= 0 && currentEndPageAsInteger.CompareTo(maxPageAsInteger) <= 0 && SameNumberingType)
                {
                    // In this case, we don't have to do anything because the current page range is within the range defined by minPage & maxPage
                }
                else if (currentStartPageAsInteger.CompareTo(maxPageAsInteger + 1) < 1 && SameNumberingType)
                {
                    maxPage = currentEndPage;
                    maxPageAsInteger = currentEndPageAsInteger;
                }
                else
                {
                    IsDiscreteRange = true;
                }

                if (IsDiscreteRange && pageRange == last)
                {
                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(previousPageRangeNumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                    minPage = currentStartPage;
                    maxPage = currentEndPage;
                    dummyPageRange = dummyPageRange.Update(pageRange.NumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                }
                else if (IsDiscreteRange)
                {
                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(previousPageRangeNumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                    minPage = currentStartPage;
                    maxPage = currentEndPage;
                    minPageAsInteger = currentStartPageAsInteger;
                    maxPageAsInteger = currentEndPageAsInteger;
                }
                else if ((newList.Count == 0 || !PreviousPageRangeWasNumber || !SameNumberingType) && pageRange == last)
                {
                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(pageRange.NumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                }
                previousPageRangeNumberingType = pageRange.NumberingType;
                PreviousPageRangeWasNumber = true;
            }
            return newList;
        } // end MergeAdjacent
    }
    class PageRangeComparer : IComparer<PageRange>
    {
        public int Compare(PageRange x, PageRange y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    return 0;
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                if (y == null)
                {
                    return 1;
                }
                else
                {
                    if (x.NumberingType.CompareTo(y.NumberingType) != 0)
                    {
                        return x.NumberingType.CompareTo(y.NumberingType);
                    }
                    else if (x.NumeralSystem.CompareTo(y.NumeralSystem) != 0)
                    {
                        return x.NumeralSystem.CompareTo(y.NumeralSystem);
                    }
                    else
                    {
                        return x.StartPage.CompareTo(y.StartPage);
                    }
                }
            }
        }
    } // end PageRangeComparer
}
