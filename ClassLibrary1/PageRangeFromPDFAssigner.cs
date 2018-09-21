using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic.Citavi;
using SwissAcademic.Pdf.Analysis;

namespace QuotationsToolbox
{
    class PageRangeFromPositionInPDFAssigner
    {
        public static void AssignPageRangeFromPositionInPDF(List<KnowledgeItem> quotations)
        {
            Reference reference = quotations.FirstOrDefault().Reference;
            if (reference == null) return;

            if (quotations.Count == 0) return;

            int startPageInt = 1;

            if (reference.PageRange.StartPage.Number != null) startPageInt = reference.PageRange.StartPage.Number.Value;

            if (reference == null) return;

            foreach (KnowledgeItem quotation in quotations)
            {
                foreach (var entityLink in quotation.EntityLinks)
                {
                    if (entityLink.Indication.Equals(EntityLink.PdfKnowledgeItemIndication, StringComparison.OrdinalIgnoreCase) && entityLink.Target is Annotation)
                    {
                        Annotation annotation = quotation.EntityLinks.FirstOrDefault().Target as Annotation;
                        if (annotation == null) continue;
                        List<int> pages = new List<int>();
                        foreach (Quad quad in annotation.Quads)
                        {
                            pages.Add(startPageInt + quad.PageIndex - 1);
                        }
                        int annotationStartPageInt = pages.Min();
                        int annotationEndPageInt = pages.Max();
                        if (annotationStartPageInt == annotationEndPageInt)
                        {
                            quotation.PageRange = annotationStartPageInt.ToString();
                        }
                        else
                        {
                            quotation.PageRange = annotationStartPageInt.ToString() + "-" + annotationEndPageInt.ToString();
                        }
                    }
                }
            }
        }
    }
}
