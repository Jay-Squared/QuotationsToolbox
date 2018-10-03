using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using pdftron.PDF;

namespace QuotationsToolbox
{
    class AnnotationExporter
    {
        public static void ExportAnnotations(List<Reference> references)
        {
            foreach (Reference reference in references)
            {
                List<KnowledgeItem> directQuotes = reference.Quotations.Where(q => q.QuotationType == QuotationType.DirectQuotation).ToList();

                foreach (KnowledgeItem directQuote in directQuotes)
                {
                    List<Annotation> annotations = directQuote.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Select(e => (Annotation)e.Target).ToList();

                    foreach (Annotation annotation in annotations)
                    {
                           
                    }
                }


            }
        }
    }
}
