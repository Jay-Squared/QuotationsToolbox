using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic.Citavi;
using SwissAcademic.Pdf;

namespace QuotationsToolbox
{
    class PageRangeFromPrecedingQuotationAssigner
    {
        public static void AssignPageRangeFromPrecedingQuotation(List<KnowledgeItem> quotations)
        {
            Reference reference = quotations.FirstOrDefault().Reference;
            if (reference == null) return;

            List<KnowledgeItem> referenceQuotations = reference.Quotations.ToList();
            if (referenceQuotations == null) return;

            var pdfLocations = reference.GetPDFLocations();

            List<PageWidth> store = new List<PageWidth>();

            if (pdfLocations != null)
            {
                foreach (Location location in pdfLocations)
                {
                    Document document = null;

                    var address = location.Address.Resolve().LocalPath;
                    document = new Document(address);

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
                }
            }


            referenceQuotations.Sort(new KnowledgeItemComparer(store));

            foreach (KnowledgeItem quotation in quotations)
            {
                int index = referenceQuotations.FindIndex(q => q == quotation);
                if (index <= 1) continue;

                KnowledgeItem previousQuotation = referenceQuotations[index - 1];
                if (previousQuotation == null) continue;

                quotation.PageRange = previousQuotation.PageRange;
                quotation.PageRange = quotation.PageRange.Update(previousQuotation.PageRange.NumberingType);
                quotation.PageRange = quotation.PageRange.Update(previousQuotation.PageRange.NumeralSystem);
            }

        }
    }
}
