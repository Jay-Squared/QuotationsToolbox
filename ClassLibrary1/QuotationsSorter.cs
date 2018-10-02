using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic.Citavi;
using SwissAcademic.Pdf;

namespace QuotationsToolbox
{
    class QuotationsSorter
    {
        public static void SortQuotations(List<KnowledgeItem> quotations)
        {
            if (quotations.Count <= 1) return;

            Reference reference = quotations.FirstOrDefault().Reference;
            if (reference == null) return;

            var locations = quotations.GetPDFLocations();

            if (locations == null) return;

            List<PageWidth> store = new List<PageWidth>();

            foreach (Location location in locations)
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

            quotations.Sort(new KnowledgeItemComparer(store));

            var firstQuotation = quotations.First();

            for (int i = 1; i < quotations.Count; i++)
            {
                reference.Quotations.Move(quotations[i], firstQuotation);
                firstQuotation = quotations[i];
            }
        }
    }
}
