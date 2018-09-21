using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Pdf;

namespace QuotationsToolbox
{
    public class KnowledgeItemInSelectionSorter
    {
        public static void SortSelectedKnowledgeItems(MainForm mainForm)
        {
            List<KnowledgeItem> knowledgeItems = mainForm.GetSelectedKnowledgeItems();
            var category = mainForm.KnowledgeOrganizerFilterSet.Filters[0].Category;

            var categoryKnowledgeItems = category.KnowledgeItems;

            if (knowledgeItems.Count > 1)
            {
                var pdfLocations = knowledgeItems.GetPDFLocations();

                List<PageWidth> store = new List<PageWidth>();

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
                } // end foreach

                knowledgeItems.Sort(new KnowledgeItemComparer(store));

                var firstKnowledgeItem = knowledgeItems.First();
                for (int i = 1; i < knowledgeItems.Count; i++)
                {
                    categoryKnowledgeItems.Move(knowledgeItems[i], firstKnowledgeItem);
                    firstKnowledgeItem = knowledgeItems[i];
                }
            }
        }
      
    }
}
