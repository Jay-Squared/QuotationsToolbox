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
    public class KnowledgeItemInCategorySorter
    {
        public static void SortKnowledgeItemsInCategorySorter(MainForm mainForm)
        {
            var category = mainForm.GetSelectedKnowledgeOrganizerCategory();

            var knowledgeItems = category.KnowledgeItems.ToList();

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
                    category.KnowledgeItems.Move(knowledgeItems[i], firstKnowledgeItem);
                    firstKnowledgeItem = knowledgeItems[i];
                }

                CreateSubheadings(knowledgeItems, category, true);
            }
        }
        static void CreateSubheadings(List<KnowledgeItem> knowledgeItems, Category category, bool overwriteSubheadings)
        {
            var mainForm = Program.ActiveProjectShell.PrimaryMainForm;
            var projectShell = Program.ActiveProjectShell;
            var project = projectShell.Project;

            var categoryKnowledgeItems = category.KnowledgeItems;
            var subheadings = knowledgeItems.Where(item => item.KnowledgeItemType == KnowledgeItemType.Subheading).ToList();

            Reference currentReference = null;
            Reference previousReference = null;

            int nextInsertionIndex = -1;

            if (subheadings.Any())
            {
                if (!overwriteSubheadings)
                {
                    DialogResult result = MessageBox.Show("The filtered list of knowledge items in the category \"" + category.Name + "\" already contains sub-headings.\r\n\r\nIf you continue, these sub-headings will be removed first.\r\n\r\nContinue?", "Citavi", MessageBoxButtons.YesNo);
                    if (result == DialogResult.No) return;
                }

                foreach (var subheading in subheadings)
                {
                    project.Thoughts.Remove(subheading);
                }

                projectShell.SaveAsync(mainForm);
            }

            foreach (var knowledgeItem in knowledgeItems)
            {
                if (knowledgeItem.KnowledgeItemType == KnowledgeItemType.Subheading)
                {
                    knowledgeItem.Categories.Remove(category);
                    continue;
                }

                if (knowledgeItem.Reference != null) currentReference = knowledgeItem.Reference;

                string headingText = "No short title available";
                if (currentReference != null)
                {
                    headingText = currentReference.ShortTitle;
                }
                else if (knowledgeItem.QuotationType == QuotationType.None)
                {
                    headingText = "Thoughts";
                }

                nextInsertionIndex = category.KnowledgeItems.IndexOf(knowledgeItem);
                category.KnowledgeItems.AddNextItemAtIndex = nextInsertionIndex;
                currentReference = knowledgeItem.Reference;

                if (nextInsertionIndex == 0)
                {
                    var subheading = new KnowledgeItem(project, KnowledgeItemType.Subheading) { CoreStatement = headingText };
                    subheading.Categories.Add(category);

                    project.Thoughts.Add(subheading);
                    projectShell.SaveAsync(mainForm);
                    previousReference = currentReference;
                    continue;
                }

                if (nextInsertionIndex > 0 && (currentReference != null && currentReference != previousReference))
                {
                    var subheading = new KnowledgeItem(project, KnowledgeItemType.Subheading) { CoreStatement = headingText };
                    subheading.Categories.Add(category);

                    project.Thoughts.Add(subheading);
                    projectShell.SaveAsync(mainForm);
                }

                previousReference = currentReference;
            }
        }
    }
}
