using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Controls;
using SwissAcademic.Citavi.Shell;

using Infragistics.Win.UltraWinToolbars;

namespace QuotationsToolbox
{
    class MenuQuotations
                :
        CitaviAddOn
    {
        public override AddOnHostingForm HostingForm
        {
            get { return AddOnHostingForm.MainForm; }
        }

        protected override void OnHostingFormLoaded(System.Windows.Forms.Form hostingForm)
        {
            MainForm mainForm = (MainForm)hostingForm;

            // Reference Editor Reference Menu

            var referencesMenu = mainForm.GetMainCommandbarManager().GetReferenceEditorCommandbar(MainFormReferenceEditorCommandbarId.Menu).GetCommandbarMenu(MainFormReferenceEditorCommandbarMenuId.References);

            var commandbarButtonConvertAnnotations = referencesMenu.AddCommandbarButton("ConvertAnnotations", "Convert external annotations to quotations in active reference");
            commandbarButtonConvertAnnotations.HasSeparator = true;

            // Quotations Pop-Up Menu

            var referenceEditorQuotationsContextMenu = CommandbarMenu.Create(mainForm.GetReferenceEditorQuotationsCommandbarManager().ToolbarsManager.Tools["ReferenceEditorQuotationsContextMenu"] as PopupMenuTool);

            var commandbarButtonKnowledgeItemsSortInReference = referenceEditorQuotationsContextMenu.AddCommandbarButton("KnowledgeItemsSortInReference", "Sort selected quotations by position in PDF");
            commandbarButtonKnowledgeItemsSortInReference.Shortcut = Shortcut.CtrlK;
            commandbarButtonKnowledgeItemsSortInReference.HasSeparator = true;
            var commandbarButtonPageAssignFromPositionInPDF = referenceEditorQuotationsContextMenu.AddCommandbarButton("PageAssignFromPositionInPDF", "Assign a page range to selected quote based on the PDF position");
            commandbarButtonPageAssignFromPositionInPDF.HasSeparator = true;
            referenceEditorQuotationsContextMenu.AddCommandbarButton("PageAssignFromPreviousQuotation", "Assign page range and numbering type from previous quote to selected quote");
            var commandbarButtonShowQuotationAndSetPageRangeManually = referenceEditorQuotationsContextMenu.AddCommandbarButton("ShowQuotationAndSetPageRangeManually", "Assign page range manually after showing selected quote in PDF");
            commandbarButtonShowQuotationAndSetPageRangeManually.Shortcut = Shortcut.CtrlDel;

            var commandbarButtonCommentAnnotation = referenceEditorQuotationsContextMenu.AddCommandbarButton("CommentAnnotation", "Link comment to same text in PDF as related quote");
            commandbarButtonCommentAnnotation.HasSeparator = true;
            var commandbarButtonCreateCommentOnQuotation = referenceEditorQuotationsContextMenu.AddCommandbarButton("CreateCommentOnQuotation", "Create comment on selected quote");
            commandbarButtonCreateCommentOnQuotation.Shortcut = Shortcut.CtrlShiftK;
            referenceEditorQuotationsContextMenu.AddCommandbarButton("SelectLinkedKnowledgeItem", "Jump to related quote or comment");

            var commandbarButtonCleanQuotationsText = referenceEditorQuotationsContextMenu.AddCommandbarButton("CleanQuotationsText", "Clean text of selected quotes");
            commandbarButtonCleanQuotationsText.Shortcut = Shortcut.ShiftDel;
            commandbarButtonCleanQuotationsText.HasSeparator = true;
            referenceEditorQuotationsContextMenu.AddCommandbarButton("QuickReferenceTitleCase", "Write core statement in title case in selected quick references");
            referenceEditorQuotationsContextMenu.AddCommandbarButton("QuotationsMerge", "Merge selected quotes");
            referenceEditorQuotationsContextMenu.AddCommandbarButton("ConvertDirectQuoteToRedHighlight", "Convert selected quotes to quick references");

            // Knowledge Item Pop-Up Menu

            var knowledgeOrganizerKnowledgeItemsContextMenu = CommandbarMenu.Create(mainForm.GetKnowledgeOrganizerKnowledgeItemsCommandbarManager().ToolbarsManager.Tools["KnowledgeOrganizerKnowledgeItemsContextMenu"] as PopupMenuTool);

            var commandBarButtonOpenKnowledgeItemAttachment = knowledgeOrganizerKnowledgeItemsContextMenu.AddCommandbarButton("OpenKnowledgeItemAttachment", "Open the attachment in new window");
            commandBarButtonOpenKnowledgeItemAttachment.HasSeparator = true;
            knowledgeOrganizerKnowledgeItemsContextMenu.AddCommandbarButton("SelectLinkedKnowledgeItem", "Jump to the linked knowledge item");
            var commandbarButtonKnowledgeItemsSortInCategory = knowledgeOrganizerKnowledgeItemsContextMenu.AddCommandbarButton("SortKnowledgeItemsInSelection", "Sort selected knowledge items in current category by reference and position");

            // Knowledge Organizer Category Pop-Up Menu

            var knowledgeOrganizerCategoriesColumnContextMenu = CommandbarMenu.Create(mainForm.GetKnowledgeOrganizerCategoriesCommandbarManager().ToolbarsManager.Tools["KnowledgeOrganizerCategoriesColumnContextMenu"] as PopupMenuTool);


            var commandbarButtonSortKnowledgeItemsInCategory = knowledgeOrganizerCategoriesColumnContextMenu.AddCommandbarButton("SortKnowledgeItemsInCategory", "Sort all knowledge items in this category by reference and position");
            commandbarButtonSortKnowledgeItemsInCategory.HasSeparator = true;

            //System.Diagnostics.Debug.WriteLine("Toolbars");

            //System.Diagnostics.Debug.WriteLine(mainForm.GetKnowledgeOrganizerKnowledgeItemsCommandbarManager().ToolbarsManager.Tools.ToString());

            //foreach (ToolBase tool in (mainForm.GetKnowledgeOrganizerKnowledgeItemsCommandbarManager().ToolbarsManager.Tools))
            //{
            //    System.Diagnostics.Debug.WriteLine(tool.Key.ToString());
            //}
            //foreach (ToolBase tool in (mainForm.GetKnowledgeOrganizerCategoriesCommandbarManager().ToolbarsManager.Tools))
            //{
            //    System.Diagnostics.Debug.WriteLine(tool.Key.ToString());
            //}

            // Fin

            base.OnHostingFormLoaded(hostingForm);
        }

        protected override void OnBeforePerformingCommand(SwissAcademic.Controls.BeforePerformingCommandEventArgs e)
        {
            Project project = Program.ActiveProjectShell.Project;

            switch (e.Key)
            {
                #region Reference-based commands
                case "ConvertAnnotations":
                    {
                        e.Handled = true;
                        Reference reference = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().FirstOrDefault();
                        AnnotationConverter.ConvertAnnotations(reference);
                    }
                    break;
                #endregion
                #region Quotations Pop-Up Menu
                case "CleanQuotationsText":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        QuotationTextCleaner.CleanQuotationsText(quotations);
                    }
                    break;
                case "ConvertDirectQuoteToRedHighlight":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        QuotationTypeConverter.ConvertDirectQuoteToRedHighlight(quotations);
                    }
                    break;
                case "CreateCommentOnQuotation":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        CommentCreator.CreateCommentOnQuotation(quotations);
                    }
                    break;
                case "KnowledgeItemsSortInReference":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        QuotationsSorter.SortQuotations(quotations);
                    }
                    break;
                case "CommentAnnotation":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        CommentAnnotationCreator.CreateCommentAnnotation(quotations);
                    }
                    break;
                case "PageAssignFromPositionInPDF":
                    {

                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        PageRangeFromPositionInPDFAssigner.AssignPageRangeFromPositionInPDF(quotations);
                    }
                    break;
                case "PageAssignFromPreviousQuotation":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        PageRangeFromPrecedingQuotationAssigner.AssignPageRangeFromPrecedingQuotation(quotations);
                    }
                    break;
                case "QuotationsMerge":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        QuotationMerger.MergeQuotations(quotations);
                    }
                    break;
                case "QuickReferenceTitleCase":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        QuickReferenceTitleCaser.TitleCaseQuickReference(quotations);
                    }
                    break;
                case "ShowQuotationAndSetPageRangeManually":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        PageRangeManualAssigner.AssignPageRangeManuallyAfterShowingAnnotation();
                    }
                    break;
                #endregion
                #region KnowledgeOrganizerKnowledgeItemsContextMenu
                case "OpenKnowledgeItemAttachment":
                    {
                        e.Handled = true;

                        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

                        if (mainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
                        {
                            KnowledgeItem knowledgeItem = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedKnowledgeItems().FirstOrDefault();
                            if (knowledgeItem.EntityLinks.FirstOrDefault() == null) break;
                            Annotation annotation = knowledgeItem.EntityLinks.FirstOrDefault().Target as Annotation;
                            if (annotation == null) break;
                            Location location = annotation.Location;

                            SwissAcademic.Citavi.Shell.Controls.Preview.PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;

                            Program.ActiveProjectShell.ShowPreviewFullScreenForm(location, previewControl, null);

                        }
                    }
                    break;
                case "SelectLinkedKnowledgeItem":
                    {
                        e.Handled = true;

                        MainForm mainForm = Program.ActiveProjectShell.PrimaryMainForm;

                        KnowledgeItem target;
                        KnowledgeItem knowledgeItem;

                        if (mainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
                        {
                            knowledgeItem = mainForm.GetSelectedQuotations().FirstOrDefault();
                        }
                        else if (mainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
                        {
                            knowledgeItem = mainForm.ActiveKnowledgeItem;
                        }
                        else
                        {
                            return;
                        }

                        if (knowledgeItem == null) return;

                        if (knowledgeItem.EntityLinks.Where(el => el.Indication == EntityLink.CommentOnQuotationIndication).Count() == 0) return;

                        if (knowledgeItem.QuotationType == QuotationType.Comment)
                        {

                            target = knowledgeItem.EntityLinks.ToList().Where(n => n != null && n.Indication == EntityLink.CommentOnQuotationIndication && n.Target as KnowledgeItem != null).ToList().FirstOrDefault().Target as KnowledgeItem;

                        }
                        else
                        {
                            target = knowledgeItem.EntityLinks.ToList().Where(n => n != null && n.Indication == EntityLink.CommentOnQuotationIndication && n.Target as KnowledgeItem != null).ToList().FirstOrDefault().Source as KnowledgeItem;
                        }

                        if (target == null) return;

                        if (mainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
                        {
                            Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
                            SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

                            Reference reference = target.Reference;
                            if (reference == null) return;

                            List<KnowledgeItem> quotations = reference.Quotations.ToList();

                            int index = quotations.FindIndex(q => q == target);

                            quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate(quotations[index]);
                        }
                        else if (mainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
                        {
                            mainForm.ActiveKnowledgeItem = target;
                        }
                        else
                        {
                            return;
                        }



                        return;
                    }
                case "SortKnowledgeItemsInSelection":
                    {
                        e.Handled = true;

                        var mainForm = (MainForm)e.Form;
                        if (mainForm.KnowledgeOrganizerFilterSet.Filters.Count() != 1 || mainForm.KnowledgeOrganizerFilterSet.Filters[0].Name == "Knowledge items without categories")
                        {
                            MessageBox.Show("You must select one category.");
                            return;
                        }

                        KnowledgeItemInSelectionSorter.SortSelectedKnowledgeItems(mainForm);
                    }
                    break;
                #endregion
                case "SortKnowledgeItemsInCategory":
                    {
                        e.Handled = true;

                        var mainForm = (MainForm)e.Form;
                        if (mainForm.KnowledgeOrganizerFilterSet.Filters.Count() != 1 || mainForm.KnowledgeOrganizerFilterSet.Filters[0].Name == "Knowledge items without categories")
                        {
                            MessageBox.Show("You must select one category.");
                            return;
                        }

                        KnowledgeItemInCategorySorter.SortKnowledgeItemsInCategorySorter(mainForm);
                    }
                    break;
            }
            base.OnBeforePerformingCommand(e);
        }
    }
}
