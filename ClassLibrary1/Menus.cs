using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Citavi.Shell.Controls.SmartRepeaters;
using SwissAcademic.Controls;
using SwissAcademic.Pdf;

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

        protected override void OnHostingFormLoaded(Form hostingForm)
        {
            MainForm mainForm = (MainForm)hostingForm;

            mainForm.ActiveWorkspaceChanged += PrimaryMainForm_ActiveWorkspaceChanged;

            // Reference Editor Reference Menu

            var referencesMenu = mainForm.GetMainCommandbarManager().GetReferenceEditorCommandbar(MainFormReferenceEditorCommandbarId.Menu).GetCommandbarMenu(MainFormReferenceEditorCommandbarMenuId.References);

            var commandbarButtonMoveAttachment = referencesMenu.AddCommandbarButton("MoveAttachment", "Move local attachments of selected references to a different folder");
            commandbarButtonMoveAttachment.HasSeparator = true;

            var commandbarButtonExportAnnotations = referencesMenu.AddCommandbarButton("ExportAnnotations", "Export quotations in selected references as PDF highlights");
            commandbarButtonExportAnnotations.HasSeparator = true;
            var commandbarButtonExportBookmarks = referencesMenu.AddCommandbarButton("ExportBookmarks", "Export quick references in selected references as PDF bookmarks");

            // Preview Tool Menu

            var previewCommandbarMenuTools = (mainForm.GetPreviewCommandbar(MainFormPreviewCommandbarId.Toolbar).GetCommandbarMenu(MainFormPreviewCommandbarMenuId.Tools));

            var annotationsImportCommandbarMenu = previewCommandbarMenuTools.AddCommandbarMenu("AnnotationsImportCommandbarMenu", "Import annotations…", CommandbarItemStyle.Default);

            annotationsImportCommandbarMenu.HasSeparator = true;

            annotationsImportCommandbarMenu.AddCommandbarButton("ImportDirectQuotations", "Import direct quotations in active document");
            annotationsImportCommandbarMenu.AddCommandbarButton("ImportIndirectQuotations", "Import indirect quotations in active document");
            annotationsImportCommandbarMenu.AddCommandbarButton("ImportComments", "Import comments in active document");
            annotationsImportCommandbarMenu.AddCommandbarButton("ImportQuickReferences", "Import quick references in active document");
            annotationsImportCommandbarMenu.AddCommandbarButton("ImportSummaries", "Import summaries in active document");

            var commandBarButtonMergeAnnotations = previewCommandbarMenuTools.AddCommandbarButton("MergeAnnotations", "Merge annotations", CommandbarItemStyle.Default);
            commandBarButtonMergeAnnotations.HasSeparator = true;

            var commandBarButtonRedrawAnnotations = previewCommandbarMenuTools.AddCommandbarButton("SimplifyAnnotations", "Redraw annotations", CommandbarItemStyle.Default);

            // Quotations Pop-Up Menu

            var referenceEditorQuotationsContextMenu = CommandbarMenu.Create(mainForm.GetReferenceEditorQuotationsCommandbarManager().ToolbarsManager.Tools["ReferenceEditorQuotationsContextMenu"] as PopupMenuTool);

            var positionContextMenu = referenceEditorQuotationsContextMenu.AddCommandbarMenu("PositionContextMenu", "Quotation Position", CommandbarItemStyle.Default);
            positionContextMenu.HasSeparator = true;

            var commandbarButtonKnowledgeItemsSortInReference = positionContextMenu.AddCommandbarButton("KnowledgeItemsSortInReference", "Sort selected quotations by position in PDF");
            commandbarButtonKnowledgeItemsSortInReference.Shortcut = Shortcut.CtrlK;
            var commandbarButtonPageAssignFromPositionInPDF = positionContextMenu.AddCommandbarButton("PageAssignFromPositionInPDF", "Assign a page range to selected quote based on the PDF position");
            commandbarButtonPageAssignFromPositionInPDF.HasSeparator = true;
            positionContextMenu.AddCommandbarButton("PageAssignFromPreviousQuotation", "Assign page range and numbering type from previous quote to selected quote");
            var commandbarButtonShowQuotationAndSetPageRangeManually = positionContextMenu.AddCommandbarButton("ShowQuotationAndSetPageRangeManually", "Assign page range manually after showing selected quote in PDF");
            commandbarButtonShowQuotationAndSetPageRangeManually.Shortcut = Shortcut.CtrlDel;

            var commentsContextMenu = referenceEditorQuotationsContextMenu.AddCommandbarMenu("CommentsContextMenu", "Comments", CommandbarItemStyle.Default);

            var commandbarButtonCommentAnnotation = commentsContextMenu.AddCommandbarButton("CommentAnnotation", "Link comment to same text in PDF as related quote");
            commentsContextMenu.AddCommandbarButton("LinkQuotations", "Link selected quote and comment");
            var commandbarButtonCreateCommentOnQuotation = commentsContextMenu.AddCommandbarButton("CreateCommentOnQuotation", "Create comment on selected quote");
            commandbarButtonCreateCommentOnQuotation.Shortcut = Shortcut.CtrlShiftK;
            var commandbarButtonSelectLinkedKnowledgeItem = commentsContextMenu.AddCommandbarButton("SelectLinkedKnowledgeItem", "Jump to related quote or comment");
            commandbarButtonSelectLinkedKnowledgeItem.Shortcut = Shortcut.AltRightArrow;

            var quickReferencesContextMenu = referenceEditorQuotationsContextMenu.AddCommandbarMenu("QuickReferencesContextMenu", "Quick References", CommandbarItemStyle.Default);
            quickReferencesContextMenu.AddCommandbarButton("QuickReferenceTitleCase", "Write core statement in title case in selected quick references");
            var commandbarButtonConvertDirectQuoteToRedHighlight = quickReferencesContextMenu.AddCommandbarButton("ConvertDirectQuoteToRedHighlight", "Convert selected quotations to quick references");


            var commandbarButtonCleanQuotationsText = referenceEditorQuotationsContextMenu.AddCommandbarButton("CleanQuotationsText", "Clean text of selected quotations");
            commandbarButtonCleanQuotationsText.Shortcut = Shortcut.ShiftDel;
            var commandbarButtonQuotationsMerge = referenceEditorQuotationsContextMenu.AddCommandbarButton("QuotationsMerge", "Merge selected quotations");
            var commandbarButtonCreateSummaryOnQuotations = referenceEditorQuotationsContextMenu.AddCommandbarButton("CreateSummaryOnQuotations", "Create summary of selected quotes");

            // Reference Editor Attachment Pop-Up Menu

            var referenceEditorUriLocationsContextMenu = CommandbarMenu.Create(mainForm.GetReferenceEditorElectronicLocationsCommandbarManager().ToolbarsManager.Tools["ReferenceEditorUriLocationsContextMenu"] as PopupMenuTool);
            
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

            // Fin

            base.OnHostingFormLoaded(hostingForm);
        }

        protected override void OnBeforePerformingCommand(SwissAcademic.Controls.BeforePerformingCommandEventArgs e)
        {
            switch (e.Key)
            {
                #region Annotation-based menus
                case "ImportComments":
                    {
                        e.Handled = true;
                        Reference reference = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().FirstOrDefault();
                        AnnotationsImporter.AnnotationsImport(QuotationType.Comment);
                    }
                    break;
                case "ImportDirectQuotations":
                    {
                        e.Handled = true;
                        Reference reference = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().FirstOrDefault();
                        AnnotationsImporter.AnnotationsImport(QuotationType.DirectQuotation);
                    }
                    break;
                case "ImportIndirectQuotations":
                    {
                        e.Handled = true;
                        Reference reference = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().FirstOrDefault();
                        AnnotationsImporter.AnnotationsImport(QuotationType.IndirectQuotation);
                    }
                    break;
                case "ImportQuickReferences":
                    {
                        e.Handled = true;
                        Reference reference = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().FirstOrDefault();
                        AnnotationsImporter.AnnotationsImport(QuotationType.QuickReference);
                    }
                    break;
                case "ImportSummaries":
                    {
                        e.Handled = true;
                        Reference reference = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().FirstOrDefault();
                        AnnotationsImporter.AnnotationsImport(QuotationType.Summary);
                    }
                    break;

                case "MergeAnnotations":
                    {
                        e.Handled = true;
                        AnnotationsAndQuotationsMerger.MergeAnnotations();
                    }
                    break;
                case "SimplifyAnnotations":
                    {
                        e.Handled = true;
                        AnnotationSimplifier.SimplifyAnnotations();
                    }
                    break;
                #endregion            
                #region Reference-based commands
                case "ExportAnnotations":
                    {
                        e.Handled = true;
                        List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().ToList();
                        AnnotationsExporter.ExportAnnotations(references);
                    }
                    break;
                case "ExportBookmarks":
                    {
                        e.Handled = true;
                        List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences().ToList();
                        QuickReferenceBookmarkExporter.ExportBookmarks(references);
                    }
                    break;
                case "MoveAttachment":
                    {
                        e.Handled = true;
                        LocationMover.MoveAttachment();
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
                case "CreateSummaryOnQuotations":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        SummaryCreator.CreatesummaryOnQuotations(quotations);
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
                case "LinkQuotations":
                    {
                        e.Handled = true;
                        List<KnowledgeItem> quotations = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().ToList();
                        QuotationLinker.LinkQuotations(quotations);
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
                        AnnotationsAndQuotationsMerger.MergeQuotations(quotations);
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
                #region ReferenceEditorUriLocationsPopupMenu

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
                    #endregion
            }
            base.OnBeforePerformingCommand(e);
        }

        void PrimaryMainForm_ActiveWorkspaceChanged(object o, EventArgs a)
        {
            if (Program.ActiveProjectShell.PrimaryMainForm.ActiveWorkspace == MainFormWorkspace.KnowledgeOrganizer)
            {
                SmartRepeater<KnowledgeItem> KnowledgeItemSmartRepeater = (SmartRepeater<KnowledgeItem>)Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("SmartRepeater", true).FirstOrDefault();

                QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("knowledgeItemPreviewSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;

                KnowledgeItemSmartRepeater.ActiveListItemChanged += KnowledgeItemPreviewSmartRepeater_ActiveListItemChanged;
            }
            else if (Program.ActiveProjectShell.PrimaryMainForm.ActiveWorkspace == MainFormWorkspace.ReferenceEditor)
            {
                QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault() as QuotationSmartRepeater;

                quotationSmartRepeaterAsQuotationSmartRepeater.ActiveListItemChanged += QuotationSmartRepeater_ActiveListItemChanged;
            }
        }

        void QuotationSmartRepeater_ActiveListItemChanged(object o, EventArgs a)
        {
            if (Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().Count == 0) return;

            KnowledgeItem activeQuotation = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().FirstOrDefault();
            if (activeQuotation.EntityLinks == null) return;
            if (activeQuotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() == 0) return;

            Annotation annotation = activeQuotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Target as Annotation;

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return;

            SwissAcademic.Citavi.Controls.Wpf.PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            pdfViewControl.GoToAnnotation(annotation);

        }

        void KnowledgeItemPreviewSmartRepeater_ActiveListItemChanged(object o, EventArgs a)
        {
            if (Program.ActiveProjectShell.PrimaryMainForm.GetSelectedKnowledgeItems().Count == 0) return;

            KnowledgeItem activeQuotation = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedKnowledgeItems().FirstOrDefault();
            if (activeQuotation.EntityLinks == null) return;
            if (activeQuotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() == 0) return;

            Annotation annotation = activeQuotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Target as Annotation;
            
            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            Document document = pdfViewControl.Document;
            if (document == null) return;

            if (previewControl.ActiveLocation != annotation.Location)
            {
                Program.ActiveProjectShell.ShowPreviewFullScreenForm(annotation.Location, previewControl, null);
            }
            pdfViewControl.GoToAnnotation(annotation);

            Program.ActiveProjectShell.PrimaryMainForm.Activate();
        }
    }
}
