﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Drawing;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using pdftron.PDF;
using pdftron.PDF.Annots;
using pdftron.SDF;

namespace QuotationsToolbox
{
    class QuickReferenceBookmarkExporter
    {
        public static void ExportBookmarks(List<Reference> references)
        {
            List<Reference> exportFailed = new List<Reference>();
            int exportCounter = 0;

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl != null)
            {
                previewControl.ShowNoPreview();
            }

            foreach (Reference reference in references)
            {
                List<KnowledgeItem> quotations = reference.Quotations.Where(q => q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                List<LinkedResource> linkedResources = reference.Locations.Where(l => l.LocationType == LocationType.ElectronicAddress).Select(l => (LinkedResource)l.Address).ToList();

                List<KnowledgeItem> quickReferences = quotations.Where(q => q.QuotationType == QuotationType.QuickReference && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();


                foreach (LinkedResource linkedResource in linkedResources)
                {
                    string pathToFile = linkedResource.Resolve().LocalPath;
                    if (!System.IO.File.Exists(pathToFile)) continue;

                    List<Rect> coveredRects = new List<Rect>();

                    Document document = new Document(pathToFile);
                    if (document == null) continue;


                    List<KnowledgeItem> quickReferencesAtThisLocation = quickReferences.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();


                    List<Annotation> quickReferenceAnnotationsAtThisLocation = quickReferencesAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

                   foreach (Annotation annotation in quickReferenceAnnotationsAtThisLocation)
                    {

                        ExportQuickReferenceAsBookmark(annotation, document);
                    }

                    DeleteExtraBookmarks(quickReferencesAtThisLocation.Select(q => q.CoreStatement).ToList(), document);

                    try
                    {
                        document.Save(pathToFile, SDFDoc.SaveOptions.e_remove_unused);
                        exportCounter++;
                    }
                    catch
                    {
                        exportFailed.Add(reference);
                    }
                    document.Close();
                }

            }
            string message = "Bookmarks exported to {0} locations.";
            message = string.Format(message, exportCounter);
            if (exportFailed.Count == 0)
            {
                MessageBox.Show(message, "Citavi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                DialogResult showFailed = MessageBox.Show(message + "\n Would you like to show a selection of references where export has failed?", "Citavi Macro", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (showFailed == DialogResult.Yes)
                {
                    var filter = new ReferenceFilter(exportFailed, "Export failed", false);
                    Program.ActiveProjectShell.PrimaryMainForm.ReferenceEditorFilterSet.Filters.ReplaceBy(new List<ReferenceFilter> { filter });
                }
            }
        }
        public static void ExportQuickReferenceAsBookmark(Annotation annotation, Document document)
        {
            KnowledgeItem quickReference = (KnowledgeItem)annotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Source;

            string bookmarkTitle = quickReference.CoreStatement;

            pdftron.PDF.Bookmark bookmark = pdftron.PDF.Bookmark.Create(document, bookmarkTitle);

            Quad firstQuad = annotation.Quads.FirstOrDefault();
            int page = firstQuad.PageIndex;
            double left = firstQuad.MinX;
            double top = firstQuad.MaxY;

            Destination destination = Destination.CreateXYZ(document.GetPage(page), left, top, 0);

            pdftron.PDF.Bookmark foo = document.GetFirstBookmark().Find(bookmarkTitle);

            if (foo.IsValid())
            {
                return;
            }
            
            document.AddRootBookmark(bookmark);

            bookmark.SetAction(pdftron.PDF.Action.CreateGoto(destination));

        }
        public static void DeleteExtraBookmarks(List<string> quickreferences, Document document)
        {
            List<pdftron.PDF.Bookmark> bookmarks = new List<pdftron.PDF.Bookmark>();

            pdftron.PDF.Bookmark bookmark = document.GetFirstBookmark();

            while (bookmark.IsValid())
            {
                bookmarks.Add(bookmark);
                bookmark = bookmark.GetNext();

            }

            foreach (pdftron.PDF.Bookmark b in bookmarks)
            {
                if (!quickreferences.Contains(b.GetTitle())) b.Delete();
            }

        }
    }
}
