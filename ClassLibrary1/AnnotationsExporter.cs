using System;
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
    class AnnotationsExporter
    {
        public static void ExportAnnotations(List<Reference> references)
        {
            List<Reference> exportFailed = new List<Reference>();
            int exportCounter = 0;

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl != null)
            {
                previewControl.ShowNoPreview();
            }

            foreach (Reference reference in references)
            {
                List<KnowledgeItem> quotations = reference.Quotations.Where(q => q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                List<LinkedResource> linkedResources = reference.Locations.Where(l => l.LocationType == LocationType.ElectronicAddress && l.Address.LinkedResourceType == LinkedResourceType.AttachmentFile).Select(l => (LinkedResource)l.Address).ToList();

                List<KnowledgeItem> comments = quotations.Where(q => q.QuotationType == QuotationType.Comment && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                List<KnowledgeItem> directQuotations = quotations.Where(q => q.QuotationType == QuotationType.DirectQuotation && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                List<KnowledgeItem> indirectQuotations = quotations.Where(q => q.QuotationType == QuotationType.IndirectQuotation && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                List<KnowledgeItem> quickReferences = quotations.Where(q => q.QuotationType == QuotationType.QuickReference && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                List<KnowledgeItem> summaries = quotations.Where(q => q.QuotationType == QuotationType.Summary && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

                foreach (Location location in reference.Locations)
                {
                    foreach (Annotation annotation in location.Annotations.Where(a => a.Visible == true && a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() == 0).ToList())
                    {
                        location.Annotations.Remove(annotation);  
                    }
                    foreach (Annotation annotation in location.Annotations.Where(a => a.Visible == false && a.EntityLinks.Where(e => e.Indication == EntityLink.SourceAnnotLinkIndication).Count() == 0 && a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() == 0).ToList())
                    {
                        location.Annotations.Remove(annotation);
                    }
                }

                foreach (LinkedResource linkedResource in linkedResources)
                {
                    string pathToFile = linkedResource.Resolve().LocalPath;
                    if (!System.IO.File.Exists(pathToFile)) continue;

                    List<Rect> coveredRects = new List<Rect>();

                    Document document = new Document(pathToFile);
                    if (document == null) continue;

                    previewControl.ShowNoPreview();


                    List<KnowledgeItem> commentsAtThisLocation = comments.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

                    List<KnowledgeItem> directQuotationsAtThisLocation = directQuotations.Where(d => d.EntityLinks != null && d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication) != null && ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

                    List<KnowledgeItem> indirectQuotationsAtThisLocation = indirectQuotations.Where(d => d.EntityLinks != null && d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication) != null && ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

                    List<KnowledgeItem> quickReferencesAtThisLocation = quickReferences.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

                    List<KnowledgeItem> summariesAtThisLocation = summaries.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

                    List<Annotation> commentAnnotationsAtThisLocation = commentsAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

                    List<Annotation> directQuotationAnnotationsAtThisLocation = directQuotationsAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

                    List<Annotation> indirectQuotationAnnotationsAtThisLocation = indirectQuotationsAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

                    List<Annotation> quickReferenceAnnotationsAtThisLocation = quickReferencesAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

                    List<Annotation> summaryAnnotationsAtThisLocation = summaries.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();


                    DeleteExistingHighlights(document);

                    foreach (Annotation annotation in commentAnnotationsAtThisLocation)
                    {
                        System.Drawing.Color highlightColor = KnownColors.AnnotationComment100;
                        CreateHighlightAnnots(annotation, highlightColor, document, coveredRects, out coveredRects);
                    }

                    foreach (Annotation annotation in directQuotationAnnotationsAtThisLocation)
                    {
                        System.Drawing.Color highlightColor = KnownColors.AnnotationDirectQuotation100;
                        CreateHighlightAnnots(annotation, highlightColor, document, coveredRects, out coveredRects);
                    }

                    foreach (Annotation annotation in indirectQuotationAnnotationsAtThisLocation)
                    {
                        System.Drawing.Color highlightColor = KnownColors.AnnotationIndirectQuotation100;
                        CreateHighlightAnnots(annotation, highlightColor, document, coveredRects, out coveredRects);
                    }

                    foreach (Annotation annotation in quickReferenceAnnotationsAtThisLocation)
                    {
                        System.Drawing.Color highlightColor = KnownColors.AnnotationQuickReference100;
                        CreateHighlightAnnots(annotation, highlightColor, document, coveredRects, out coveredRects);
                    }

                    foreach (Annotation annotation in summaryAnnotationsAtThisLocation)
                    {
                        System.Drawing.Color highlightColor = KnownColors.AnnotationSummary100;
                        CreateHighlightAnnots(annotation, highlightColor, document, coveredRects, out coveredRects);
                    }

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
            string message = "Annotations exported to {0} locations.";
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

        public static void DeleteExistingHighlights(Document document)
            {
                for (int i = 1; i <= document.GetPageCount(); i++)
                {
                    pdftron.PDF.Page page = document.GetPage(i);
                    int annotCount = page.GetNumAnnots();

                    List<Annot> annotsToDelete = new List<Annot>();

                    for (int j = 0; j < page.GetNumAnnots(); j++)
                    {
                        Annot annot = page.GetAnnot(j);
                        if (annot.GetSDFObj() != null &&
                            annot.GetType() == Annot.Type.e_Highlight)
                        {
                            annotsToDelete.Add(annot);
                        }
                        else
                        {

                        }
                    }

                    foreach (Annot annot in annotsToDelete)
                    {
                        page.AnnotRemove(annot);
                    }
                }
            }

        static void CreateHighlightAnnots(Annotation annotation, System.Drawing.Color highlightColor, Document document, List<Rect> coveredRectsBefore, out List<Rect> coveredRectsAfter)
        {
            List<Quad> quads = annotation.Quads.Where(q => q.IsContainer == false).ToList();
            List<int> pages = quads.Select(q => q.PageIndex).ToList().Distinct().ToList();
            pages.Sort();

            KnowledgeItem quotation = (KnowledgeItem)annotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Source;

            string quotationText = quotation.Text;
            if (string.IsNullOrEmpty(quotationText)) quotationText = quotation.CoreStatement;
            bool TextAlreadyWrittenToAnAnnotation = false;

            coveredRectsAfter = coveredRectsBefore;

            double currentR = highlightColor.R;
            double currentG = highlightColor.G;
            double currentB = highlightColor.B;
            double currentA = highlightColor.A;

            ColorPt colorPt = new ColorPt(
                currentR / 255,
                currentG / 255,
                currentB / 255 
                );

            double highlightOpacity = currentA / 255;

            foreach (int pageIndex in pages)
            {
                pdftron.PDF.Page page = document.GetPage(pageIndex);
                if (page == null) continue;
                if (!page.IsValid()) continue;
                List<Quad> quadsOnThisPage = annotation.Quads.Where(q => q.PageIndex == pageIndex && q.IsContainer == false).Distinct().ToList();

                List<Rect> boxes = new List<Rect>();

                foreach (Quad quad in quadsOnThisPage)
                {
                    boxes.Add(new Rect(quad.MinX, quad.MinY, quad.MaxX, quad.MaxY));
                }

                Annot newAnnot = null;

                // If we want to make the later annotation invisible, we should uncomment the following if then else, and comment out the line after that

                if (boxes.Select(b => new { b.x1, b.y1, b.x2, b.y2 }).Intersect(coveredRectsAfter.Select(b => new { b.x1, b.y1, b.x2, b.y2 })).Count() == boxes.Count())
                {
                    newAnnot = CreateSingleHighlightAnnot(document, boxes, colorPt, 0);
                }
                else
                {
                    newAnnot = CreateSingleHighlightAnnot(document, boxes, colorPt, highlightOpacity);
                }

                // newAnnot = CreateSingleHighlightAnnot(document, boxes, colorPt, highlightOpacity);

                if (!TextAlreadyWrittenToAnAnnotation)
                {
                    newAnnot.SetContents(quotationText);
                    TextAlreadyWrittenToAnAnnotation = true;
                }

                page.AnnotPushBack(newAnnot);

                if (newAnnot.IsValid() == false) continue;
                if (newAnnot.GetAppearance() == null)
                {
                    newAnnot.RefreshAppearance();
                }

                coveredRectsAfter.AddRange(boxes);
            }
        }

        static Annot CreateSingleHighlightAnnot(Document document, List<Rect> boxes, ColorPt colorPt, double highlightOpacity)
        {
            Annot annotation = Annot.Create(document, Annot.Type.e_Highlight, RectangleUnion(boxes));

            annotation.SetColor(colorPt);

            annotation.SetBorderStyle(new Annot.BorderStyle
            (Annot.BorderStyle.Style.e_solid, 1));

            Obj quads = annotation.GetSDFObj().PutArray("QuadPoints");
            List<Rect> lineRectangles = boxes.GroupBy(box => box.y1).Select(line => RectangleUnion(line.ToList())).ToList();
            lineRectangles.ForEach(lineRect => PushBackBox(quads, lineRect));
            annotation.SetAppearance(CreateHighlightAppearance(lineRectangles, colorPt, highlightOpacity, document));

            return annotation;
        }

        static void PushBackBox(Obj quads, Rect box)
        {
            quads.PushBackNumber(box.x1);
            quads.PushBackNumber(box.y2);
            quads.PushBackNumber(box.x2);
            quads.PushBackNumber(box.y2);
            quads.PushBackNumber(box.x1);
            quads.PushBackNumber(box.y1);
            quads.PushBackNumber(box.x2);
            quads.PushBackNumber(box.y1);
        }

        static Obj CreateHighlightAppearance(List<Rect> boxes, ColorPt highlightColor, double highlightOpacity, Document document)
        {
            var elementBuilder = new ElementBuilder();
            elementBuilder.PathBegin();

            boxes.ForEach(box => elementBuilder.Rect(box.x1 - 2, box.y1, box.x2 - box.x1, box.y2 - box.y1));

            Element element = elementBuilder.PathEnd();

            element.SetPathFill(true);
            element.SetPathStroke(false);

            GState elementGraphicState = element.GetGState();
            elementGraphicState.SetFillColorSpace(ColorSpace.CreateDeviceRGB());
            elementGraphicState.SetFillColor(highlightColor);
            elementGraphicState.SetFillOpacity(highlightOpacity);
            elementGraphicState.SetBlendMode(GState.BlendMode.e_bl_multiply);

            var elementWriter = new ElementWriter();
            elementWriter.Begin(document);

            elementWriter.WriteElement(element);
            Obj highlightAppearance = elementWriter.End();

            elementBuilder.Dispose();
            elementWriter.Dispose();

            Rect boundingBox = RectangleUnion(boxes);
            highlightAppearance.PutRect("BBox", boundingBox.x1, boundingBox.y1, boundingBox.x2, boundingBox.y2);
            highlightAppearance.PutName("Subtype", "Form");

            return highlightAppearance;
        }

        static Rect RectangleUnion(List<Rect> boxes)
        {
            double y1 = boxes.Select(box => box.y1).Min();
            double y2 = boxes.Select(box => box.y2).Max();
            double x1 = boxes.Select(box => box.x1).Min();
            double x2 = boxes.Select(box => box.x2).Max();
            return new Rect(x1, y1, x2, y2);
        }     
    }
}
