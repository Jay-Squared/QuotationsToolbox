using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Controls.WindowsForms;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using pdftron;
using pdftron.Common;
using pdftron.Filters;
using pdftron.FDF;
using pdftron.PDF;
using pdftron.PDF.Annots;
using pdftron.SDF;

namespace QuotationsToolbox
{
    static class AnnotationsImporter
    {
        public static void AnnotationsImport(QuotationType quotationType)
        {
            // Static Variables

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            Document document = pdfViewControl.Document;
            if (document == null) return;

            Location location = previewControl.ActiveLocation;
            if (location == null) return;

            List<Reference> references = Program.ActiveProjectShell.Project.References.Where(r => r.Locations.Contains(location)).ToList();
            if (references == null) return;
            if (references.Count != 1)
            {
                MessageBox.Show("The document is not linked with exactly one reference. Import aborted.");
            }

            Reference reference = references.FirstOrDefault();
            if (references == null) return;

            LinkedResource linkedResource = location.Address;

            string pathToFile = linkedResource.Resolve().LocalPath;

            PreviewBehaviour previewBehaviour = location.PreviewBehaviour;

            // Dynamic Variables

            string colorPickerCaption = string.Empty;

            List<ColorPt> selectedColorPts = new List<ColorPt>();

            int annotationsImportedCount = 0;

            bool ImportEmptyAnnotations = false;
            bool RedrawAnnotations = true;

            // The Magic

            Form commentAnnotationsColorPicker = new AnnotationsImporterColorPicker(quotationType, document.ExistingColors(), out selectedColorPts);

            DialogResult dialogResult = commentAnnotationsColorPicker.ShowDialog();

            RedrawAnnotations = AnnotationsImporterColorPicker.RedrawAnnotationsSelected;
            ImportEmptyAnnotations = AnnotationsImporterColorPicker.ImportEmptyAnnotationsSelected;

            if (dialogResult == DialogResult.Cancel)
            {                
                MessageBox.Show("Import of external highlights cancelled.");
                return;
            }

            List<Annotation> temporaryAnnotations = new List<Annotation>();

            if (ImportEmptyAnnotations || quotationType == QuotationType.Comment)
            { 
                for (int pageIndex = 1; pageIndex <= document.GetPageCount(); pageIndex++)
                {
                    pdftron.PDF.Page page = document.GetPage(pageIndex);
                    if (page.IsValid())
                    {
                        List<Annot> annotsToDelete = new List<Annot>();
                        for (int j = 0; j < page.GetNumAnnots(); j++)
                        {
                            Annot annot = page.GetAnnot(j);

                            if (annot.GetSDFObj() != null && (annot.GetType() == Annot.Type.e_Highlight || annot.GetType() == Annot.Type.e_Unknown))
                            {
                                Highlight highlight = new Highlight(annot);
                                if (highlight == null) continue;

                                ColorPt annotColorPt = annot.GetColorAsRGB();

                                if (selectedColorPts.Where(e =>
                                        e.Get(0) == annotColorPt.Get(0)
                                        && e.Get(1) == annotColorPt.Get(1)
                                        && e.Get(2) == annotColorPt.Get(2)
                                        && e.Get(3) == annotColorPt.Get(3)
                                    ).Count() == 0)
                                {
                                    continue;
                                }
                                Annotation temporaryAnnotation = highlight.TemporaryAnnotation();
                                location.Annotations.Add(temporaryAnnotation);
                                temporaryAnnotations.Add(temporaryAnnotation);
                            }
                        } // end for (int j = 1; j <= page.GetNumAnnots(); j++)
                    } // end if (page.IsValid())
                } // end for (int i = 1; i <= document.GetPageCount(); i++)

                previewControl.ShowNoPreview();
                location.PreviewBehaviour = PreviewBehaviour.SkipEntryPage;
                previewControl.ShowLocationPreview(location);
                document = new Document(pathToFile);
            }

            //Uncomment here to get an overview of all annotations
            //int x = 0;
            //foreach (Annotation a in location.Annotations)
            //{
            //    System.Diagnostics.Debug.WriteLine("Annotation " + x.ToString());
            //    int y = 0;
            //    foreach (Quad q in a.Quads)
            //    {
            //        System.Diagnostics.Debug.WriteLine("Quad " + y.ToString());
            //        System.Diagnostics.Debug.WriteLine("IsContainer: " + q.IsContainer.ToString());
            //        System.Diagnostics.Debug.WriteLine("MinX: " + q.MinX.ToString());
            //        System.Diagnostics.Debug.WriteLine("MinY: " + q.MinY.ToString());
            //        System.Diagnostics.Debug.WriteLine("MaxX: " + q.MaxX.ToString());
            //        System.Diagnostics.Debug.WriteLine("MaxY: " + q.MaxY.ToString());
            //        y = y + 1;
            //    }
            //    x = x + 1;
            //}

            int v = 0;
            for (int pageIndex = 1; pageIndex <= document.GetPageCount(); pageIndex++)
            {
                pdftron.PDF.Page page = document.GetPage(pageIndex);
                if (page.IsValid())
                {
                    List<Annot> annotsToDelete = new List<Annot>();
                    for (int j = 0; j < page.GetNumAnnots(); j++)
                    {
                        Annot annot = page.GetAnnot(j);

                        if (annot.GetSDFObj() != null && (annot.GetType() == Annot.Type.e_Highlight || annot.GetType() == Annot.Type.e_Unknown))
                        {
                            Highlight highlight = new Highlight(annot);
                            if (highlight == null) continue;

                            ColorPt annotColorPt = annot.GetColorAsRGB();

                            if (selectedColorPts.Where(e =>
                                    e.Get(0) == annotColorPt.Get(0)
                                    && e.Get(1) == annotColorPt.Get(1)
                                    && e.Get(2) == annotColorPt.Get(2)
                                    && e.Get(3) == annotColorPt.Get(3)
                                ).Count() == 0)
                            {
                                continue;
                            }

                            // Uncomment here to get an overview of all highlights
                            //System.Diagnostics.Debug.WriteLine("Highlight " + v.ToString());
                            //int w = 0;
                            //foreach (Quad q in highlight.AsAnnotationQuads())
                            //{
                            //    System.Diagnostics.Debug.WriteLine("Quad " + w.ToString());
                            //    System.Diagnostics.Debug.WriteLine("IsContainer: " + q.IsContainer.ToString());
                            //    System.Diagnostics.Debug.WriteLine("MinX: " + q.MinX.ToString());
                            //    System.Diagnostics.Debug.WriteLine("MinY: " + q.MinY.ToString());
                            //    System.Diagnostics.Debug.WriteLine("MaxX: " + q.MaxX.ToString());
                            //    System.Diagnostics.Debug.WriteLine("MaxY: " + q.MaxY.ToString());
                            //    w = w + 1;
                            //}
                            //v = v + 1;

                            if (highlight.UpdateExistingQuotation(quotationType, location))
                            {
                                annotationsImportedCount++;
                                continue;
                            }
                            if (highlight.CreateNewQuotationAndAnnotationFromHighlight(quotationType, ImportEmptyAnnotations, RedrawAnnotations, temporaryAnnotations))
                            {
                                annotationsImportedCount++;
                                continue;
                            }
                        }
                    } // end for (int j = 1; j <= page.GetNumAnnots(); j++)
                    foreach (Annot annot in annotsToDelete)
                    {
                        page.AnnotRemove(annot);
                    }
                } // end if (page.IsValid())
            } // end for (int i = 1; i <= document.GetPageCount(); i++)

            foreach (Annotation annotation in temporaryAnnotations)
            {
                location.Annotations.Remove(annotation);
            }

            location.PreviewBehaviour = previewBehaviour;

            MessageBox.Show(annotationsImportedCount.ToString() + " highlights have been imported.");
        }

        public static List<ColorPt> ExistingColors(this Document document)
        {
            List<ColorPt> existingColorPts = new List<ColorPt>();
            for (int i = 1; i <= document.GetPageCount(); i++)
            {
                pdftron.PDF.Page page = document.GetPage(i);
                if (page.IsValid())
                {
                    for (int j = 0; j < page.GetNumAnnots(); j++)
                    {
                        Annot annot = page.GetAnnot(j);

                        if (annot.GetSDFObj() != null && (annot.GetType() == Annot.Type.e_Highlight || annot.GetType() == Annot.Type.e_Unknown))
                        {
                            ColorPt existingColorPt = annot.GetColorAsRGB();
                            if (existingColorPts.Where(e =>
                                    e.Get(0) == existingColorPt.Get(0)
                                    && e.Get(1) == existingColorPt.Get(1)
                                    && e.Get(2) == existingColorPt.Get(2)
                                    && e.Get(3) == existingColorPt.Get(3)
                                ).Count() == 0)
                            {
                                existingColorPts.Add(existingColorPt);
                            }
                        }
                    }
                }
            }
            return existingColorPts;
        }

        public static bool UpdateExistingQuotation(this Highlight highlight, QuotationType quotationType, Location location)
        {
            // Static Variables

            Reference reference = location.Reference;

            List<Annotation> annotationsAtThisLocation = location.Annotations.ToList();

            List<KnowledgeItem> quotations = reference.Quotations.Where(q => q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

            List<KnowledgeItem> quotationsOfRelevantQuotationType = quotations.Where(q => q.QuotationType == quotationType && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

            List<KnowledgeItem> quotationsOfRelevantQuotationTypeAtThisLocation = quotationsOfRelevantQuotationType.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location == location).ToList();

            List<Annotation> annotationsOfRelevantTypeAtThisLocation = quotationsOfRelevantQuotationTypeAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

            List<Annotation> equivalentAnnotationsWithQuotation = highlight.EquivalentAnnotations().Where(a => a.EntityLinks != null && a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

            if (equivalentAnnotationsWithQuotation == null) return false;

            Annotation existingAnnotation = equivalentAnnotationsWithQuotation.Intersect(annotationsOfRelevantTypeAtThisLocation).FirstOrDefault();
            if (existingAnnotation == null) return false;

            // Dynamic Variables

            string highlightContents = highlight.GetContents();

            // The Magic

            if (!string.IsNullOrEmpty(highlightContents))
            {
                if (existingAnnotation != null)
                {
                    if (existingAnnotation.EntityLinks != null && existingAnnotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0)
                    {
                        KnowledgeItem existingQuotation = (KnowledgeItem)existingAnnotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Source;

                        switch (quotationType)
                        {
                            case QuotationType.QuickReference:
                                existingQuotation.CoreStatement = highlightContents;
                                break;
                            default:
                                existingQuotation.Text = highlightContents;
                                break;
                        }
                        return true;
                    }
                }
            }
            else
            {
                return true;
            }
            return false;
        }

        public static bool CreateNewQuotationAndAnnotationFromHighlight(this Highlight highlight, QuotationType quotationType, bool ImportEmptyAnnotations, bool RedrawAnnotations, List<Annotation> temporaryAnnotations)
        {
            string highlightContents = highlight.GetContents();
            if (string.IsNullOrEmpty(highlightContents) && !ImportEmptyAnnotations) return false;
          
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return false;

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return false;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return false;

            Document document = pdfViewControl.Document;
            if (document == null) return false;

            Location location = previewControl.ActiveLocation;
            if (location == null) return false;

            Reference reference = location.Reference;
            if (reference == null) return false;


            // Dynamic variables

            KnowledgeItem newQuotation = null;
            KnowledgeItem newDirectQuotation = null;

            TextContent textContent = null;

            // Does any other annotation with the same quads already exist?

            Annotation existingAnnotation = highlight.EquivalentAnnotations().FirstOrDefault();

            if ((string.IsNullOrEmpty(highlightContents) && ImportEmptyAnnotations) || quotationType == QuotationType.Comment)

            {
                Annotation temporaryAnnotation = temporaryAnnotations.Where(a => !highlight.AsAnnotationQuads().TemporaryQuads().Except(a.Quads.ToList()).Any()).FirstOrDefault();
                if (temporaryAnnotation != null)
                {
                    pdfViewControl.GoToAnnotation(temporaryAnnotation);
                    textContent = (TextContent)pdfViewControl.GetSelectedContentFromType(pdfViewControl.GetSelectedContentType(), -1, false, true);
                    location.Annotations.Remove(temporaryAnnotation);
                }
                else
                {
                    return false;
                }
            }

            int startPage = 1;

            if (reference.PageRange.StartPage.Number != null) startPage = reference.PageRange.StartPage.Number.Value;

            string pageRangeString = (startPage + highlight.GetPage().GetIndex() - 1).ToString();

            Annotation knowledgeItemIndicationAnnotation = highlight.CreateKnowledgeItemIndicationAnnotation(RedrawAnnotations);

            switch (quotationType)
            {
                case QuotationType.Comment:
                    if (!string.IsNullOrEmpty(highlightContents))
                    {
                        newQuotation = highlightContents.CreateNewQuotationFromHighlightContents(pageRangeString, reference, quotationType);
                        reference.Quotations.Add(newQuotation);
                        project.AllKnowledgeItems.Add(newQuotation);
                        newQuotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);

                        if (AnnotationsImporterColorPicker.ImportDirectQuotationLinkedWithCommentSelected)
                        {
                            Annotation newDirectQuotationIndicationAnnotation = highlight.CreateKnowledgeItemIndicationAnnotation(RedrawAnnotations);
                            newDirectQuotation = textContent.CreateNewQuotationFromAnnotationContent(pageRangeString, reference, QuotationType.DirectQuotation);
                            reference.Quotations.Add(newDirectQuotation);
                            project.AllKnowledgeItems.Add(newDirectQuotation);
                            newDirectQuotation.LinkWithKnowledgeItemIndicationAnnotation(newDirectQuotationIndicationAnnotation);

                            EntityLink commentDirectQuotationLink = new EntityLink(project);
                            commentDirectQuotationLink.Source = newQuotation;
                            commentDirectQuotationLink.Target = newDirectQuotation;
                            commentDirectQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                            project.EntityLinks.Add(commentDirectQuotationLink);

                            newQuotation.CoreStatement = newDirectQuotation.CoreStatement + " (Comment)";
                            newQuotation.CoreStatementUpdateType = UpdateType.Manual;
                        }
                    }
                    else if (string.IsNullOrEmpty(highlightContents) && ImportEmptyAnnotations)
                    {
                        newQuotation = textContent.CreateNewQuotationFromAnnotationContent(pageRangeString, reference, quotationType);
                        reference.Quotations.Add(newQuotation);
                        project.AllKnowledgeItems.Add(newQuotation);
                        newQuotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);

                        if (AnnotationsImporterColorPicker.ImportDirectQuotationLinkedWithCommentSelected)
                        {
                            Annotation newDirectQuotationIndicationAnnotation = highlight.CreateKnowledgeItemIndicationAnnotation(RedrawAnnotations);
                            newDirectQuotation = textContent.CreateNewQuotationFromAnnotationContent(pageRangeString, reference, QuotationType.DirectQuotation);
                            reference.Quotations.Add(newDirectQuotation);
                            project.AllKnowledgeItems.Add(newDirectQuotation);
                            newDirectQuotation.LinkWithKnowledgeItemIndicationAnnotation(newDirectQuotationIndicationAnnotation);

                            EntityLink commentDirectQuotationLink = new EntityLink(project);
                            commentDirectQuotationLink.Source = newQuotation;
                            commentDirectQuotationLink.Target = newDirectQuotation;
                            commentDirectQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                            project.EntityLinks.Add(commentDirectQuotationLink);

                            newQuotation.CoreStatement = newDirectQuotation.CoreStatement + " (Comment)";
                            newQuotation.CoreStatementUpdateType = UpdateType.Manual;
                        }
                    }
                    break;
                case QuotationType.QuickReference:
                    if (!string.IsNullOrEmpty(highlightContents))
                    {
                        newQuotation = highlightContents.CreateNewQuickReferenceFromHighlightContents(pageRangeString, reference, quotationType);
                    }
                    else if (string.IsNullOrEmpty(highlightContents) && ImportEmptyAnnotations)
                    {
                        newQuotation = textContent.CreateNewQuickReferenceFromAnnotationContent(pageRangeString, reference, quotationType);
                    }
                    reference.Quotations.Add(newQuotation);
                    project.AllKnowledgeItems.Add(newQuotation);
                    newQuotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);
                    break;
                default:
                    if (!string.IsNullOrEmpty(highlightContents))
                    {
                        newQuotation = highlightContents.CreateNewQuotationFromHighlightContents(pageRangeString, reference, quotationType);
                        reference.Quotations.Add(newQuotation);
                        project.AllKnowledgeItems.Add(newQuotation);
                        newQuotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);
                    }
                    else if (string.IsNullOrEmpty(highlightContents) && ImportEmptyAnnotations)
                    {
                        newQuotation = textContent.CreateNewQuotationFromAnnotationContent(pageRangeString, reference, quotationType);
                        reference.Quotations.Add(newQuotation);
                        project.AllKnowledgeItems.Add(newQuotation);
                        newQuotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);
                    }
                    break;
            }

            if (existingAnnotation != null)
            {
                existingAnnotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);
                existingAnnotation.Visible = false;
            }

            return true;

        }

        public static KnowledgeItem CreateNewQuotationFromHighlightContents(this string highlightContents, string pageRangeString, Reference reference, QuotationType quotationType)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            KnowledgeItem newQuotation = new KnowledgeItem(reference, quotationType);

            newQuotation.Text = highlightContents;

            newQuotation.PageRange = pageRangeString;

            return newQuotation;
        }

        public static KnowledgeItem CreateNewQuotationFromAnnotationContent(this TextContent textContent, string pageRangeString, Reference reference, QuotationType quotationType)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            KnowledgeItem newQuotation = new KnowledgeItem(reference, quotationType);

            newQuotation.TextRtf = textContent.Rtf;

            newQuotation.PageRange = pageRangeString;

            return newQuotation;
        }

        public static KnowledgeItem CreateNewQuickReferenceFromHighlightContents(this string highlightContents, string pageRangeString, Reference reference, QuotationType quotationType)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            KnowledgeItem newQuotation = new KnowledgeItem(reference, quotationType);

            newQuotation.CoreStatement = highlightContents;

            newQuotation.PageRange = pageRangeString;
            
            return newQuotation;
        }

        public static KnowledgeItem CreateNewQuickReferenceFromAnnotationContent(this TextContent textContent, string pageRangeString, Reference reference, QuotationType quotationType)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            KnowledgeItem newQuotation = new KnowledgeItem(reference, quotationType);

            newQuotation.CoreStatement = textContent.Text;

            newQuotation.PageRange = pageRangeString;

            return newQuotation;
        }

        public static List<Annotation> EquivalentAnnotations(this Highlight highlight)
        {
            List<Annotation> equivalentAnnotations = new List<Annotation>();

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return null;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return null;

            Document document = pdfViewControl.Document;
            if (document == null) return null;

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            Location location = previewControl.ActiveLocation;
            if (location == null) return null;

            List<Annotation> annotationsAtThisLocation = location.Annotations.ToList();
            if (annotationsAtThisLocation == null) return null;

            equivalentAnnotations.AddRange(annotationsAtThisLocation.Where(a => !highlight.AsAnnotationQuads().RoundedQuads().Except(a.Quads.ToList().RoundedQuads()).Any()));
            equivalentAnnotations.AddRange(annotationsAtThisLocation.Where(a => !highlight.AsAnnotationQuads().SimpleQuads().RoundedQuads().Except(a.Quads.ToList().RoundedQuads()).Any()));

            return equivalentAnnotations;
        }

        public static List<Quad> AsAnnotationQuads(this Highlight highlight)
        {
            int pageIndex = highlight.GetPage().GetIndex();

            List<Quad> highlightAsAnnotationQuads = new List<Quad>();

            for (int l = 0; l < highlight.GetQuadPointCount(); l++)
            {
                List<double> xValues = new List<double>();
                List<double> yValues = new List<double>();

                QuadPoint quadPoint = highlight.GetQuadPoint(l);

                xValues.AddRange(new List<double> { quadPoint.p1.x, quadPoint.p2.x, quadPoint.p3.x, quadPoint.p4.x });
                yValues.AddRange(new List<double> { quadPoint.p1.y, quadPoint.p2.y, quadPoint.p3.y, quadPoint.p4.y });

                highlightAsAnnotationQuads.Add(new Quad(pageIndex, xValues.Min(), yValues.Min(), xValues.Max(), yValues.Max()));
            }

            return highlightAsAnnotationQuads;
        }

        public static Annotation CreateKnowledgeItemIndicationAnnotation(this Highlight highlight, bool RedrawAnnotations)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return null;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return null;

            Document document = pdfViewControl.Document;
            if (document == null) return null;

            Location location = previewControl.ActiveLocation;
            if (location == null) return null;

            Reference reference = location.Reference;
            if (reference == null) return null;

            Annotation newAnnotation = new Annotation(location);

            List<Quad> quads = new List<Quad>();

            if (RedrawAnnotations)
            {
                quads = highlight.AsAnnotationQuads().SimpleQuads();
            }
            else
            {
                quads = highlight.AsAnnotationQuads();
            }


            newAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
            newAnnotation.Quads = quads;
            newAnnotation.Visible = false;

            location.Annotations.Add(newAnnotation);

            return newAnnotation;
        }
               
        public static void LinkWithKnowledgeItemIndicationAnnotation(this KnowledgeItem quotation, Annotation annotation)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return;

            EntityLink newEntityLink = new EntityLink(project);
            newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
            newEntityLink.Source = quotation;
            newEntityLink.Target = annotation;
            project.EntityLinks.Add(newEntityLink);
        }

        public static void LinkWithKnowledgeItemIndicationAnnotation(this Annotation existingAnnotation, Annotation newAnnotation)
        {
            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return;

            EntityLink sourceAnnotLink = new EntityLink(project);
            sourceAnnotLink.Indication = EntityLink.SourceAnnotLinkIndication;
            sourceAnnotLink.Source = existingAnnotation;
            sourceAnnotLink.Target = newAnnotation;
            project.EntityLinks.Add(sourceAnnotLink);
        }

        public static Annotation TemporaryAnnotation(this Highlight highlight)
        {
            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return null;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return null;

            Document document = pdfViewControl.Document;
            if (document == null) return null;

            Location location = previewControl.ActiveLocation;
            if (location == null) return null;

            Annotation temporaryAnnotation = new Annotation(location);

            temporaryAnnotation.Quads = highlight.AsAnnotationQuads().TemporaryQuads();
            temporaryAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
            temporaryAnnotation.Visible = true;

            return temporaryAnnotation;
        }

        public static List<Quad> TemporaryQuads(this List<Quad> quads)
        {
            //List<Quad> tempQuads = new List<Quad>();

            //quads = quads.OrderByDescending(q => MidPoint(q.MaxY, q.MinY)).ToList();

            //double maxX = quads.Select(q => q.MaxX).Max();
            //double minX = quads.Select(q => q.MinX).Min();

            //if (quads.Count > 2)
            //{
            //    tempQuads.Add(quads[0]);
            //    for (int i = 1; i < quads.Count - 1; i++)
            //    {
            //        if ((quads[i].MaxX - quads[i].MinX) > 0.25 * (maxX - minX)) tempQuads.Add(quads[i]);
            //    }
            //    tempQuads.Add(quads[quads.Count - 1]);
            //    quads = tempQuads;
            //}

            //List<Quad> temporaryAnnotationQuads = new List<Quad>();

            //int l = quads.Count - 1;

            //double scalingFactor = 0.5;

            //double averageLineHeight = 0;
            //double yValuesOffset = 0;

            //if (quads.Count == 0) return null;

            //if (quads.Count == 1)
            //{
            //    averageLineHeight = (quads[0].MaxY - quads[0].MinY);
            //    yValuesOffset = (averageLineHeight - averageLineHeight * scalingFactor) / 2;
            //}
            //else
            //{
            //    averageLineHeight = (MidPoint(quads[0].MaxY, quads[0].MinY) - MidPoint(quads[l].MaxY, quads[l].MinY)) / (l);
            //    yValuesOffset = (averageLineHeight - averageLineHeight * scalingFactor) / 2;
            //}

            //switch (quads.Count)
            //{
            //    case 1:
            //        temporaryAnnotationQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].X1, MidPoint(quads[0].MaxY, quads[0].MinY) - yValuesOffset / 2, quads[0].X2, MidPoint(quads[0].MaxY, quads[0].MinY) + yValuesOffset));
            //        break;
            //    case 2:
            //        if (MidPoint(quads[l].MaxY, quads[l].MinY) > quads[0].MinY)
            //        {
            //            temporaryAnnotationQuads.Add(new Quad(quads[0].PageIndex, false, minX, quads[l].MinY, maxX, quads[0].MaxY));
            //        }
            //        else
            //        {
            //            temporaryAnnotationQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].X1, quads[0].MinY + yValuesOffset, quads[0].X2, quads[0].MaxY - yValuesOffset));
            //            temporaryAnnotationQuads.Add(new Quad(quads[l].PageIndex, false, quads[l].X1, MidPoint(quads[l].MaxY, quads[l].MinY) - averageLineHeight / 2, quads[l].X2, quads[l].MaxY - yValuesOffset));
            //        }
            //        break;
            //    default:
            //        quads[1] = quads[1];

            //        temporaryAnnotationQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].X1, quads[0].MinY + yValuesOffset, maxX, quads[0].MaxY- yValuesOffset));
            //        temporaryAnnotationQuads.Add(new Quad(quads[1].PageIndex, false, minX, quads[l-1].MinY + yValuesOffset, maxX, quads[1].MaxY - yValuesOffset));
            //        temporaryAnnotationQuads.Add(new Quad(quads[l].PageIndex, false, minX, MidPoint(quads[l].MaxY, quads[l].MinY) - averageLineHeight/2, quads[l].X2, quads[l].MaxY - yValuesOffset));

            //        break;
            //}
            //return temporaryAnnotationQuads;

            List<Quad> newQuads = new List<Quad>();
            List<Quad> tempQuads = new List<Quad>();

            quads = quads.OrderByDescending(q => q.MinY).ToList();

            double maxX = quads.Select(q => q.MaxX).Max();
            double minX = quads.Select(q => q.MinX).Min();

            double scalingFactor = 0.5;
            double averageLineHeight = 0;
            double yValuesOffset = 0;

            if (quads.Count == 0) return null;

            if (quads.Count > 2)
            {
                if (MidPoint(quads[0].MinY, quads[0].MaxY) > quads[1].MaxY) tempQuads.Add(quads[0]);
                for (int i = 1; i < quads.Count - 1; i++)
                {
                    if ((quads[i].MaxX - quads[i].MinX) > 0.50 * (maxX - minX)) tempQuads.Add(quads[i]);
                }
                tempQuads.Add(quads[quads.Count - 1]);
                quads = tempQuads;
            }

            int l = quads.Count - 1;

            if (quads.Count == 1)
            {
                averageLineHeight = (quads[0].MaxY - quads[0].MinY);
                yValuesOffset = (averageLineHeight - averageLineHeight * (1- scalingFactor)) / 2;
            }
            else
            {
                averageLineHeight = (MidPoint(quads[0].MaxY, quads[0].MinY) - MidPoint(quads[l].MaxY, quads[l].MinY)) / (l);
                yValuesOffset = (averageLineHeight * scalingFactor) / 2;
            }

            quads = quads.OrderByDescending(q => MidPoint(q.MaxY, q.MinY)).ToList();

            switch (quads.Count)
            {
                case 0:
                    break;
                case 1:
                    newQuads = quads;
                    break;
                case 2:
                    if (MidPoint(quads[1].MaxY, quads[1].MinY) > quads[0].MinY)
                    {
                        newQuads.Add(new Quad(quads[0].PageIndex, false, minX, quads[1].MinY, maxX, quads[0].MaxY));
                    }
                    else
                    {
                        newQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].MinX, MidPoint(quads[0].MinY, quads[0].MaxY) - yValuesOffset, maxX, MidPoint(quads[0].MinY, quads[0].MaxY) + yValuesOffset));
                        newQuads.Add(new Quad(quads[l].PageIndex, false, minX, MidPoint(quads[l].MinY, quads[l].MaxY) - yValuesOffset, quads[l].MaxX, MidPoint(quads[l].MinY, quads[l].MaxY) + yValuesOffset));
                    }
                    break;
                default:
                    newQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].MinX, quads[0].MinY, maxX, MidPoint(quads[0].MaxY, quads[0].MinY) + yValuesOffset));

                    for (int i = 1; i < l; i++)
                    {
                        newQuads.Add(new Quad(quads[i].PageIndex, false, minX, MidPoint(quads[i].MaxY, quads[i].MinY) - yValuesOffset, maxX, MidPoint(quads[i].MaxY, quads[i].MinY) + yValuesOffset));
                    }

                    newQuads.Add(new Quad(quads[l].PageIndex, false, minX, quads[l].MinY, quads[l].MaxX, MidPoint(quads[l].MinY, quads[l].MaxY) + yValuesOffset));
                    break;
            }
            newQuads = newQuads.OrderByDescending(q => MidPoint(q.MaxY, q.MinY)).ToList();

            return newQuads;
        }

        public static List<Quad> SimpleQuads(this List<Quad> quads)
        {
            List<Quad> newQuads = new List<Quad>();
            List<Quad> tempQuads = new List<Quad>();

            quads = quads.OrderByDescending(q => q.MinY).ToList();

            double maxX = quads.Select(q => q.MaxX).Max();
            double minX = quads.Select(q => q.MinX).Min();

            if (quads.Count == 0) return null;

            if (quads.Count > 2)
            {
                for (int i = 0; i < quads.Count; i++)
                {
                    if (i > 0 && i < quads.Count - 2 && (quads[i].MaxX - quads[i].MinX) < 0.51 * (maxX - minX)) continue;
                    if (tempQuads.Where(q => q.MaxY < quads[i].MaxY).Any()) continue;

                    tempQuads.Add(quads[i]);
                }
                quads = tempQuads;
            }

            quads = quads.OrderByDescending(q => q.QuadMidPoint()).ToList();

            int l = quads.Count;

            double minY = quads.Select(q => q.MinY).Min();
            double maxY = quads.Select(q => q.MaxY).Max();
            double height = (maxY - minY) / l;

            switch (quads.Count)
            {
                case 0:
                    break;
                case 1:
                    newQuads = quads;
                    break;
                default:
                    newQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].MinX, maxY - height, maxX, maxY ));
                    for (int i = 1; i < l -1; i++)
                    {
                        newQuads.Add(new Quad(quads[i].PageIndex, false, minX, maxY - (i + 1) * height, maxX, maxY - (i) * height ));
                    }
                    newQuads.Add(new Quad(quads[l-1].PageIndex, false, minX, minY, quads[l-1].MaxX, minY + height));

                    break;
            }
            newQuads = newQuads.OrderByDescending(q => q.QuadMidPoint()).ToList();
            return newQuads;
        }

        public static double MidPoint(double firstValue, double secondValue)
        {
            double halfwayPoint = 0;

            halfwayPoint = Math.Round(((firstValue + secondValue) / 2), 2);

            return halfwayPoint;
        }
        public static double QuadMidPoint(this Quad quad)
        {
            return MidPoint(quad.MaxY, quad.MinY);
        }
        public static List<Quad> RoundedQuads(this List<Quad> quads)
        {
            int p = 0;
            List<Quad> newQuads = new List<Quad>();
            foreach (Quad q in quads)
            {
                newQuads.Add(new Quad(q.PageIndex, q.IsContainer, Math.Round(q.MinX, p), Math.Round(q.MinY, p), Math.Round(q.MaxX, p), Math.Round(q.MaxY, p)));
            }
            quads = newQuads;
            return quads;
        }
    }
}
