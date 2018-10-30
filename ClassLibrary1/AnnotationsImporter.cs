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
        public static void AnnotationsImport(Reference reference, QuotationType quotationType)
        {
            // Static Variables

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            Document document = previewControl.GetDocument();
            if (document == null) return;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return;

            LinkedResource linkedResource = location.Address;

            string pathToFile = linkedResource.Resolve().LocalPath;

            PreviewBehaviour previewBehaviour = location.PreviewBehaviour;

            // Dynamic Variables

            string colorPickerCaption = string.Empty;

            List<ColorPt> selectedColorPts = new List<ColorPt>();

            int annotationsImportedCount = 0;

            bool ImportEmptyAnnotations = false;
            bool RedrawAnnotations = false;

            // The Magic

            switch (quotationType)
            {
                case QuotationType.Comment:
                    colorPickerCaption = "Select the colors of all comments.";
                    break;
                case QuotationType.DirectQuotation:
                    colorPickerCaption = "Select the colors of all direct quotations.";
                    ImportEmptyAnnotations = true;
                    break;
                case QuotationType.QuickReference:
                    colorPickerCaption = "Select the colors of all quick references.";
                    ImportEmptyAnnotations = true;
                    break;
                case QuotationType.Summary:
                    colorPickerCaption = "Select the colors of all summaries.";
                    break;
                default:
                    break;
            }

            Form commentAnnotationsColorPicker = new AnnotationsImporterColorPicker(colorPickerCaption, document.ExistingColors(), ImportEmptyAnnotations, out selectedColorPts);

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

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return false;

            Document document = previewControl.GetDocument();
            if (document == null) return false;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return false;

            Reference reference = location.Reference;
            if (reference == null) return false;

            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return false;

            PdfViewControl pdfViewControl = propertyInfo.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl) as PdfViewControl;
            
            // Dynamic variables

            KnowledgeItem newQuotation = null;
            KnowledgeItem newDirectQuotation = null;

            TextContent textContent = null;

            // Does any other annotation with the same quads already exist?

            Annotation existingAnnotation = highlight.EquivalentAnnotations().FirstOrDefault();

            if ((string.IsNullOrEmpty(highlightContents) && ImportEmptyAnnotations) || quotationType == QuotationType.Comment)

            {
                Annotation temporaryAnnotation = temporaryAnnotations.Where(a => !highlight.EquivalentAnnotationQuads().TemporaryQuads().Except(a.Quads.ToList()).Any()).FirstOrDefault();
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
                        Annotation newDirectQuotationIndicationAnnotation = highlight.CreateKnowledgeItemIndicationAnnotation(RedrawAnnotations);
                        newDirectQuotation = textContent.CreateNewQuotationFromAnnotationContent(pageRangeString, reference, QuotationType.DirectQuotation);
                        reference.Quotations.Add(newDirectQuotation);
                        project.AllKnowledgeItems.Add(newDirectQuotation);
                        newDirectQuotation.LinkWithKnowledgeItemIndicationAnnotation(newDirectQuotationIndicationAnnotation);

                        newQuotation = highlightContents.CreateNewQuotationFromHighlightContents(pageRangeString, reference, quotationType);
                        reference.Quotations.Add(newQuotation);
                        project.AllKnowledgeItems.Add(newQuotation);
                        newQuotation.LinkWithKnowledgeItemIndicationAnnotation(knowledgeItemIndicationAnnotation);

                        EntityLink commentDirectQuotationLink = new EntityLink(project);
                        commentDirectQuotationLink.Source = newQuotation;
                        commentDirectQuotationLink.Target = newDirectQuotation;
                        commentDirectQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                        project.EntityLinks.Add(commentDirectQuotationLink);

                        newQuotation.CoreStatement = newDirectQuotation.CoreStatement + " (Comment)";
                        newQuotation.CoreStatementUpdateType = UpdateType.Manual;
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

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return null;

            Document document = previewControl.GetDocument();
            if (document == null) return null;

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return null;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return null;

            List<Annotation> annotationsAtThisLocation = location.Annotations.ToList();
            if (annotationsAtThisLocation == null) return null;

            equivalentAnnotations = annotationsAtThisLocation.Where(a => !highlight.EquivalentAnnotationQuads().Except(a.Quads.ToList()).Any()).ToList();
            equivalentAnnotations.AddRange(annotationsAtThisLocation.Where(a => !highlight.EquivalentAnnotationQuads().SimpleQuads().Except(a.Quads.ToList()).Any()).ToList());

            return equivalentAnnotations;
        }

        public static List<Quad> EquivalentAnnotationQuads(this Highlight highlight)
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

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return null;

            Document document = previewControl.GetDocument();
            if (document == null) return null;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return null;

            Reference reference = location.Reference;
            if (reference == null) return null;

            Annotation newAnnotation = new Annotation(location);

            List<Quad> quads = new List<Quad>();

            if (RedrawAnnotations)
            {
                quads = highlight.EquivalentAnnotationQuads().SimpleQuads();
            }
            else
            {
                quads = highlight.EquivalentAnnotationQuads();
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
            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return null;

            Document document = previewControl.GetDocument();
            if (document == null) return null;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return null;

            Annotation temporaryAnnotation = new Annotation(location);

            temporaryAnnotation.Quads = highlight.EquivalentAnnotationQuads().TemporaryQuads();
            temporaryAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
            temporaryAnnotation.Visible = true;

            return temporaryAnnotation;
        }

        public static List<Quad> TemporaryQuads(this List<Quad> quads)
        {
            List<Quad> tempQuads = new List<Quad>();

            quads = quads.OrderByDescending(q => q.MinY).ToList();

            double maxX = quads.Select(q => q.MaxX).Max();
            double minX = quads.Select(q => q.MinX).Min();

            if (quads.Count > 2)
            {
                tempQuads.Add(quads[0]);
                for (int i = 1; i < quads.Count - 1; i++)
                {
                    if ((quads[i].MaxX - quads[i].MinX) > 0.25 * (maxX - minX)) tempQuads.Add(quads[i]);
                }
                tempQuads.Add(quads[quads.Count - 1]);
                quads = tempQuads;
            }

            List<Quad> temporaryAnnotationQuads = new List<Quad>();

            double scalingFactor = 0.5;

            Quad firstQuad = new Quad();
            Quad secondQuad = new Quad();
            Quad secondToLastQuad = new Quad();
            Quad lastQuad = new Quad();

            double firstQuadMidY = 0;
            double lastQuadMidY = 0;
            double averageLineHeight = 0;
            double yValuesOffset = 0;

            if (quads.Count == 0) return null;
            if (quads.Count >= 1)
            {
                firstQuad = quads[0];
                lastQuad = quads[quads.Count - 1];
                firstQuadMidY = (firstQuad.MaxY + firstQuad.MinY) / 2;
                lastQuadMidY = (lastQuad.MaxY + lastQuad.MinY) / 2;
            }
            if (quads.Count == 1)
            {
                averageLineHeight = (firstQuad.MaxY - firstQuad.MinY);
                yValuesOffset = (averageLineHeight - averageLineHeight * scalingFactor) / 2;
            }
            else
            {
                averageLineHeight = (firstQuadMidY - lastQuadMidY) / (quads.Count - 1);
                yValuesOffset = (averageLineHeight - averageLineHeight * scalingFactor) / 2;
            }

            switch (quads.Count)
            {
                case 1:
                    temporaryAnnotationQuads.Add(new Quad(firstQuad.PageIndex, false, firstQuad.X1, firstQuadMidY - yValuesOffset / 2, firstQuad.X2, firstQuadMidY + yValuesOffset));
                    break;
                case 2:
                    temporaryAnnotationQuads.Add(new Quad(firstQuad.PageIndex, false, firstQuad.X1, firstQuad.MinY + yValuesOffset, firstQuad.X2, firstQuad.MaxY - yValuesOffset));
                    temporaryAnnotationQuads.Add(new Quad(lastQuad.PageIndex, false, lastQuad.X1, lastQuadMidY - averageLineHeight / 2, lastQuad.X2, lastQuad.MaxY - yValuesOffset));
                    break;
                default:
                    secondQuad = quads[1];
                    secondToLastQuad = quads[quads.Count - 2];

                    temporaryAnnotationQuads.Add(new Quad(firstQuad.PageIndex, false, firstQuad.X1, firstQuad.MinY + yValuesOffset, firstQuad.X2, firstQuad.MaxY- yValuesOffset));
                    temporaryAnnotationQuads.Add(new Quad(secondQuad.PageIndex, false, minX, secondToLastQuad.MinY + yValuesOffset, maxX, secondQuad.MaxY - yValuesOffset));
                    temporaryAnnotationQuads.Add(new Quad(lastQuad.PageIndex, false, lastQuad.X1, lastQuadMidY - averageLineHeight/2, lastQuad.X2, lastQuad.MaxY - yValuesOffset));

                    break;
            }
            return temporaryAnnotationQuads;
        }

        public static List<Quad> SimpleQuads(this List<Quad> quads)
        {
            // return quads;
            List<Quad> newQuads = new List<Quad>();
            List<Quad> tempQuads = new List<Quad>();

            quads = quads.OrderByDescending(q => q.MinY).ToList();

            double maxX = quads.Select(q => q.MaxX).Max();
            double minX = quads.Select(q => q.MinX).Min();

            if (quads.Count > 2)
            {
                tempQuads.Add(quads[0]);
                for (int i = 1; i < quads.Count - 1; i++)
                {
                    if ((quads[i].MaxX - quads[i].MinX) > 0.30 * (maxX - minX)) tempQuads.Add(quads[i]);
                }
                tempQuads.Add(quads[quads.Count - 1]);
                quads = tempQuads;
            }

            switch (quads.Count)
            {
                case 0:
                    break;
                case 1:
                    newQuads = quads;
                    break;
                case 2:
                    newQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].MinX, MidPoint(quads[0].MaxY, quads[1].MinY), quads[0].MaxX, quads[0].MaxY));
                    newQuads.Add(new Quad(quads[1].PageIndex, false, quads[1].MinX, quads[1].MinY, quads[1].MaxX, MidPoint(quads[0].MaxY, quads[1].MinY)));
                    break;
                default:
                    newQuads.Add(new Quad(quads[0].PageIndex, false, quads[0].MinX, MidPoint(quads[0].MinY, quads[1].MaxY), maxX, quads[0].MaxY));

                    for (int i = 1; i < quads.Count - 1; i++)
                    {
                        newQuads.Add(new Quad(quads[i].PageIndex, false, minX, MidPoint(quads[i - 1].MaxY, quads[i].MinY), maxX, MidPoint(quads[i].MaxY, quads[i + 1].MinY)));
                    }

                    newQuads.Add(new Quad(quads[quads.Count - 1].PageIndex, false, minX, quads[quads.Count - 1].MinY, quads[quads.Count - 1].MaxX, MidPoint(quads[quads.Count - 2].MaxY, quads[quads.Count - 1].MinY)));
                    break;
            }
            return newQuads;
        }

        public static double MidPoint(double firstValue, double secondValue)
        {
            double halfwayPoint = 0;

            halfwayPoint = Math.Round(((firstValue + secondValue) / 2), 2);

            return halfwayPoint;
        }
    }
}
