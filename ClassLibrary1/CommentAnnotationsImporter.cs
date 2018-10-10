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
    class CommentAnnotationsImporter
    {
        public static void ConvertComments(Reference reference)
        {
            
            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            Document document = previewControl.GetDocument();
            if (document == null) return;

            Project project = reference.Project;
            if (project == null) return;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return;

            List<PageWidth> store = new List<PageWidth>();

            int startPageInt = 1;

            if (reference.PageRange.StartPage.Number != null) startPageInt = reference.PageRange.StartPage.Number.Value;

            List<Annot> annotationsAtThisLocationToDelete = new List<Annot>();

            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return;

            PdfViewControl pdfViewControl = propertyInfo.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl) as PdfViewControl;

            Content contentx = pdfViewControl.GetSelectedContentFromType(pdfViewControl.GetSelectedContentType(), -1, false, true);
            SwissAcademic.Citavi.Controls.Wpf.TextContent textContentx = contentx as TextContent;

            List<Annotation> annotationsAtThisLocation = location.Annotations.ToList();

            List<KnowledgeItem> quotations = reference.Quotations.Where(q => q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

            List<KnowledgeItem> comments = quotations.Where(q => q.QuotationType == QuotationType.Comment && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

            List<KnowledgeItem> directQuotations = quotations.Where(q => q.QuotationType == QuotationType.DirectQuotation && q.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0).ToList();

            LinkedResource linkedResource = location.Address;

            string pathToFile = linkedResource.Resolve().LocalPath;

            List<KnowledgeItem> commentsAtThisLocation = comments.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

            List<Annotation> commentAnnotationsAtThisLocation = commentsAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();

            List<KnowledgeItem> directQuotationsAtThisLocation = directQuotations.Where(d => ((Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).Location.Address == linkedResource).ToList();

            List<Annotation> directQuotationAnnotationsAtThisLocation = directQuotationsAtThisLocation.Select(d => (Annotation)d.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).ToList().FirstOrDefault().Target).ToList();


            List<ColorPt> existingColorPts = new List<ColorPt>();

            for (int i = 1; i <= document.GetPageCount(); i++)
            {
                pdftron.PDF.Page page = document.GetPage(i);
                if (page.IsValid())
                {

                    for (int j = 1; j <= page.GetNumAnnots(); j++)
                    {
                        Annot annot = page.GetAnnot(j);
                        if (annot.GetSDFObj() != null && annot.GetType() == Annot.Type.e_Highlight)
                        {
                            ColorPt existingColorPt = annot.GetColorAsRGB();
                            if (existingColorPts.Where(e =>
                                e.Get(0) == existingColorPt.Get(0)
                                && e.Get(1) == existingColorPt.Get(1)
                                && e.Get(2) == existingColorPt.Get(2)
                                && e.Get(3) == existingColorPt.Get(3)
                            ).Count() == 0)
                            {
                                existingColorPts.Add(annot.GetColorAsRGB());
                            }

                        }
                    }
                }
            }

            List<ColorPt> selectedColorPts = new List<ColorPt>();

            Form commentAnnotationsColorPicker = new CommentAnnotationsColorPicker(existingColorPts, out selectedColorPts);

            DialogResult dialogResult = commentAnnotationsColorPicker.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
            
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                MessageBox.Show("Importing comments aborted.");
                return;
            }

            for (int i = 1; i <= document.GetPageCount(); i++)
            {
                pdftron.PDF.Page page = document.GetPage(i);
                if (page.IsValid())
                {

                    for (int j = 1; j <= page.GetNumAnnots(); j++)
                    {
                        Annot annot = page.GetAnnot(j);

                        if (annot.GetSDFObj() != null && annot.GetType() == Annot.Type.e_Highlight)
                        {
                            Highlight highlight = new Highlight(annot);
                            if (string.IsNullOrEmpty(highlight.GetContents())) continue;

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
                            else
                            {
                            }                            

                            List<Quad> commentAnnotationQuads = new List<Quad>();

                            for (int l = 0; l < highlight.GetQuadPointCount(); l++)
                            {

                                List<double> xValues = new List<double>();
                                List<double> yValues = new List<double>();

                                QuadPoint quadPoint = highlight.GetQuadPoint(l);

                                Quad quad = new Quad(highlight.GetPage().GetIndex(), quadPoint.p1.x, quadPoint.p1.y, quadPoint.p2.x, quadPoint.p2.y);

                                xValues.AddRange(new List<double> { quadPoint.p1.x, quadPoint.p2.x, quadPoint.p3.x, quadPoint.p4.x });
                                yValues.AddRange(new List<double> { quadPoint.p1.y, quadPoint.p2.y, quadPoint.p3.y, quadPoint.p4.y });

                                commentAnnotationQuads.Add(new Quad(i, xValues.Min(), yValues.Min(), xValues.Max(), yValues.Max()));
                            }

                            // Does a comment annotation with the same quads already exist? If so, just update its comment's text and go to the next PDFNet Annot

                            Annotation existingCommentAnnotation = commentAnnotationsAtThisLocation.Where(c => !c.Quads.ToList().Except(commentAnnotationQuads).Any()).FirstOrDefault();

                            if (existingCommentAnnotation != null)
                            {
                                if (existingCommentAnnotation.EntityLinks != null && existingCommentAnnotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0)
                                {
                                    KnowledgeItem existingComment = (KnowledgeItem)existingCommentAnnotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Source;
                                    existingComment.Text = highlight.GetContents();
                                    annotationsAtThisLocationToDelete.Add(annot);
                                    continue;
                                }
                            }

                            // There was no equivalent comment annotation, so we have to create one

                            Annotation newCommentAnnotation = new Annotation(location);

                            newCommentAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
                            newCommentAnnotation.Quads = commentAnnotationQuads;
                            newCommentAnnotation.Visible = false;
                            location.Annotations.Add(newCommentAnnotation);

                            KnowledgeItem comment = new KnowledgeItem(reference, QuotationType.Comment);
                            comment.Text = highlight.GetContents();
                            comment.CoreStatementUpdateType = UpdateType.Automatic;
                            reference.Quotations.Add(comment);

                            EntityLink commentAnnotationLink = new EntityLink(project);
                            commentAnnotationLink.Source = comment;

                            commentAnnotationLink.Target = newCommentAnnotation;
                            commentAnnotationLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                            project.EntityLinks.Add(commentAnnotationLink);

                            annotationsAtThisLocationToDelete.Add(annot);

                            // Check if there already is a Citavi annotation with the same quads that we can link the new comment to. Since we already checked for comment annotations above, this should only find non-comment annotations

                            Annotation annotation = annotationsAtThisLocation.Where(a => !a.Quads.ToList().Except(commentAnnotationQuads).Any()).FirstOrDefault();

                            if (annotation == null) continue;

                            // If that annotation is already linked to a knowledge item, we link the new comment to that knowledge item

                            if (annotation.EntityLinks != null && annotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0)
                            {
                                EntityLink commentDirectQuotationLink = new EntityLink(project);
                                commentDirectQuotationLink.Source = comment;
                                commentDirectQuotationLink.Target = (KnowledgeItem)annotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Target;
                                commentDirectQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                                project.EntityLinks.Add(commentDirectQuotationLink);
                            }
                            // If the corresponding Citavi annotation is not linked to a knowledge item, we create a direct quotation from the annotation and link the comment to that one
                            else
                            {
                                pdfViewControl.GoToAnnotation(annotation);

                                List<Quad> annotationQuads = annotation.Quads.ToList();
                                Content content = pdfViewControl.GetSelectedContentFromType(pdfViewControl.GetSelectedContentType(), -1, false, true);
                                SwissAcademic.Citavi.Controls.Wpf.TextContent textContent = content as TextContent;

                                KnowledgeItem newDirectQuotation = new KnowledgeItem(reference, QuotationType.DirectQuotation);
                                List<int> pages = new List<int>();
                                List<Quad> newQuads = annotation.Quads.ToList();

                                foreach (Quad quad in annotationQuads)
                                {
                                    Quad newQuad = new Quad(quad.PageIndex, true, quad.X1, quad.Y1, quad.X2, quad.Y1);
                                    newQuads.Add(newQuad);

                                    pages.Add(startPageInt + quad.PageIndex - 1);
                                }

                                annotation.Visible = false;

                                Annotation newAnnotation = new Annotation(location);

                                newAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
                                newAnnotation.Quads = newQuads;
                                newAnnotation.Visible = false;

                                location.Annotations.Add(newAnnotation);

                                EntityLink sourceAnnotLink = new EntityLink(project);
                                sourceAnnotLink.Indication = EntityLink.SourceAnnotLinkIndication;
                                sourceAnnotLink.Source = annotation;
                                sourceAnnotLink.Target = newAnnotation;
                                project.EntityLinks.Add(sourceAnnotLink);

                                if (pages.Min() == pages.Max())
                                {
                                    newDirectQuotation.PageRange = pages.Min().ToString();
                                }
                                else
                                {
                                    newDirectQuotation.PageRange = pages.Min().ToString() + "-" + pages.Max().ToString();
                                }

                                newDirectQuotation.TextRtf = textContent.Rtf;

                                reference.Quotations.Add(newDirectQuotation);
                                project.AllKnowledgeItems.Add(newDirectQuotation);

                                EntityLink newEntityLink = new EntityLink(project);
                                newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                                newEntityLink.Source = newDirectQuotation;
                                newEntityLink.Target = newAnnotation;
                                project.EntityLinks.Add(newEntityLink);

                                EntityLink commentDirectQuotationLink = new EntityLink(project);
                                commentDirectQuotationLink.Source = comment;
                                commentDirectQuotationLink.Target = newDirectQuotation;
                                commentDirectQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                                project.EntityLinks.Add(commentDirectQuotationLink);

                                comment.CoreStatement = newDirectQuotation.CoreStatement + " (Comment)";
                                comment.PageRange = newDirectQuotation.PageRange;
                            }
                        }
                    } // end for (int j = 1; j <= page.GetNumAnnots(); j++)
                    foreach (Annot annotation in annotationsAtThisLocationToDelete)
                    {
                        page.AnnotRemove(annotation);
                    }
                } // end if (page.IsValid())
            } // for (int i = 1; i <= document.GetPageCount(); i++)
            try
            {
                document.Save(pathToFile, SDFDoc.SaveOptions.e_remove_unused);
            }
            catch
            {

            }
        }
    }
}
