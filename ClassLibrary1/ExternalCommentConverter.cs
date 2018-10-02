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
using pdftron.Filters;
using pdftron.FDF;
using pdftron.Common;
using pdftron.SDF;

using pdftron.PDF;
using pdftron.PDF.Annots;

namespace QuotationsToolbox
{
    class ExternalCommentConverter
    {
        public static void ConvertComments(Reference reference)
        {
            var pdfLocations = reference.GetPDFLocations();

            List<PageWidth> store = new List<PageWidth>();

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            Document document = previewControl.GetDocument();
            if (document == null) return;

            Project project = reference.Project;
            if (project == null) return;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return;

            int startPageInt = 1;

            if (reference.PageRange.StartPage.Number != null) startPageInt = reference.PageRange.StartPage.Number.Value;

            int overall_num_annots = 0;
            List<Annot> annotationsToDelete = new List<Annot>();

            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return;

            PdfViewControl pdfViewControl = propertyInfo.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl) as PdfViewControl;

            Content contentx = pdfViewControl.GetSelectedContentFromType(pdfViewControl.GetSelectedContentType(), -1, false, true);
            SwissAcademic.Citavi.Controls.Wpf.TextContent textContentx = contentx as TextContent;
            System.Diagnostics.Debug.WriteLine(textContentx.Text);

            List<Annotation> annotations = location.Annotations.ToList();

            for (int i = 1; i <= document.GetPageCount(); i++)
            {
                pdftron.PDF.Page page = document.GetPage(i);
                if (page.IsValid())
                {
                    overall_num_annots = overall_num_annots + page.GetNumAnnots();
                    for (int j = 1; j <= page.GetNumAnnots(); j++)
                    {
                        Annot annot = page.GetAnnot(j);
                        if (annot.GetSDFObj() != null && annot.GetType() == Annot.Type.e_Highlight)
                        {
                            Highlight highlight = new Highlight(annot);

                            System.Diagnostics.Debug.WriteLine("PDFTron Highlight");

                            if (string.IsNullOrEmpty(highlight.GetContents())) continue;

                            List<Quad> commentAnnotationQuads = new List<Quad>();

                            for (int l = 0; l < highlight.GetQuadPointCount(); l++)
                            {

                                List<double> xValues = new List<double>();
                                List<double> yValues = new List<double>();


                                QuadPoint quadPoint = highlight.GetQuadPoint(l);

                                System.Diagnostics.Debug.WriteLine("QUAD POINT");

                                System.Diagnostics.Debug.WriteLine("p1.x: " + quadPoint.p1.x + ", p1.y: " + quadPoint.p1.y
                                    + "; p2.x: " + quadPoint.p2.x + ", p2.y: " + quadPoint.p2.y
                                    + "; p3.x: " + quadPoint.p3.x + ", p3.y: " + + quadPoint.p3.y
                                    + "; p4.x: " + quadPoint.p4.x + ", p4.y: " + quadPoint.p4.y);

                                Quad quad = new Quad(highlight.GetPage().GetIndex(), quadPoint.p1.x, quadPoint.p1.y, quadPoint.p2.x, quadPoint.p2.y);

                                xValues.AddRange(new List<double> { quadPoint.p1.x, quadPoint.p2.x, quadPoint.p3.x, quadPoint.p4.x });
                                yValues.AddRange(new List<double> { quadPoint.p1.y, quadPoint.p2.y, quadPoint.p3.y, quadPoint.p4.y });

                                commentAnnotationQuads.Add(new Quad(i, xValues.Min(), yValues.Min(), xValues.Max(), yValues.Max()));
                            }

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

                            // Now let's look at the corresponding Citavi annotation


                            Annotation annotation = annotations.Where(a => !a.Quads.ToList().Except(commentAnnotationQuads).Any()).FirstOrDefault();

                            if (annotation == null) continue;

                            // If it is already linked to a knowledge item, we link the new comment to that knowledge item

                            if (annotation.EntityLinks != null && annotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() > 0)
                            {
                                EntityLink commentQuotationLink = new EntityLink(project);
                                commentQuotationLink.Source = comment;
                                commentQuotationLink.Target = (KnowledgeItem)annotation.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Target;
                                commentQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                                project.EntityLinks.Add(commentQuotationLink);
                            }
                            // If the corresponding Citavi annotation is not linked to a knowledge item, we create a direct quotation and link the comment to that one
                            else
                            { 
                                pdfViewControl.GoToAnnotation(annotation);

                                List<Quad> annotationQuads = annotation.Quads.ToList();
                                Content content = pdfViewControl.GetSelectedContentFromType(pdfViewControl.GetSelectedContentType(), -1, false, true);
                                SwissAcademic.Citavi.Controls.Wpf.TextContent textContent = content as TextContent;

                                KnowledgeItem newQuotation = new KnowledgeItem(reference, QuotationType.DirectQuotation);
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
                                    newQuotation.PageRange = pages.Min().ToString();
                                }
                                else
                                {
                                    newQuotation.PageRange = pages.Min().ToString() + "-" + pages.Max().ToString();
                                }

                                newQuotation.TextRtf = textContent.Rtf;

                                reference.Quotations.Add(newQuotation);
                                project.AllKnowledgeItems.Add(newQuotation);

                                EntityLink newEntityLink = new EntityLink(project);
                                newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                                newEntityLink.Source = newQuotation;
                                newEntityLink.Target = newAnnotation;
                                project.EntityLinks.Add(newEntityLink);

                                EntityLink commentQuotationLink = new EntityLink(project);
                                commentQuotationLink.Source = comment;
                                commentQuotationLink.Target = newQuotation;
                                commentQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                                project.EntityLinks.Add(commentQuotationLink);

                                comment.CoreStatement = newQuotation.CoreStatement + " (Comment)";
                                comment.PageRange = newQuotation.PageRange;
                            }

                            annotationsToDelete.Add(annot);
                        }
                    }
                    foreach (Annot annotation in annotationsToDelete)
                    {
                        page.AnnotRemove(annotation);
                    }
                }
            }
            foreach (Annotation annotation in annotations)
            {
                System.Diagnostics.Debug.WriteLine("CITAVI ANNOTATION");
                foreach (Quad quad in annotation.Quads)
                {
                    System.Diagnostics.Debug.WriteLine("QUAD");
                    System.Diagnostics.Debug.WriteLine("MinX: " + quad.MinX + ", MinY: " + quad.MinY + ", MaxX: " + quad.MaxX + ", MaxY: " + quad.MaxY);
                }

            }
        }
    }
}
