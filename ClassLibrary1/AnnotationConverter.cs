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
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using pdftron.PDF;

namespace QuotationsToolbox
{
    class AnnotationConverter
    {
        public static void ConvertAnnotations(Reference reference)
        {
            var project = reference.Project;

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            Document document = previewControl.GetDocument();
            if (document == null) return;

            Location location = document.GetPDFLocationOfDocument();
            if (location == null) return;

            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return;

            PdfViewControl pdfViewControl = propertyInfo.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl) as PdfViewControl;

            int startPageInt = 1;

            if (reference.PageRange.StartPage.Number != null) startPageInt = reference.PageRange.StartPage.Number.Value;

            if (reference == null) return;

            if (document != null)
            {
                List<Annotation> annotations = location.Annotations.Where(a => a.Visible == true && a.EntityLinks.Where(e => e.Indication == EntityLink.PdfKnowledgeItemIndication).Count() == 0).ToList();

                foreach (Annotation annotation in annotations)
                {
                    pdfViewControl.GoToAnnotation(annotation);

                    List<Quad> quads = annotation.Quads.ToList();
                    Content content = pdfViewControl.GetSelectedContentFromType(pdfViewControl.GetSelectedContentType(), -1, false, true);
                    SwissAcademic.Citavi.Controls.Wpf.TextContent textContent = content as TextContent;

                    KnowledgeItem newQuotation = new KnowledgeItem(reference, QuotationType.DirectQuotation);
                    List<int> pages = new List<int>();
                    List<Quad> newQuads = annotation.Quads.ToList();

                    foreach (Quad quad in quads)
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
                }
                
            }
        }
    }
}
