using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Controls.WindowsForms;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

namespace QuotationsToolbox
{
    class AnnotationSimplifier
    {
        public static void SimplifyAnnotations()
        {
            PreviewControl previewControl = PreviewMethods.GetPreviewControl();
            if (previewControl == null) return;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            Document document = pdfViewControl.Document;
            if (document == null) return;

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return;

            Location location = previewControl.ActiveLocation;
            if (location == null) return;

            Reference reference = location.Reference;
            if (reference == null) return;

            List<Annotation> annotations = previewControl.GetPdfViewControl().GetSelectedAnnotations().ToList();

            foreach (Annotation annotation in annotations)
            {
                List<Quad> newQuads = new List<Quad>();
                foreach (int i in annotation.Quads.Select(q => q.PageIndex).Distinct())
                {
                    List<Quad> quadsOnPage = annotation.Quads.Where(q => q.PageIndex == i).ToList();

                    newQuads.AddRange(quadsOnPage.SimpleQuads());

                }
                annotation.Quads = newQuads;
            }
        }
    }
}
