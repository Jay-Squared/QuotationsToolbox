using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;

using System.Collections.Generic;
using System.Linq;

using System.Windows.Forms;

namespace QuotationsToolbox
{
    public static class PreviewMethods
    {
        public static PreviewControl GetPreviewControl()
        {
            MainForm previewMainForm = null;
            List<Form> forms = new List<Form>();
            PreviewControl previewControl = null;

            foreach (Form form in Application.OpenForms)
            {
                if (!form.Visible) continue;
                MainForm mainForm = form as MainForm;
                if (mainForm == null) continue;
                if (!mainForm.IsPreviewFullScreenForm) continue;
                previewMainForm = mainForm;
                break;
            }

            if (previewMainForm == null) previewMainForm = Program.ActiveProjectShell.PrimaryMainForm;

            previewControl = previewMainForm.PreviewControl;

            return previewControl;
        }

        public static PdfViewControl GetPdfViewControl(this PreviewControl previewControl)
        {
            return previewControl?
                   .GetType()?
                   .GetField
                    (
                       "_pdfViewControl",
                       System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic
                    )?
                   .GetValue(previewControl) as PdfViewControl;
        }

        public static IEnumerable<Annotation> GetSelectedAnnotations(this PdfViewControl pdfViewControl)
        {
            return pdfViewControl?
                   .Tool
                   .GetSelectedHighlights()?
                   .Select(adornmentCanvas => adornmentCanvas.Annotation)
                   .Where(annotation => annotation != null)
                   .Distinct()
                   .ToList();
        }

        public static IEnumerable<ICitaviEntity> GetSelectedCitaviEntities(this PdfViewControl pdfViewControl)
        {
            return pdfViewControl?
                   .Tool
                   .GetSelectedHighlights()?
                   .Where(adornmentCanvas => adornmentCanvas.CitaviEntity != null)
                   .Select(adornmentCanvas => adornmentCanvas.CitaviEntity)
                   .Where(entity => entity != null)
                   .Distinct()
                   .ToList();
        }

        static List<AdornmentCanvas> GetSelectedHighlights(this Tool tool)
        {
            return tool?
                  .GetType()?
                  .GetField
                   (
                       "SelectedAdornmentContainers",
                       System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic
                   )?
                  .GetValue(tool) as List<AdornmentCanvas>;
        }
    }
}
