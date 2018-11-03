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

namespace QuotationsToolbox
{
    class PageRangeManualAssigner
    {
        public static void AssignPageRangeManuallyAfterShowingAnnotation()
        {
            KnowledgeItem quotation = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedQuotations().FirstOrDefault();
            if (quotation.EntityLinks.FirstOrDefault() == null) return;

            Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
            SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

            Reference reference = quotation.Reference;
            if (reference == null) return;

            List<KnowledgeItem> quotations = reference.Quotations.ToList();

            int index = quotations.FindIndex(q => q == quotation);

            Annotation annotation = quotation.EntityLinks.FirstOrDefault().Target as Annotation;
            if (annotation == null) return;

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            pdfViewControl.GoToAnnotation(annotation);

            KnowledgeItemsForms.NewPageRangeForm("Please enter the new page range:", out string data);

            if (!String.IsNullOrEmpty(data)) quotation.PageRange = quotation.PageRange.Update(data);

            if (quotations[index - 1] == null) return;

            Program.ActiveProjectShell.PrimaryMainForm.ActiveControl = quotationSmartRepeater;
            quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate(quotations[index - 1]);

            annotation = quotations[index - 1].EntityLinks.FirstOrDefault().Target as Annotation;
            if (annotation == null) return;

            pdfViewControl.GoToAnnotation(annotation);
        }
    }
}
