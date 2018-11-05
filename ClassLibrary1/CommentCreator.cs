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
    class CommentCreator
    {
        public static void CreateCommentOnQuotation(List<KnowledgeItem> quotations)
        {
            Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
            SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;
            if (previewControl == null) return;

            PdfViewControl pdfViewControl = previewControl.GetPdfViewControl();
            if (pdfViewControl == null) return;

            Annotation lastAnnotation = null;
            KnowledgeItem lastComment = null;

            foreach (KnowledgeItem quotation in quotations)
            {
                Reference reference = quotation.Reference;
                if (reference == null) return;

                Project project = reference.Project;
                if (project == null) return;

                Annotation mainQuotationAnnotation = quotation.EntityLinks.Where(link => link.Target is Annotation).FirstOrDefault().Target as Annotation;
                if (mainQuotationAnnotation == null) return;

                Location location = mainQuotationAnnotation.Location;
                if (location == null) return;

                KnowledgeItem comment = new KnowledgeItem(reference, QuotationType.Comment);
                comment.PageRange = quotation.PageRange;
                comment.PageRange.Update(quotation.PageRange.NumberingType);
                comment.PageRange.Update(quotation.PageRange.NumeralSystem);
                comment.CoreStatement = quotation.CoreStatement + " (Comment)";
                comment.CoreStatementUpdateType = UpdateType.Manual;
                reference.Quotations.Add(comment);

                EntityLink commentQuotationLink = new EntityLink(project);
                commentQuotationLink.Source = comment;
                commentQuotationLink.Target = quotation;
                commentQuotationLink.Indication = EntityLink.CommentOnQuotationIndication;
                project.EntityLinks.Add(commentQuotationLink);

                Annotation newAnnotation = new Annotation(location);

                newAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
                newAnnotation.Quads = mainQuotationAnnotation.Quads;
                newAnnotation.Visible = false;

                location.Annotations.Add(newAnnotation);

                EntityLink commentAnnotationLink = new EntityLink(project);
                commentAnnotationLink.Source = comment;
                commentAnnotationLink.Target = newAnnotation;
                commentAnnotationLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                project.EntityLinks.Add(commentAnnotationLink);

                lastComment = comment;
                lastAnnotation = newAnnotation;
            }
            quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate(lastComment, true);
            pdfViewControl.GoToAnnotation(lastAnnotation);


            Program.ActiveProjectShell.ShowKnowledgeItemFormForExistingItem(Program.ActiveProjectShell.PrimaryMainForm, lastComment);
        }
    }
}
