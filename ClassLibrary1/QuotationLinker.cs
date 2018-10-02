using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;

namespace QuotationsToolbox
{
    class QuotationLinker
    {
        public static void LinkQuotations(List<KnowledgeItem> quotations)
        {
            Control quotationSmartRepeater = Program.ActiveProjectShell.PrimaryMainForm.Controls.Find("quotationSmartRepeater", true).FirstOrDefault();
            SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater quotationSmartRepeaterAsQuotationSmartRepeater = quotationSmartRepeater as SwissAcademic.Citavi.Shell.Controls.SmartRepeaters.QuotationSmartRepeater;

            List<KnowledgeItem> quotesOrSummaries = quotations.Where(q => q.QuotationType == QuotationType.DirectQuotation || q.QuotationType == QuotationType.IndirectQuotation || q.QuotationType == QuotationType.Summary).ToList();
            
            List<KnowledgeItem> comments = quotations.Where(q => q.QuotationType == QuotationType.Comment).ToList();

            if (quotesOrSummaries.Count != 1 || comments.Count != 1)
            {
                MessageBox.Show("Please select exactly one direct or indirect quote or summary and exactly one comment.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Project project = Program.ActiveProjectShell.Project;
            if (project == null) return;

            KnowledgeItem comment = comments.FirstOrDefault();
            KnowledgeItem quoteOrSummary = quotesOrSummaries.FirstOrDefault();

            if (comment.EntityLinks.Where(e => e.Indication == EntityLink.CommentOnQuotationIndication).Count() > 0)
            {
                DialogResult dialogResult = MessageBox.Show("The selected comment is already linked to a knowledge item. Do you want to reset the comment's core statement?", "Comment Already Linked", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    comment.CoreStatement = ((KnowledgeItem)comment.EntityLinks.Where(e => e.Indication == EntityLink.CommentOnQuotationIndication).FirstOrDefault().Target).CoreStatement + " (Comment)";
                }
                else if (dialogResult == DialogResult.No)
                {
                    quotationSmartRepeaterAsQuotationSmartRepeater.SelectAndActivate((KnowledgeItem)comment.EntityLinks.Where(e => e.Indication == EntityLink.CommentOnQuotationIndication).FirstOrDefault().Target, true);
                }
                return;
            }

            EntityLink commentDirectQuoteLink = new EntityLink(project);
            commentDirectQuoteLink.Source = comment;
            commentDirectQuoteLink.Target = quoteOrSummary;
            commentDirectQuoteLink.Indication = EntityLink.CommentOnQuotationIndication;
            project.EntityLinks.Add(commentDirectQuoteLink);

            comment.CoreStatement = quoteOrSummary.CoreStatement + " (Comment)";

            return;
        }

    }
}
