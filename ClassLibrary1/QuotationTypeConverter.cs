using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic.Citavi;

namespace QuotationsToolbox
{
    class QuotationTypeConverter
    {
        public static void ConvertDirectQuoteToRedHighlight(List<KnowledgeItem> quotations)
        {
            foreach (KnowledgeItem quotation in quotations)
            {
                if (quotation.QuotationType == QuotationType.QuickReference) continue;
                string quotationText = quotation.Text;
                quotation.QuotationType = QuotationType.QuickReference;
                quotation.Text = string.Empty;
                quotation.CoreStatement = quotationText;
            }
        }
    }
}
