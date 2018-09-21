using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using SwissAcademic.Citavi;

namespace QuotationsToolbox
{
    class QuickReferenceTitleCaser
    {
        public static void TitleCaseQuickReference(List<KnowledgeItem> quotations)
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Citavi 6";
            string file = "CitaviAccronyms.txt";
            string stringPath = folder + "\\" + file;

            bool NoAccronymFile = false;

            if (!File.Exists(stringPath)) NoAccronymFile = true;

            string replacements = string.Empty;

            if (!NoAccronymFile) replacements = System.IO.File.ReadAllText(stringPath);

            List<string> accronymList = replacements.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            foreach (KnowledgeItem quotation in quotations)
            {
                if (quotation.QuotationType != QuotationType.QuickReference) continue;
                string text = quotation.CoreStatement;
                text = TitleCase(text);
                quotation.CoreStatement = text;               
            }
        }
        public static string TitleCase(string text)
        {
            string output;
            output = text.ToTitleCaseProperEnglish();
            return output;
        }
    }
}
