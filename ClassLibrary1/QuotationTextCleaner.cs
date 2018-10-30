using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SwissAcademic.Citavi.Controls.Wpf;

using SwissAcademic.Citavi;

namespace QuotationsToolbox
{
    class QuotationTextCleaner
    {
        public static void CleanQuotationsText(List<KnowledgeItem> quotations)
        {
            foreach (KnowledgeItem quotation in quotations)
            {
                string text = string.Empty;

                if (quotation.QuotationType == QuotationType.QuickReference)
                {
                    text = quotation.CoreStatement;
                    quotation.CoreStatement = TextCleaner(text);
                }
                else
                {
                    text = quotation.Text;
                    quotation.Text = TextCleaner(text);
                }
            }
        }
        public static string TextCleaner(string text)
        {
            string output = string.Empty;

            string folder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Citavi 6";
            string file = "CitaviReplacements.txt";
            string stringPath = folder + "\\" + file;
            if (!File.Exists(stringPath)) return null;

            string replacements = System.IO.File.ReadAllText(stringPath);
            List<string> replacementsList = replacements.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            foreach (String replacement in replacementsList)
            {
                if (replacement.StartsWith("//")) continue;
                List<string> replacementList = replacement.Split(new[] { ";;" }, StringSplitOptions.None).ToList();
                text = Regex.Replace(text, replacementList.FirstOrDefault(), replacementList.LastOrDefault());
            }

            output = text;

            return output;
        }
    }
}
