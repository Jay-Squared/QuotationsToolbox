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
    class QuotationTextCleaner
    {
        public static void CleanQuotationsText(List<KnowledgeItem> quotations)
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Citavi 6";
            string file = "CitaviReplacements.txt";
            string stringPath = folder + "\\" + file;
            if (!File.Exists(stringPath)) return;

            string replacements = System.IO.File.ReadAllText(stringPath);
            List<string> replacementsList = replacements.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            foreach (KnowledgeItem quotation in quotations)
            {
                string text = quotation.Text;

                foreach (String replacement in replacementsList)
                {
                    if (replacement.StartsWith("//")) continue;
                    List<string> replacementList = replacement.Split(new[] { ";;" }, StringSplitOptions.None).ToList();
                    text = Regex.Replace(text, replacementList.FirstOrDefault(), replacementList.LastOrDefault());
                }
                quotation.Text = text;
            }
        }
    }
}
