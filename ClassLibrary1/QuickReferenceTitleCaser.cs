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
                text = QuotationTextCleaner.TextCleaner(text);
                quotation.CoreStatement = text;               
            }
        }
        public static string TitleCase(string text)
        {
            string output;

            text = Regex.Replace(text, @"(?<!^)([^IVX])", m => m.ToString().ToLower());
            text = Regex.Replace(text, @"[IVX](?![IVX\s\p{P}])", m => m.ToString().ToLower());

            text = Regex.Replace(text, @"^[a-z](?=[a-z]{2,})", s => (s.Value.ToUpper()));
            text = Regex.Replace(text, @"\s[a-z](?=[a-z]{2,})", s => (s.Value.ToUpper()));

            text = Regex.Replace(text, " And ", " and ");
            text = Regex.Replace(text, " For ", " for ");
            text = Regex.Replace(text, " From ", " from ");
            text = Regex.Replace(text, " The ", " the ");
            text = Regex.Replace(text, " With ", " with ");

            text = Regex.Replace(text, @"[\p{P}]\s[a-z]", s => (s.Value.ToUpper()));


            output = text;
            return output;
        }
    }
}
