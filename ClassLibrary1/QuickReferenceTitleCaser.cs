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

            // Anything but IVX not preceded by the beginning of the line
            text = Regex.Replace(text, @"(?<!^)([^IVX])", m => m.ToString().ToLower());
            // IVX not preceded by Whitespace or Punctuation
            text = Regex.Replace(text, @"(?<!\b|[IVX])[IVX]", m => m.ToString().ToLower());
            // IVX not followed by IVX or white space or punctuation
            text = Regex.Replace(text, @"[IVX](?!\b|[IVX])", m => m.ToString().ToLower());



            // (1) Word Boundary, (2) One Letter, (Positive Lookahead) At Least Two More Letters
            text = Regex.Replace(text, @"(\b)([\p{L}])(?=[\p{L}]{2,})", s => (s.Value.ToUpper()));

            text = Regex.Replace(text, @"(\b)([IVX])([ivx]+)(\b)", s => (s.Value.ToLower()));

            // (1) Word Boundary, (2) One Letter, (3) Period
            text = Regex.Replace(text, @"(\b)([\p{L}])(\.)", s => (s.Value.ToUpper()));

            text = Regex.Replace(text, " And ", " and ");
            text = Regex.Replace(text, " For ", " for ");
            text = Regex.Replace(text, " From ", " from ");
            text = Regex.Replace(text, " The ", " the ");
            text = Regex.Replace(text, " With ", " with ");

            // (1) Punctuation, (2) White Space, (3) Letter
            text = Regex.Replace(text, @"([\.\-\)\:\?\!\-–])([\s]*)([\p{L}])", s => (s.Value.ToUpper()));

            // (1) Beginning of Line, (2) Optional White Space, Digit, or Punctuation, (3) Letter, (Negative Lookahead) Unless Followed by IVX or Punctuation
            text = Regex.Replace(text, @"(^)([\s0-9\p{P}]*)([\p{L}])(?![ivx\p{P}])", s => (s.Value.ToUpper()));

            // (1) Beginning of Line, (2) Upper-Case Letter, (3) Mandatory Whitespace or Punctuation, (4) Letter
            text = Regex.Replace(text, @"(^)([\p{Lu}]+)([\s\p{P}]+)([\p{L}])", s => (s.Value.ToUpper()));

            output = text;
            return output;
        }
    }
}
