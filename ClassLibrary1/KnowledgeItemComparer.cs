using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Controls.WindowsForms;
using SwissAcademic.Pdf;

namespace QuotationsToolbox
{
    class KnowledgeItemComparer : IComparer<KnowledgeItem>
    {
        List<PageWidth> _store;
        public KnowledgeItemComparer(List <PageWidth> store)
        {
            _store = store;
        }
        public int Compare(KnowledgeItem x, KnowledgeItem y)
        {
             // First, we compare the reference (identified by short title) of each knowledge item

            var xCriterionOne = string.Empty;            
            var yCriterionOne = string.Empty;
                
            if (x.Reference != null) xCriterionOne = x.Reference.ShortTitle;
            if (y.Reference != null) yCriterionOne = y.Reference.ShortTitle;

            if (xCriterionOne.Equals(yCriterionOne, StringComparison.Ordinal) == false) return xCriterionOne.CompareTo(yCriterionOne);

            // Now we start looking at the actual annotations

            if (x.EntityLinks.Any() && y.EntityLinks.Any())
            {
                Annotation xAnnotation = null;
                Annotation yAnnotation = null;

                if (x.EntityLinks.ToList().Where(link => link.Target is Annotation).FirstOrDefault() != null)
                {
                    xAnnotation = x.EntityLinks.ToList().Where(link => link.Target is Annotation).FirstOrDefault().Target as Annotation;
                }
                else
                {
                    if (x.EntityLinks.Any() && x.QuotationType == QuotationType.Comment)
                    {
                        try
                        {
                            var target = x.EntityLinks.ToList().Where(n => n != null && n.Target as KnowledgeItem != null).ToList().FirstOrDefault().Target as KnowledgeItem;
                            if (target != null) xAnnotation = target.EntityLinks.FirstOrDefault(link => link.Target is Annotation).Target as Annotation;
                        }
                        catch { }
                    }
                }

                if (y.EntityLinks.ToList().Where(link => link.Target is Annotation).FirstOrDefault() != null)
                {
                    yAnnotation = y.EntityLinks.ToList().Where(link => link.Target is Annotation).FirstOrDefault().Target as Annotation;
                }
                else
                {
                    if (y.EntityLinks.Any() && y.QuotationType == QuotationType.Comment)
                    {
                        try
                        {
                            var target = y.EntityLinks.ToList().Where(n => n != null && n.Target as KnowledgeItem != null).ToList().FirstOrDefault().Target as KnowledgeItem;
                            if (target != null) yAnnotation = target.EntityLinks.FirstOrDefault(link => link.Target is Annotation).Target as Annotation;
                        }
                        catch { }
                    }
                }

                if (xAnnotation != null && yAnnotation != null)
                {
                    // Second, we compare in which attachment each knowledge item resides

                    var xCriterionTwo = xAnnotation.Location.Address;
                    var yCriterionTwo = yAnnotation.Location.Address;

                    if (xCriterionTwo != yCriterionTwo) return xCriterionTwo.CompareTo(yCriterionTwo);

                    // Third, we compare on which page of the attachment the knowledge item resides

                    int xCriterionThree = 0;
                    int yCriterionThree = 0;

                    foreach (SwissAcademic.Pdf.Analysis.Quad quad in xAnnotation.Quads)
                    {
                        if (quad.PageIndex > xCriterionThree) xCriterionThree = quad.PageIndex;

                    }
                    foreach (SwissAcademic.Pdf.Analysis.Quad quad in yAnnotation.Quads)
                    {
                        if (quad.PageIndex > yCriterionThree) yCriterionThree = quad.PageIndex;

                    }

                    if (xCriterionThree != yCriterionThree) return xCriterionThree.CompareTo(yCriterionThree);

                    // Eight, we compare the page range of each knowledge item

                    try
                    {
                        var xCriterionEight = Decimal.Parse(x.PageRange.StartPage.OriginalString);
                        var yCriterionEight = Decimal.Parse(y.PageRange.StartPage.OriginalString);
                        if (xCriterionEight != yCriterionEight) return xCriterionEight.CompareTo(yCriterionEight);

                        xCriterionEight = Decimal.Parse(x.PageRange.EndPage.OriginalString);
                        yCriterionEight = Decimal.Parse(y.PageRange.EndPage.OriginalString);
                        if (xCriterionEight != yCriterionEight) return xCriterionEight.CompareTo(yCriterionEight);

                    }
                    catch (Exception)
                    {
                        var xCriterionEight = x.PageRange.OriginalString;
                        var yCriterionEight = y.PageRange.OriginalString;

                        if (xCriterionEight != null && yCriterionEight != null && xCriterionEight != yCriterionEight) return xCriterionEight.CompareTo(yCriterionEight);
                    }

                    // Fourth, we compare whether each knowledge item is on the left side or the right side of a two-column text                

                    if (x.Reference.CustomField8 == "2")
                    {
                        var xCriterionFour = 0;
                        var yCriterionFour = 0;

                        // Der Fall kann eintreten dass eine Markierung über 2 Seiten des Dokumentes geht, der ist hier nicht abgedeckt!

                        int pageNumber = xAnnotation.Quads.GetStartPage();
                        double pageWidth = 0.00;

                        if (_store.Count() > 0)
                        {
                            pageWidth = _store.Find(c => (c.Location == xAnnotation.Location) && (c.PageIndex == pageNumber)).Width;
                        }

                        if (xAnnotation.Quads.Any(q => (q.MinX <= 9 * pageWidth / 20))) xCriterionFour = -1;

                        if (yAnnotation.Quads.Any(q => (q.MinX <= 9 * pageWidth / 20))) yCriterionFour = -1;

                        if (xCriterionFour != yCriterionFour) return xCriterionFour.CompareTo(yCriterionFour);
                    }

                    // Fifth, we compare the Y position of each knowledge item

                    var xCriterionFive = -xAnnotation.Quads.FirstOrDefault().Y1;
                    var yCriterionFive = -yAnnotation.Quads.FirstOrDefault().Y1;

                    foreach (SwissAcademic.Pdf.Analysis.Quad quad in xAnnotation.Quads)
                    {
                        if (quad.Y1 > -xCriterionFive) xCriterionFive = -quad.Y1;

                    }
                    foreach (SwissAcademic.Pdf.Analysis.Quad quad in yAnnotation.Quads)
                    {
                        if (quad.Y1 > -yCriterionFive) yCriterionFive = -quad.Y1;

                    }

                    if (xCriterionFive != yCriterionFive) return xCriterionFive.CompareTo(yCriterionFive);

                    // Sixth, we compare the X position of each knowledge item

                    var xCriterionSix = xAnnotation.Quads.FirstOrDefault().X1;
                    var yCriterionSix = yAnnotation.Quads.FirstOrDefault().X1;

                    if (xCriterionSix != yCriterionSix) return xCriterionSix.CompareTo(yCriterionSix);

                    // Seven, we compare what type each knowledge item is

                    var xCriterionSeven = GetRankByQuotationType(x);
                    var yCriterionSeven = GetRankByQuotationType(y);

                    if (xCriterionSeven != yCriterionSeven) return xCriterionSeven.CompareTo(yCriterionSeven);
                }
            }

            // Eight, we compare the page range of each knowledge item

            try
            {
                var xCriterionEight = Decimal.Parse(x.PageRange.StartPage.OriginalString);
                var yCriterionEight = Decimal.Parse(y.PageRange.StartPage.OriginalString);
                if (xCriterionEight != yCriterionEight) return xCriterionEight.CompareTo(yCriterionEight);

                xCriterionEight = Decimal.Parse(x.PageRange.EndPage.OriginalString);
                yCriterionEight = Decimal.Parse(y.PageRange.EndPage.OriginalString);
                if (xCriterionEight != yCriterionEight) return xCriterionEight.CompareTo(yCriterionEight);

            }
            catch (Exception)
            {
                var xCriterionEight = x.PageRange.OriginalString;
                var yCriterionEight = y.PageRange.OriginalString;

                if (xCriterionEight != null && yCriterionEight != null && xCriterionEight != yCriterionEight) return xCriterionEight.CompareTo(yCriterionEight);
            }

            return 0;
        }
        int GetRankByQuotationType(KnowledgeItem knowledgeItem)
        {
            switch (knowledgeItem.QuotationType)
            {
                case QuotationType.QuickReference: return 1;
                case QuotationType.DirectQuotation: return 2;
                case QuotationType.IndirectQuotation: return 3;
                case QuotationType.Comment: return 4;
                default: return 5;
            }
        }
    }
}
