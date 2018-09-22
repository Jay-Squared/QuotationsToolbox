using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic;
using SwissAcademic.Citavi;

namespace QuotationsToolbox
{
    class PageRangeMerger
    {
        public static string PageRangeListToString(List<PageRange> pageRangesList)
        {
            string pageRangeText = string.Empty;
            int j = 0;
            foreach (PageRange pageRange in MergeAdjacent(pageRangesList))
            {
                if (j > 0) pageRangeText += ", ";
                pageRangeText += pageRange.StartPage;
                if (pageRange.EndPage != null && !string.IsNullOrEmpty(pageRange.EndPage.OriginalString) && pageRange.EndPage != pageRange.StartPage) pageRangeText += "-" + pageRange.EndPage;
                j = j + 1;
            }
            return pageRangeText;
        }
        public static List<PageRange> MergeAdjacent(List<PageRange> pageRangesList)
        {
            int i = pageRangesList.Count;

            if (pageRangesList.Count < 1) return pageRangesList;

            List<PageRange> newList = new List<PageRange>();

            PageNumber minPage = new PageNumber();
            PageNumber maxPage = new PageNumber();

            int minPageAsInteger = 0;
            int maxPageAsInteger = 0;

            PageRange first = pageRangesList.FirstOrDefault();
            PageRange last = pageRangesList.Last();

            NumberingType previousPageRangeNumberingType = new NumberingType();
            bool PreviousPageRangeWasNumber = false;

            foreach (PageRange pageRange in pageRangesList)
            {
                PageNumber currentStartPage = pageRange.StartPage;
                PageNumber currentEndPage = pageRange.EndPage;
                if (currentEndPage == null) currentEndPage = currentStartPage;
                if (string.IsNullOrEmpty(currentEndPage.OriginalString)) currentEndPage = currentStartPage;

                int currentStartPageAsInteger = 0;
                int currentEndPageAsInteger = 0;

                bool startPageIsNumber = Int32.TryParse(currentStartPage.OriginalString.Replace(".", ""), out currentStartPageAsInteger);
                bool endPageIsNumber = Int32.TryParse(currentEndPage.OriginalString.Replace(".", ""), out currentEndPageAsInteger);

                bool IsDiscreteRange = false;

                bool SameNumberingType = (previousPageRangeNumberingType == pageRange.NumberingType);

                if (pageRangesList.Count == 1)
                {

                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(pageRange.NumberingType);

                    Double result;
                    if (Double.TryParse(pageRange.StartPage.ToString(), out result) == false)
                    {
                        newList.Add(dummyPageRange.Update(pageRange.OriginalString));
                    }
                    else
                    {
                        newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", currentStartPage.OriginalString, currentEndPage.OriginalString)));
                    }
                    continue;
                }
                else if (pageRange == first)
                {
                    minPage = currentStartPage;
                    maxPage = currentEndPage;
                    minPageAsInteger = currentStartPageAsInteger;
                    maxPageAsInteger = currentEndPageAsInteger;
                }
                else if (!PreviousPageRangeWasNumber || !SameNumberingType || !startPageIsNumber || !endPageIsNumber || pageRange.StartPage.Number == null)
                {
                    IsDiscreteRange = true;
                }
                else if (currentStartPageAsInteger.CompareTo(minPageAsInteger) >= 0 && currentEndPageAsInteger.CompareTo(maxPageAsInteger) <= 0 && SameNumberingType)
                {
                    // In this case, we don't have to do anything because the current page range is within the range defined by minPage & maxPage
                }
                else if (currentStartPageAsInteger.CompareTo(maxPageAsInteger + 1) < 1 && SameNumberingType)
                {
                    maxPage = currentEndPage;
                    maxPageAsInteger = currentEndPageAsInteger;
                }
                else
                {
                    IsDiscreteRange = true;
                }

                if (IsDiscreteRange && pageRange == last)
                {
                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(previousPageRangeNumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                    minPage = currentStartPage;
                    maxPage = currentEndPage;
                    dummyPageRange = dummyPageRange.Update(pageRange.NumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                }
                else if (IsDiscreteRange)
                {
                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(previousPageRangeNumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                    minPage = currentStartPage;
                    maxPage = currentEndPage;
                    minPageAsInteger = currentStartPageAsInteger;
                    maxPageAsInteger = currentEndPageAsInteger;
                }
                else if ((newList.Count == 0 || !PreviousPageRangeWasNumber || !SameNumberingType) && pageRange == last)
                {
                    PageRange dummyPageRange = new PageRange();
                    dummyPageRange = dummyPageRange.Update(pageRange.NumberingType);
                    newList.Add(dummyPageRange.Update(string.Format("{0}-{1}", minPage.OriginalString, maxPage.OriginalString)));
                }
                previousPageRangeNumberingType = pageRange.NumberingType;
                PreviousPageRangeWasNumber = true;
            }
            return newList;
        } // end MergeAdjacent
    }
    class PageRangeComparer : IComparer<PageRange>
    {
        public int Compare(PageRange x, PageRange y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    return 0;
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                if (y == null)
                {
                    return 1;
                }
                else
                {
                    if (x.NumberingType.CompareTo(y.NumberingType) != 0)
                    {
                        return x.NumberingType.CompareTo(y.NumberingType);
                    }
                    else if (x.NumeralSystem.CompareTo(y.NumeralSystem) != 0)
                    {
                        return x.NumeralSystem.CompareTo(y.NumeralSystem);
                    }
                    else
                    {
                        return x.StartPage.CompareTo(y.StartPage);
                    }
                }
            }
        }
    } // end PageRangeComparer
}
