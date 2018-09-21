using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Pdf;
using SwissAcademic.Pdf.Analysis;

using pdftron;
using pdftron.Filters;
using pdftron.FDF;
using pdftron.Common;
using pdftron.SDF;
using pdftron.PDF;

namespace QuotationsToolbox
{
    public class PageWidth
    {
        public Location Location { get; set; }
        public int PageIndex { get; set; }
        public double Width { get; set; }

        public PageWidth(Location location, int pageIndex, double width)
        {
            Location = location;
            PageIndex = pageIndex;
            Width = width;
        }
    }
}
