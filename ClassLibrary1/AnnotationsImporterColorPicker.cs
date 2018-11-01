using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;

using pdftron;
using pdftron.Common;
using pdftron.Filters;
using pdftron.FDF;
using pdftron.PDF;
using pdftron.PDF.Annots;
using pdftron.SDF;

namespace QuotationsToolbox
{
    public partial class AnnotationsImporterColorPicker : Form
    {
        public AnnotationsImporterColorPicker(QuotationType quotationType, List<ColorPt> existingColorPts, out List<ColorPt> selectedColorPts)
        {
            InitializeComponent(quotationType, existingColorPts, out selectedColorPts);
        }
    }
}
