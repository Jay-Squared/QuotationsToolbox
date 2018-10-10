using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using pdftron;
using pdftron.Common;
using pdftron.Filters;
using pdftron.FDF;
using pdftron.PDF;
using pdftron.PDF.Annots;
using pdftron.SDF;

namespace QuotationsToolbox
{
    public partial class CommentAnnotationsColorPicker : Form
    {
        public CommentAnnotationsColorPicker(List<ColorPt> existingColorPts, out List<ColorPt> selectedColorPts)
        {
            InitializeComponent(existingColorPts, out selectedColorPts);
        }

        private void CommentAnnotationsColorPicker_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
