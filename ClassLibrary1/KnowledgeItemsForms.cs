using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;


using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
namespace QuotationsToolbox
{
    class KnowledgeItemsForms
    {
        public static string NewPageRangeForm(string title, out string output)
        {
            Form form = new Form();
            form.Text = title;

            form.HelpButton = form.MinimizeBox = form.MaximizeBox = false;
            form.ShowIcon = form.ShowInTaskbar = false;
            form.TopMost = true;

            form.Height = 100;
            form.Width = 300;
            form.MinimumSize = new Size(form.Width, form.Height);

            int margin = 5;
            Size size = form.ClientSize;

            TextBox textbox = new TextBox();
            textbox.TextAlign = HorizontalAlignment.Right;
            textbox.Height = 20;
            textbox.Width = size.Width - 2 * margin;
            textbox.Location = new Point(margin, margin);
            textbox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            form.Controls.Add(textbox);

            Button ok = new Button();
            ok.Text = "Ok";
            ok.Click += new EventHandler(OK_Click);
            ok.Height = 23;
            ok.Width = 75;
            ok.Location = new Point(size.Width / 2 - ok.Width / 2, size.Height / 2);
            ok.Anchor = AnchorStyles.Bottom;
            form.Controls.Add(ok);
            form.AcceptButton = ok;

            form.ShowDialog();

            output = textbox.Text;

            return output;
        }
        private static void OK_Click(object sender, EventArgs e)
        {
            Form form = (sender as Control).Parent as Form;
            form.DialogResult = DialogResult.OK;
            form.Close();
        }
    }
}
