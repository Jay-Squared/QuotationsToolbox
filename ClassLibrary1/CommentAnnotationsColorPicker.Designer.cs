using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

using pdftron;
using pdftron.Common;
using pdftron.Filters;
using pdftron.FDF;
using pdftron.PDF;
using pdftron.PDF.Annots;
using pdftron.SDF;

namespace QuotationsToolbox
{
    partial class CommentAnnotationsColorPicker
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent(List<ColorPt> existingColorPts, out List<ColorPt> selectedColorPts)
        {
            selectedColorPts = new List<ColorPt>();

            int formWidth = System.Convert.ToInt32(Math.Ceiling(new decimal(existingColorPts.Count) / 3) * 300);
            int formHeight = existingColorPts.Count * 50 + 160;

            int buttonWidth = 75;
            int buttonHeight = 25;

            int padding = 10;

            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textbox1 = new TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(formWidth / 5, formHeight - 40);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(buttonWidth, buttonHeight);
            this.button1.TabIndex = 0;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(formWidth - buttonWidth - formWidth /5, formHeight - 40);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(buttonWidth, buttonHeight);
            this.button2.TabIndex = 1;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            //
            // textbox 1
            this.textbox1.Location = new System.Drawing.Point(padding, padding);
            this.textbox1.Text = "Select the annotation colors to be imported as comments:";
            this.textbox1.Size = new System.Drawing.Size(formWidth - 20, buttonHeight);
            //
            // 
            // CommentAnnotationsColorPicker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(formWidth, formHeight);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(textbox1);
            this.CancelButton = this.button2;
            this.Name = "CommentAnnotationsColorPicker";
            this.Text = "CommentAnnotationsColorPicker";
            this.Load += new System.EventHandler(this.CommentAnnotationsColorPicker_Load);
            this.ResumeLayout(false);

            int i = 0;

            List<CheckBox> checkBoxes = new List<CheckBox>();

            foreach (ColorPt existingColorPt in existingColorPts)
            {
                List<ColorPt> selectedColorPtsInBox = selectedColorPts;

                int existingColorPtR = System.Convert.ToInt32(existingColorPt.Get(0) * 255);
                int existingColorPtG = System.Convert.ToInt32(existingColorPt.Get(1) * 255);
                int existingColorPtB = System.Convert.ToInt32(existingColorPt.Get(2) * 255);
                int existingColorPtA = System.Convert.ToInt32((1 - existingColorPt.Get(3)) * 255);

                checkBoxes.Add(new CheckBox());

                CheckBox checkBox = checkBoxes[i];

                checkBox.Tag = i.ToString();
                checkBox.Text = existingColorPtR.ToString() + ", " + existingColorPtG.ToString() + ", " + existingColorPtB.ToString() + ", " + existingColorPtA.ToString();
                System.Drawing.Color foreColor = System.Drawing.Color.FromArgb(existingColorPtA, existingColorPtR, existingColorPtG, existingColorPtB);
                checkBox.ForeColor = foreColor;
                checkBox.Font = new System.Drawing.Font(checkBox.Font.FontFamily, checkBox.Font.Size, System.Drawing.FontStyle.Bold);
                checkBox.AutoSize = true;
                checkBox.Location = new System.Drawing.Point(padding, padding + buttonHeight + padding + i * padding + i * buttonHeight);

                checkBox.Checked = foreColor == (System.Drawing.Color)SwissAcademic.Drawing.KnownColors.AnnotationComment100;

                if (foreColor == (System.Drawing.Color)SwissAcademic.Drawing.KnownColors.AnnotationComment100) selectedColorPtsInBox.Add(existingColorPt);

                this.Controls.Add(checkBox);

                checkBox.Click += new EventHandler((sender, e) => checkBoxClick(sender, e, checkBox, existingColorPt, selectedColorPtsInBox, out selectedColorPtsInBox));
                selectedColorPts = selectedColorPtsInBox;

                i++;
            }

            button1.Click += new EventHandler(cmdOK_Click);
            button2.Click += new EventHandler(cmdCancel_Click);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textbox1;


        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void checkBoxClick(object sender, System.EventArgs e, CheckBox checkBox, ColorPt selectedColorPt, List<ColorPt> selectedColorPts, out List<ColorPt> selectedColorPtsAfter)
        {
            selectedColorPtsAfter = selectedColorPts;
            if (checkBox.Checked)
            {
                selectedColorPtsAfter.Add(selectedColorPt);
            }
            else
            {
                if (selectedColorPtsAfter.Where(s => s == selectedColorPt).ToList().Count > 0) selectedColorPtsAfter.Remove(selectedColorPt);
            }
        }
    }
}