using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

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
    partial class AnnotationsImporterColorPicker
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
        private void InitializeComponent(QuotationType quotationType, List<ColorPt> existingColorPts, out List<ColorPt> selectedColorPts)
        {
            selectedColorPts = new List<ColorPt>();

            ImportEmptyAnnotationsSelected = false;
            RedrawAnnotationsSelected = true;
            ImportDirectQuotationLinkedWithCommentSelected = true;

            string colorPickerCaption = string.Empty;

            switch (quotationType)
            {
                case QuotationType.Comment:
                    colorPickerCaption = "Select the colors of all comments.";
                    break;
                case QuotationType.DirectQuotation:
                    colorPickerCaption = "Select the colors of all direct quotations.";
                    ImportEmptyAnnotationsSelected = true;
                    break;
                case QuotationType.IndirectQuotation:
                    colorPickerCaption = "Select the colors of all indirect quotations.";
                    break;
                case QuotationType.QuickReference:
                    colorPickerCaption = "Select the colors of all quick references.";
                    ImportEmptyAnnotationsSelected = true;
                    break;
                case QuotationType.Summary:
                    colorPickerCaption = "Select the colors of all summaries.";
                    break;
                default:
                    break;
            }


            int formWidth = 300;
            int formHeight = existingColorPts.Count * 50 + 160;

            if (formWidth < 400) formWidth = 375;

            int buttonWidth = 75;
            int buttonHeight = 25;

            int padding = 10;

            this.BackColor = System.Drawing.Color.White;

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

            System.Drawing.Font font = new System.Drawing.Font("Calibri", 13.0f);

            this.textbox1.Location = new System.Drawing.Point(padding, padding);
            this.textbox1.Text = colorPickerCaption;
            this.textbox1.Font = font;
            this.textbox1.ForeColor = System.Drawing.Color.DarkSlateBlue ;
            this.textbox1.Size = new System.Drawing.Size(formWidth - 20, buttonHeight);
            this.textbox1.BackColor = this.BackColor;
            this.textbox1.BorderStyle = BorderStyle.None;
            //
            // 
            // AnnotationsImporterColorPicker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(formWidth, formHeight);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(textbox1);
            this.CancelButton = this.button2;
            this.Name = "AnnotationsImporterColorPicker";
            this.Text = "Color Picker";
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
                System.Drawing.Color foreColor = System.Drawing.Color.FromArgb(existingColorPtA, existingColorPtR, existingColorPtG, existingColorPtB);
                checkBox.ForeColor = foreColor;
                checkBox.BackColor = System.Drawing.Color.White;
                checkBox.Width = formWidth - 30;
                checkBox.Font = new System.Drawing.Font(checkBox.Font.FontFamily, checkBox.Font.Size, System.Drawing.FontStyle.Bold);
                checkBox.AutoSize = true;
                checkBox.Location = new System.Drawing.Point(padding, padding + buttonHeight + padding + i * padding + i * buttonHeight);

                this.Controls.Add(checkBox);

                checkBox.Click += new EventHandler((sender, e) => ColorCheckboxClick(sender, e, checkBox, existingColorPt, selectedColorPtsInBox, out selectedColorPtsInBox));
                selectedColorPts = selectedColorPtsInBox;

                i++;
            }

            // The Options

            if (quotationType == QuotationType.Comment)
            {

                CheckBox ImportDirectQuotationLinkedWithCommentCheckBox = new CheckBox();
                ImportDirectQuotationLinkedWithCommentCheckBox.Checked = ImportDirectQuotationLinkedWithCommentSelected;
                ImportDirectQuotationLinkedWithCommentCheckBox.Width = formWidth - 30;
                ImportDirectQuotationLinkedWithCommentCheckBox.Text = "Import direct quotation linked with comment.";
                ImportDirectQuotationLinkedWithCommentCheckBox.Location = new System.Drawing.Point(padding, formHeight - (40 + 3 * (buttonHeight + padding)));

                this.Controls.Add(ImportDirectQuotationLinkedWithCommentCheckBox);

                ImportDirectQuotationLinkedWithCommentCheckBox.Click += new EventHandler((sender, e) => YesNoCheckboxClick(sender, e, ImportDirectQuotationLinkedWithCommentCheckBox, ImportDirectQuotationLinkedWithCommentSelected, "ImportDirectQuotationLinkedWithComment"));
            }

            CheckBox redrawAnnotationsCheckBox = new CheckBox();
            redrawAnnotationsCheckBox.Checked = RedrawAnnotationsSelected;
            redrawAnnotationsCheckBox.Width = formWidth - 30;
            redrawAnnotationsCheckBox.Text = "Redraw annotations.";
            redrawAnnotationsCheckBox.Location = new System.Drawing.Point(padding, formHeight - (40 + 2 * (buttonHeight + padding)));

            this.Controls.Add(redrawAnnotationsCheckBox);

            redrawAnnotationsCheckBox.Click += new EventHandler((sender, e) => YesNoCheckboxClick(sender, e, redrawAnnotationsCheckBox, RedrawAnnotationsSelected, "RedrawAnnotations"));

            CheckBox emptyAnnotationsCheckBox = new CheckBox();
            emptyAnnotationsCheckBox.Checked = ImportEmptyAnnotationsSelected;
            emptyAnnotationsCheckBox.Width = formWidth - 30;
            emptyAnnotationsCheckBox.Text = "Import empty annotations.";
            emptyAnnotationsCheckBox.Location = new System.Drawing.Point(padding, formHeight - (40 + buttonHeight + padding));

            this.Controls.Add(emptyAnnotationsCheckBox);

            emptyAnnotationsCheckBox.Click += new EventHandler((sender, e) => YesNoCheckboxClick(sender, e, emptyAnnotationsCheckBox, ImportEmptyAnnotationsSelected, "ImportEmptyAnnotations"));

            // The Buttons

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

        private void ColorCheckboxClick(object sender, System.EventArgs e, CheckBox checkBox, ColorPt selectedColorPt, List<ColorPt> selectedColorPts, out List<ColorPt> selectedColorPtsAfter)
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

        private void YesNoCheckboxClick (object sender, System.EventArgs e, CheckBox checkBox, bool BoolIn, string optionName)
        {
            bool Result = BoolIn;
            if (checkBox.Checked)
            {
                Result = true;
            }
            else
            {
                Result = false;
            }
            switch (optionName)
            {
                case "ImportDirectQuotationLinkedWithComment":
                    ImportDirectQuotationLinkedWithCommentSelected = Result;
                    break;
                case "ImportEmptyAnnotations":
                    ImportEmptyAnnotationsSelected = Result;
                    break;
                case "RedrawAnnotations":
                    RedrawAnnotationsSelected = Result;
                    break;
                default:
                    break;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            foreach (Control control in this.Controls)
            {
                if (control is CheckBox)
                {
                    CheckBox checkBox = control as CheckBox;
                    System.Drawing.SolidBrush myBrush = new System.Drawing.SolidBrush(checkBox.ForeColor);
                    System.Drawing.Graphics formGraphics;
                    formGraphics = this.CreateGraphics();
                    formGraphics.FillRectangle(myBrush, new System.Drawing.Rectangle(checkBox.Left + 20 , checkBox.Top, 60, checkBox.Height));
                    myBrush.Dispose();
                    formGraphics.Dispose();
                }

            }
        }

        public static bool RedrawAnnotationsSelected
        {
            get;
            private set;
        }

        public static bool ImportEmptyAnnotationsSelected
        {
            get;
            private set;
        }

        public static bool ImportDirectQuotationLinkedWithCommentSelected
        {
            get;
            private set;
        }

    }
}