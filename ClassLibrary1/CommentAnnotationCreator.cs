﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Citavi.Controls.Wpf;


namespace QuotationsToolbox
{
    class CommentAnnotationCreator
    {
        public static void CreateCommentAnnotation(List<KnowledgeItem> quotations)
        {
            PreviewControl previewControl = Program.ActiveProjectShell.PrimaryMainForm.PreviewControl;

            var type = previewControl.GetType();
            var propertyInfos = type.GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var propertyInfo = propertyInfos.FirstOrDefault(prop => prop.Name.Equals("PdfViewControl", StringComparison.OrdinalIgnoreCase));

            if (propertyInfo == null) return;

            PdfViewControl pdfViewControl = propertyInfo.GetValue(Program.ActiveProjectShell.PrimaryMainForm.PreviewControl) as PdfViewControl;

            Annotation lastAnnotation = null;

            foreach (KnowledgeItem quotation in quotations)
            {
                if (quotation.EntityLinks.Any() && quotation.EntityLinks.Where(link => link.Target is KnowledgeItem).FirstOrDefault() != null)
                {
                    Reference reference = quotation.Reference;
                    if (reference == null) return;

                    Project project = reference.Project;
                    if (project == null) return;

                    KnowledgeItem mainQuotation = quotation.EntityLinks.ToList().Where(n => n != null && n.Indication == EntityLink.CommentOnQuotationIndication && n.Target as KnowledgeItem != null).ToList().FirstOrDefault().Target as KnowledgeItem;

                    Annotation mainQuotationAnnotation = mainQuotation.EntityLinks.Where(link => link.Target is Annotation && link.Indication == EntityLink.PdfKnowledgeItemIndication).FirstOrDefault().Target as Annotation;
                    if (mainQuotationAnnotation == null) return;

                    Location location = mainQuotationAnnotation.Location;
                    if (location == null) return;

                    Annotation newAnnotation = new Annotation(location);

                    newAnnotation.OriginalColor = System.Drawing.Color.FromArgb(255, 255, 255, 0);
                    newAnnotation.Quads = mainQuotationAnnotation.Quads;
                    newAnnotation.Visible = false;
                    location.Annotations.Add(newAnnotation);


                    EntityLink newEntityLink = new EntityLink(project);
                    newEntityLink.Source = quotation;
                    newEntityLink.Target = newAnnotation;
                    newEntityLink.Indication = EntityLink.PdfKnowledgeItemIndication;
                    project.EntityLinks.Add(newEntityLink);

                    lastAnnotation = newAnnotation;
                }
                pdfViewControl.GoToAnnotation(lastAnnotation);
            }
        }
    }
}
