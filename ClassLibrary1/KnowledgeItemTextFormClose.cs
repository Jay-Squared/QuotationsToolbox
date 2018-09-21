using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;

namespace QuotationsToolbox
{
    class KnowledgeItemTextFormClose
 :
        CitaviAddOn
    {
        public override AddOnHostingForm HostingForm
        {
            get { return AddOnHostingForm.KnowledgeItemTextForm; }
        }

        protected override void OnHostingFormLoaded(System.Windows.Forms.Form hostingForm)
        {

            KnowledgeItemTextForm knowledgeItemTextForm = (KnowledgeItemTextForm)hostingForm;
            var button = knowledgeItemTextForm.GetCommandbar(KnowledgeItemTextFormCommandbarId.Menu).GetCommandbarMenu(KnowledgeItemTextFormCommandbarMenuId.File).AddCommandbarButton("KnowledgeItemTextFormClose", "Save and close");

            button.Shortcut = System.Windows.Forms.Shortcut.AltDownArrow;

            base.OnHostingFormLoaded(hostingForm);
        }

        protected override void OnBeforePerformingCommand(SwissAcademic.Controls.BeforePerformingCommandEventArgs e)
        {
            switch (e.Key)
            {
                case "KnowledgeItemTextFormClose":
                    {

                        e.Handled = true;

                        if (Program.ActiveProjectShell.ActiveForm.GetType().ToString().Contains("KnowledgeItemTextForm"))
                        {
                            KnowledgeItemTextForm form = Program.ActiveProjectShell.ActiveForm as KnowledgeItemTextForm;
                            form.Save();
                            form.Close();
                        }

                    }
                    break;
            }
            base.OnBeforePerformingCommand(e);
        }
    }
}
