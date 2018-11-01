using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading;
using System.Globalization;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

using System.Reflection;
using SwissAcademic.Citavi.Shell.Controls.Preview;

namespace QuotationsToolbox
{
    class LocationMover
    {
        public static void MoveAttachment()
        {
            if (Program.ProjectShells.Count == 0) return;       //no project open

            Program.ActiveProjectShell.PrimaryMainForm.PreviewControl.ShowNoPreview();

            // User Options

            // Static Variables

            List<Reference> references = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedReferences();

            List<Location> locations = references.SelectMany(r => r.Locations).Where(l => l.Address.LinkedResourceType == LinkedResourceType.AttachmentFile).ToList();

            Project activeProject = Program.ActiveProjectShell.Project;

            string citaviAttachmentsFolderPath = activeProject.Addresses.AttachmentsFolderPath;

            // Dynamic Variables

            string targetFolder = string.Empty;

            var fbd = new FolderBrowserDialog();

            List<Reference> renamingFailed = new List<Reference>();
            int renameCounter = 0;

            // The Magic

            targetFolder = citaviAttachmentsFolderPath;

            fbd.SelectedPath = targetFolder;
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.Cancel) return;
            targetFolder = fbd.SelectedPath;

            foreach (Location location in locations)
            {
                string oldFilePath = location.Address.Resolve().LocalPath;

                FileInfo fileInfo = new FileInfo(oldFilePath);

                string fileExtension = fileInfo.Extension;
                if (fileExtension.Equals("pdf", StringComparison.InvariantCultureIgnoreCase))
                {
                    return;
                }

                string fileName = fileInfo.Name;

                string newAbsoluteFilePath = targetFolder + @"\" + fileName;

                if (newAbsoluteFilePath.Equals(oldFilePath)) continue;

                if (File.Exists(newAbsoluteFilePath))
                {
                    renamingFailed.Add(location.Reference);
                    continue;
                }

                try
                {
                    location.Address.ChangeFilePathAsync(new System.Uri(newAbsoluteFilePath), AttachmentAction.Move);
                    renameCounter++;
                }
                catch (Exception e)
                {
                    MessageBox.Show("An error has occured while renaming " + location.Address + ":\n" + e.Message, "Citavi Macro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    renamingFailed.Add(location.Reference);
                    return;
                }
            }
            string message = "{0} locations have been moved.";
            message = string.Format(message, renameCounter);

            if (renamingFailed.Count == 0)
            {
                MessageBox.Show(message, "Citavi Macro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                DialogResult showFailed = MessageBox.Show(message + "\n Would you like to show a selection of references where the move has failed?", "Citavi Macro", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (showFailed == DialogResult.Yes)
                {
                    var filter = new ReferenceFilter(renamingFailed, "Renaming failed", false);
                    Program.ActiveProjectShell.PrimaryMainForm.ReferenceEditorFilterSet.Filters.ReplaceBy(new List<ReferenceFilter> { filter });
                    return;
                }
            }
            return;
        }
    }
}
