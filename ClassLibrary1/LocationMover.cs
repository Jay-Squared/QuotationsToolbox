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

            Location location = Program.ActiveProjectShell.PrimaryMainForm.GetSelectedElectronicLocations().FirstOrDefault();

            Project activeProject = Program.ActiveProjectShell.Project;

            string citaviAttachmentsFolderPath = activeProject.Addresses.AttachmentsFolderPath;

            Reference reference = location.Reference;

            if (location.LocationType != LocationType.ElectronicAddress) return;
            if (location.Address.LinkedResourceType != LinkedResourceType.AttachmentFile && string.IsNullOrEmpty(location.Address.Resolve().LocalPath)) return;

            // Dynamic Variables

            string targetFolder = string.Empty;

            var fbd = new FolderBrowserDialog();

            // The Magic

            targetFolder = citaviAttachmentsFolderPath;

            fbd.SelectedPath = targetFolder;
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.Cancel) return;
            targetFolder = fbd.SelectedPath;

            string oldFilePath =  location.Address.Resolve().LocalPath;

            FileInfo fileInfo = new FileInfo(oldFilePath);

            string fileExtension = fileInfo.Extension;
            if (fileExtension.Equals("pdf", StringComparison.InvariantCultureIgnoreCase))
            {
                return;
            }

            string fileName = fileInfo.Name;

            string newAbsoluteFilePath = targetFolder + @"\" + fileName;

            if (newAbsoluteFilePath.Equals(oldFilePath)) return;

            if (File.Exists(newAbsoluteFilePath))
            {
                // renamingFailed.Add(reference);
                // continue;
            }

            try
            {
                location.Address.ChangeFilePathAsync(new System.Uri(newAbsoluteFilePath), AttachmentAction.Move);
            }
            catch (Exception e)
            {
                MessageBox.Show("An error has occured while renaming " + location.Address + ":\n" + e.Message, "Citavi Macro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            MessageBox.Show("“" + location.FullName + "” has been moved to “" + newAbsoluteFilePath + "”.", "Citavi Macro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
    }
}
