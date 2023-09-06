using SharpShell;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenWithCmdExt
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.DirectoryBackground)]
    public class OpenWithCMD : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip menuStrip = new ContextMenuStrip();

            ToolStripMenuItem menuItem = new ToolStripMenuItem() { Text = "Open in Command prompt" };

            ToolStripMenuItem runAsUserItem = new ToolStripMenuItem() { Text = "As Current User" },
                              runAsAdmin = new ToolStripMenuItem() { Text = "As Administrator" };

            runAsUserItem.Click += (s, e) => { RunCMD(false); };
            runAsAdmin.Click += (s, e) => { RunCMD(true); };

            menuItem.DropDownItems.AddRange(new ToolStripItem[] {runAsUserItem, runAsAdmin});

            menuStrip.Items.Add(menuItem);
            return menuStrip;
        }

        private void RunCMD(bool asadmin)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            
            if(asadmin)
            {
                cmd.StartInfo.UseShellExecute = true;
                cmd.StartInfo.Arguments = $"/k cd /d \"{FolderPath}\"";
                cmd.StartInfo.Verb = "runas";
            }
            else {
                cmd.StartInfo.WorkingDirectory = FolderPath;
            }

            cmd.Start();
        }
    }
}
