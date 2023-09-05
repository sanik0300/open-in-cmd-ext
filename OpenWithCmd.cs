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
            menuItem.Click += OnClicked;

            menuStrip.Items.Add(menuItem);
            return menuStrip;
        }

        private void OnClicked(object sender, EventArgs e)
        {
            Process cmd = new Process();

            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.WorkingDirectory = FolderPath;

            cmd.Start();
        }
    }
}
