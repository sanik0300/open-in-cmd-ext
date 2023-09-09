using SharpShell;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenWithCmdExt
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.DirectoryBackground)]
    public class OpenWithCMD : SharpContextMenu
    {
        static string pathToCmdExe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "cmd.exe"),
                      admin_icon_res_name = Assembly.GetExecutingAssembly().GetManifestResourceNames().Where(x => x.EndsWith(".png")).First(); 

        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip menuStrip = new ContextMenuStrip();

            ToolStripMenuItem menuItem = new ToolStripMenuItem() { Text = resources.StringTable.open_kw };

            Icon cmdIcon = Icon.ExtractAssociatedIcon(pathToCmdExe);
            menuItem.Image = cmdIcon.ToBitmap();
            menuItem.ImageAlign = ContentAlignment.MiddleLeft;
            menuItem.ImageScaling = ToolStripItemImageScaling.SizeToFit;

            ToolStripMenuItem runAsUserItem = new ToolStripMenuItem() { Text = resources.StringTable.as_user },
                              runAsAdmin = new ToolStripMenuItem() { Text = resources.StringTable.as_admin };

            using(Stream str = Assembly.GetExecutingAssembly().GetManifestResourceStream(admin_icon_res_name))
            {
                runAsAdmin.Image = Image.FromStream(str);
            }
            runAsAdmin.ImageAlign = ContentAlignment.MiddleRight;
            runAsAdmin.ImageScaling = ToolStripItemImageScaling.SizeToFit;

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
