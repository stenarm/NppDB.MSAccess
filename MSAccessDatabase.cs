using System.Windows.Forms;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class MSAccessDatabase : TreeNode, IRefreshable, IMenuProvider
    {
        public MSAccessDatabase()
        {
            SelectedImageKey = ImageKey = "Database";
            Refresh();
        }
        
        public void Refresh()
        {
            Nodes.Clear();
            Nodes.Add(new MSAccessTableGroup());
            Nodes.Add(new MSAccessViewGroup());
            // add other categories as stored procedures
        }

        public ContextMenuStrip GetMenu()
        {
            var menuList = new ContextMenuStrip { ShowImageMargin = false };
            menuList.Items.Add(new ToolStripButton("Refresh", null, (s, e) => { Refresh(); }));
            return menuList;
        }
    }
}
