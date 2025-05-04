using System.Windows.Forms;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class MsAccessDatabase : TreeNode, IRefreshable, IMenuProvider
    {
        public MsAccessDatabase()
        {
            SelectedImageKey = ImageKey = "Database";
            Refresh();
        }
        
        public void Refresh()
        {
            Nodes.Clear();

            var tablesNode = new MsAccessTableGroup();
            tablesNode.Nodes.Add(new TreeNode(""));
            Nodes.Add(tablesNode);

            var viewsNode = new MSAccessViewGroup();
            viewsNode.Nodes.Add(new TreeNode(""));
            Nodes.Add(viewsNode);
        }
        public ContextMenuStrip GetMenu()
        {
            var menuList = new ContextMenuStrip { ShowImageMargin = false };
            menuList.Items.Add(new ToolStripButton("Refresh", null, (s, e) => { Refresh(); }));
            return menuList;
        }
    }
}
