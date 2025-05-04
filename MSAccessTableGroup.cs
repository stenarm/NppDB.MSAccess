using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class MsAccessTableGroup : TreeNode, IRefreshable, IMenuProvider
    {
        public MsAccessTableGroup()
        {
            SchemaName = OleDbMetaDataCollectionNames.Tables;
            Text = "Tables";
            SelectedImageKey = ImageKey = "Group";
        }

        protected string SchemaName { get;  set; }

        protected virtual TreeNode CreateTreeNode(DataRow dataRow)
        {
            var tableNode = new MsAccessTable
            {
                Text = dataRow["table_name"].ToString()
            };

            tableNode.Nodes.Add(new TreeNode(""));

            return tableNode;
        }

        public void Refresh()
        {
            var conn = (MsAccessConnect)Parent.Parent;
            using (var cnn = conn.GetConnection())
            {
                TreeView.Cursor = Cursors.WaitCursor;
                TreeView.Enabled = false;
                try
                {
                    cnn.Open();
                    var dt = cnn.GetSchema(SchemaName);
                    Nodes.Clear();
                    foreach (DataRow row in dt.Rows)
                    {
                        var tableName = row["table_name"] as string;
                        if (SchemaName == OleDbMetaDataCollectionNames.Tables && row["table_type"] as string == "VIEW") continue;
                        if (tableName != null && (tableName.ToUpper().StartsWith("MSYS") ||
                                                  tableName.ToUpper().StartsWith("USYS")))
                            continue;
                        Nodes.Add(CreateTreeNode(row));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, @"Exception");
                }
                finally
                {
                    TreeView.Enabled = true;
                    TreeView.Cursor = null;
                }
            }
        }

        public virtual ContextMenuStrip GetMenu()
        {
            var menuList = new ContextMenuStrip { ShowImageMargin = false };
            menuList.Items.Add(new ToolStripButton("Refresh", null, (s, e) => { Refresh(); }));

            return menuList;
        }
    }
}
