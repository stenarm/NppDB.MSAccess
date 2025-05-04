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
                        if (SchemaName == OleDbMetaDataCollectionNames.Tables && row.Table.Columns.Contains("TABLE_TYPE") && row["TABLE_TYPE"] as string == "VIEW") continue;
                        if (tableName != null && (tableName.ToUpper().StartsWith("MSYS") ||
                                                  tableName.ToUpper().StartsWith("USYS")))
                            continue;

                        var childNode = CreateTreeNode(row);

                        if (NodeHasChildrenCheck(cnn, tableName))
                        {
                            childNode.Nodes.Add(new TreeNode(""));
                        }

                        Nodes.Add(childNode);
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

        private static bool NodeHasChildrenCheck(OleDbConnection conn, string tableOrViewName)
        {
            DataTable dt = null;
            try
            {
                dt = conn.GetSchema("Columns", new[] { null, null, tableOrViewName, null });
                return (dt != null && dt.Rows.Count > 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking columns for {tableOrViewName}: {ex.Message}");
                return true;
            }
            finally
            {
                 dt?.Dispose();
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