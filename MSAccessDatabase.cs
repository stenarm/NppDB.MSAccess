using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class MsAccessDatabase : TreeNode, IRefreshable, IMenuProvider
    {
        public MsAccessDatabase()
        {
            SelectedImageKey = ImageKey = "Database";
        }

        public void Refresh()
        {
            Nodes.Clear();
            var connect = GetDbConnect();
            if (connect == null) return;

            using (var checkConn = connect.GetConnection())
            {
                try
                {
                    checkConn.Open();

                    var tablesNode = new MsAccessTableGroup();
                    if (SchemaGroupHasChildren(checkConn, "TABLE"))
                    {
                        tablesNode.Nodes.Add(new TreeNode(""));
                    }
                    Nodes.Add(tablesNode);

                    var viewsNode = new MsAccessViewGroup();
                    if (SchemaGroupHasChildren(checkConn, "VIEW"))
                    {
                        viewsNode.Nodes.Add(new TreeNode(""));
                    }
                    Nodes.Add(viewsNode);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error checking children for Access DB: {ex.Message}");
                    if(Nodes.Count == 0)
                    {
                        Nodes.Add(new MsAccessTableGroup());
                        Nodes.Add(new MsAccessViewGroup());
                    }
                }
            }
        }

        private static bool SchemaGroupHasChildren(OleDbConnection conn, string objectType)
        {
            var schemaCollection = "";
            switch (objectType)
            {
                case "TABLE":
                    schemaCollection = OleDbMetaDataCollectionNames.Tables;
                    break;
                case "VIEW":
                    schemaCollection = OleDbMetaDataCollectionNames.Views;
                    break;
                default:
                    return true;
            }

            DataTable dt = null;
            try
            {
                dt = conn.GetSchema(schemaCollection);
                foreach (DataRow row in dt.Rows)
                {
                    if (objectType != "TABLE") return true;
                    var tableName = row["TABLE_NAME"] as string;
                    if (row["TABLE_TYPE"] is string tableType && tableType.Contains("VIEW")) continue;
                    if (tableName != null && (tableName.ToUpper().StartsWith("MSYS") || tableName.ToUpper().StartsWith("USYS"))) continue;

                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking {objectType} count for Access: {ex.Message}");
                return true;
            }
            finally
            {
                 dt?.Dispose();
            }
        }


        public ContextMenuStrip GetMenu()
        {
            var menuList = new ContextMenuStrip { ShowImageMargin = false };
            menuList.Items.Add(new ToolStripButton("Refresh", null, (s, e) => { Refresh(); }));
            return menuList;
        }

        private MsAccessConnect GetDbConnect()
        {
            return Parent as MsAccessConnect;
        }
    }
}