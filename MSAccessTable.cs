using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Windows.Forms;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class MSAccessTable : TreeNode, IRefreshable, IMenuProvider
    {
        public string Definition { get; set; }
        protected string TypeName { get; set; } = "TABLE";

        public MSAccessTable()
        {
            SelectedImageKey = ImageKey = "Table";
        }

        public virtual void Refresh()
        {
            var conn = (MsAccessConnect)Parent.Parent.Parent;
            using (var cnn = conn.GetConnection())
            {
                TreeView.Enabled = false;
                TreeView.Cursor = Cursors.WaitCursor;
                try
                {
                    cnn.Open();

                    Nodes.Clear();
                    
                    var columns = new List<MSAccessColumnInfo>();

                    var primaryKeyColumnNames = CollectPrimaryKeys(cnn, ref columns);
                    var foreignKeyColumnNames = CollectForeignKeys(cnn, ref columns);
                    var indexedColumnNames = CollectIndices(cnn, ref columns);
                    
                    var columnCount = CollectColumns(cnn, ref columns, primaryKeyColumnNames, foreignKeyColumnNames, indexedColumnNames);
                    if (columnCount == 0) return;
                    
                    var maxLength = columns.Max(c => c.ColumnName.Length);
                    columns.ForEach(c => c.AdjustColumnNameFixedWidth(maxLength));
                    Nodes.AddRange(columns.ToArray<TreeNode>());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Exception");
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
            var connect = GetDBConnect();
            menuList.Items.Add(new ToolStripButton("Refresh", null, (s, e) => { Refresh(); }));
            menuList.Items.Add(new ToolStripSeparator());
            if (connect?.CommandHost == null) return menuList;
                    
            var host = connect.CommandHost;
            menuList.Items.Add(new ToolStripButton("Select top 100 ...", null, (s, e) =>
            {
                host.Execute(NppDBCommandType.NewFile, null);
                var id = host.Execute(NppDBCommandType.GetActivatedBufferID, null);
                var query = "SELECT TOP 100 * FROM " + Text;
                host.Execute(NppDBCommandType.AppendToCurrentView, new object[] { query });
                host.Execute(NppDBCommandType.CreateResultView, new[] { id, connect, connect.CreateSqlExecutor() });
                host.Execute(NppDBCommandType.ExecuteSQL, new[] { id, query });
            }));
            menuList.Items.Add(new ToolStripButton("Drop", null, (s, e) =>
            {
                var id = host.Execute(NppDBCommandType.GetActivatedBufferID, null);
                var query = $"DROP {TypeName} {Text}";
                host.Execute(NppDBCommandType.ExecuteSQL, new[] { id, query });
            }));
            return menuList;
        }

        private MsAccessConnect GetDBConnect()
        {
            var connect = Parent.Parent.Parent as MsAccessConnect;
            return connect;
        }

        private int CollectColumns(OleDbConnection connection, ref List<MSAccessColumnInfo> columns,
            in List<string> primaryKeyColumnNames,
            in List<string> foreignKeyColumnNames,
            in List<string> indexedColumnNames)
        {
            var dt = connection.GetSchema(OleDbMetaDataCollectionNames.Columns, new[] {null, null, Text, null});

            var count = 0;
            foreach (var row in dt.AsEnumerable().OrderBy(r => r["ordinal_position"]))
            {
                var typename = ((OleDbType)int.Parse(row["data_type"].ToString())).ToString().ToUpper();
                var columnName = row["column_name"].ToString();
                var columnType = $"{typename}{(row["character_maximum_length"] is DBNull ? "" : "(" + row["character_maximum_length"] + ")")}";

                var options = 0;

                if (!Convert.ToBoolean(row["is_nullable"].ToString())) options += 1;
                if (indexedColumnNames.Contains(columnName)) options += 10;
                if (primaryKeyColumnNames.Contains(columnName)) options += 100;
                if (foreignKeyColumnNames.Contains(columnName)) options += 1000;
                
                columns.Insert(count++, new MSAccessColumnInfo(columnName, columnType, 0, options));
            }
            return count;
        }
        
        private List<string> CollectPrimaryKeys(OleDbConnection connection, ref List<MSAccessColumnInfo> columns)
        {
            // Restriction columns: TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME
            var dataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys, new object[] { null, null, Text });
            
            var names = new List<string>();
            if (dataTable == null) return names;
            
            foreach (DataRow row in dataTable.Rows)
            {
                var primaryKeyName = row["pk_name"].ToString();
                var primaryKeyType = $"({row["column_name"]})";
                columns.Add(new MSAccessColumnInfo(primaryKeyName, primaryKeyType, 1, 0));
                names.Add(row["column_name"].ToString());
            }
            return names;
        }
        
        private List<string> CollectForeignKeys(OleDbConnection connection, ref List<MSAccessColumnInfo> columns)
        {
            // Restriction columns: PK_TABLE_CATALOG, PK_TABLE_SCHEMA, PK_TABLE_NAME, FK_TABLE_CATALOG, FK_TABLE_SCHEMA, FK_TABLE_NAME
            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, 
                new object[] { null, null, null, null, null, Text });

            var names = new List<string>();
            if (schema == null) return names;
            
            foreach (DataRow row in schema.Rows)
            {
                var foreignKeyName = row["fk_name"].ToString();
                var foreignKeyType = $"({row["fk_column_name"]}) -> {row["pk_table_name"]} ({row["pk_column_name"]})";
                columns.Add(new MSAccessColumnInfo(foreignKeyName, foreignKeyType, 2, 0));
                names.Add(row["fk_column_name"].ToString());
            }
            return names;
        }
        
        private List<string> CollectIndices(OleDbConnection connection, ref List<MSAccessColumnInfo> columns)
        {
            // Restriction columns: TABLE_CATALOG, TABLE_SCHEMA, INDEX_NAME, TYPE, TABLE_NAME
            var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Indexes, 
                new object[] { null, null, null, null, Text });

            var names = new List<string>();
            if (schema == null) return names;

            foreach (DataRow row in schema.Rows) // unique, primary_key, column_name, collation
            {
                var indexName = row["index_name"].ToString();
                var indexType = $"({row["column_name"]})";
                columns.Add(new MSAccessColumnInfo(indexName, indexType, Convert.ToBoolean(row["unique"].ToString()) ? 4 : 3, 0));
                names.Add(row["column_name"].ToString());
            }
            return names;
        }
    }
}
