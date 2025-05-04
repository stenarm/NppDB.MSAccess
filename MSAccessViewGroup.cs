using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    internal class MSAccessViewGroup : MsAccessTableGroup
    {
        public MSAccessViewGroup()
        {
            SchemaName = OleDbMetaDataCollectionNames.Views;
            Text = "Views";
        }

        protected override TreeNode CreateTreeNode(DataRow dataRow)
        {
            return new MSAccessView
            {
                Text = dataRow["table_name"].ToString(),
                Definition = dataRow["view_definition"].ToString(),
            };
        }
    }
}
