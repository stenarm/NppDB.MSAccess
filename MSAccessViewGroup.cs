using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    internal class MsAccessViewGroup : MsAccessTableGroup
    {
        public MsAccessViewGroup()
        {
            SchemaName = OleDbMetaDataCollectionNames.Views;
            Text = "Views";
        }

        protected override TreeNode CreateTreeNode(DataRow dataRow)
        {
            var viewNode = new MSAccessView
            {
                Text = dataRow["TABLE_NAME"].ToString(),
            };

            return viewNode;
        }
    }
}