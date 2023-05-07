using System.Drawing;
using System.Windows.Forms;

namespace NppDB.MSAccess
{
    public class MSAccessColumnInfo : TreeNode
    {
        public string ColumnName { get; }
        public string ColumnType { get; }
        
        public MSAccessColumnInfo(string columnName, string columnType, int type, int options)
        {
            NodeFont = new Font("Consolas", 8F, FontStyle.Regular);
            ColumnName = columnName;
            ColumnType = columnType;
            AdjustColumnNameFixedWidth(0);
            
            switch (type)
            {
                case 1:
                    SelectedImageKey = ImageKey = "Primary_Key";
                    break;
                case 2:
                    SelectedImageKey = ImageKey = "Foreign_Key";
                    break;
                case 3:
                    SelectedImageKey = ImageKey = "Index";
                    break;
                case 4:
                    SelectedImageKey = ImageKey = "Unique_Index";
                    break;
                default:
                    // FK, PK, Indexed, Not Null
                    SelectedImageKey = ImageKey = $"Column_{options:0000}";
                    break;
            }
        }

        public void AdjustColumnNameFixedWidth(int fixedWidth)
        {
            Text = ColumnName.PadRight(fixedWidth) + (string.IsNullOrEmpty(ColumnType) ? "" : "  " + ColumnType);
        }
    }
}
