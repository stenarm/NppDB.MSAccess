
namespace NppDB.MSAccess
{
    internal class MSAccessView : MsAccessTable
    {
        public MSAccessView()
        {
            TypeName = "VIEW";
            SelectedImageKey = ImageKey = "Table";
        }
    }
}
