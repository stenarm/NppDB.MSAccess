
namespace NppDB.MSAccess
{
    internal class MSAccessView : MSAccessTable
    {
        public MSAccessView()
        {
            TypeName = "VIEW";
            SelectedImageKey = ImageKey = "Table";
        }
    }
}
