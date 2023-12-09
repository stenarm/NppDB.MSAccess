using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;
using System.Windows.Forms;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    [XmlRoot]
    [ConnectAttr(Id = "MSAccessConnect", Title = "MS Access")]
    public class MSAccessConnect : TreeNode, IDBConnect, IRefreshable, IMenuProvider, IIconProvider, INppDBCommandClient
    {
        [XmlElement]
        public string Title { set => Text = value; get => Text; }
        [XmlElement]
        public string ServerAddress { set; get; }
        public string Account { set; get; }
        [XmlIgnore]
        public string Password { set; get; }
        private OleDbConnection _conn;

        public string GetDefaultTitle()
        {
            return string.IsNullOrEmpty(ServerAddress) ? "" : Path.GetFileName(ServerAddress);
        }

        public System.Drawing.Bitmap GetIcon()
        {
            return Properties.Resources.Access;
        }

        public bool CheckLogin()
        {
            ToolTipText = ServerAddress;
            var shownPwd = false;
            if (string.IsNullOrEmpty(ServerAddress) || !File.Exists(ServerAddress))
            {
                if (!string.IsNullOrEmpty(ServerAddress) &&
                    MessageBox.Show("file(" + ServerAddress + ") don't existed.\ncreate a new database?", "Alert", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    return false;

                var dlg = new frmMSAccessConnect();
                dlg.ServerAddress = ServerAddress;

                if (dlg.ShowDialog() != DialogResult.OK) return false;

                Stream resource = null;
                FileDialog fdlg;
                shownPwd = dlg.IsNew;
                if (dlg.IsNew)
                {
                    fdlg = new SaveFileDialog();
                    fdlg.Title = "New SQLite File";
                    resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("NppDB.MSAccess.Resources.empty.accdb");
                }
                else
                {
                    fdlg = new OpenFileDialog();
                    fdlg.Title = "Open SQLite File";
                }
                if (!string.IsNullOrEmpty(ServerAddress))
                    fdlg.InitialDirectory = Path.GetDirectoryName(ServerAddress);
                fdlg.AddExtension = false;
                fdlg.DefaultExt = ".*";
                fdlg.Filter = "All Files(*.*)|*.*";
                var result = fdlg.ShowDialog();
                if (result != DialogResult.OK) return false;

                if (resource != null)
                {
                    using (var destinationFile = File.Create(fdlg.FileName))
                    {
                        resource.Seek(0, SeekOrigin.Begin);
                        resource.CopyTo(destinationFile);
                    }
                }
                ServerAddress = fdlg.FileName;
            }
            var pdlg = new frmPassword { VisiblePassword = shownPwd };
            if (pdlg.ShowDialog() != DialogResult.OK) return false;
            Password = pdlg.Password;
            return true;
        }

        public void Connect()
        {
            if (_conn == null) _conn = new OleDbConnection();
            var curConnStr = GetConnectionString();
            if (_conn.ConnectionString != curConnStr) _conn.ConnectionString = curConnStr;
            if (_conn.State == ConnectionState.Open) return;
            
            try
            {
                _conn.Open();
            }
            catch (Exception ex)
            {
                throw new ApplicationException("connect fail", ex);
            }
        }

        public void Attach()
        {
            var id = CommandHost.Execute(NppDBCommandType.GetAttachedBufferID, null);
            if (id != null)
            {
                CommandHost.Execute(NppDBCommandType.NewFile, null);
            }
            id = CommandHost.Execute(NppDBCommandType.GetActivatedBufferID, null);
            CommandHost.Execute(NppDBCommandType.CreateResultView, new[] { id, this, CreateSQLExecutor() });
        }

        public string ConnectAndAttach()
        {
            if (IsOpened) return "CONTINUE";
            if (!CheckLogin()) return "FAIL";
            try
            {
                Connect();
                Attach();
                Refresh();
                return "FRESH_NODES";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + (ex.InnerException != null ? " : " + ex.InnerException.Message : ""));
            }
            return "FAIL";
        }

        public void Disconnect()
        {
            if (_conn == null || _conn.State == ConnectionState.Closed) return;
            
            _conn.Close();
        }

        public bool IsOpened => _conn != null && _conn.State == ConnectionState.Open;
        internal INppDBCommandHost CommandHost { get; private set; }

        internal OleDbConnection GetConnection()
        {
            return new OleDbConnection(GetConnectionString());
        }

        public ISQLExecutor CreateSQLExecutor()
        {
            return new MSAccessExecutor(GetConnection);
        }

        public void Refresh()
        {
            using (var conn = GetConnection())
            {
                TreeView.Cursor = Cursors.WaitCursor;
                TreeView.Enabled = false;
                try
                {
                    conn.Open();
                    var dt = conn.GetSchema(OleDbMetaDataCollectionNames.Catalogs);
                    
                    Nodes.Clear();
                    foreach (DataRow row in dt.Rows)
                    {
                        var db = new MSAccessDatabase { Text = row["CATALOG_NAME"].ToString() };
                        Nodes.Add(db);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Nodes.Clear();
                    Nodes.Add(new MSAccessDatabase { Text = "default" });
                }
                finally
                {
                    TreeView.Enabled = true;
                    TreeView.Cursor = null;
                }
            }
        }

        public ContextMenuStrip GetMenu()
        {
            var menuList = new ContextMenuStrip { ShowImageMargin = false };
            var connect = this;
            var host = CommandHost;
            if (host != null)
            {
                menuList.Items.Add(new ToolStripButton("Open", null, (s, e) =>
                {
                    host.Execute(NppDBCommandType.NewFile, null);
                    var id = host.Execute(NppDBCommandType.GetActivatedBufferID, null);
                    host.Execute(NppDBCommandType.CreateResultView, new[] { id, connect, CreateSQLExecutor() });
                }));
                if (host.Execute(NppDBCommandType.GetAttachedBufferID, null) == null)
                {
                    menuList.Items.Add(new ToolStripButton("Attach", null, (s, e) =>
                    {
                        var id = host.Execute(NppDBCommandType.GetActivatedBufferID, null);
                        host.Execute(NppDBCommandType.CreateResultView, new[] { id, connect, CreateSQLExecutor() });
                    }));
                }
                else
                {
                    menuList.Items.Add(new ToolStripButton("Detach", null, (s, e) =>
                    {
                        host.Execute(NppDBCommandType.DestroyResultView, null);
                    }));
                }
                menuList.Items.Add(new ToolStripSeparator());
            }
            menuList.Items.Add(new ToolStripButton("Refresh", null, (s, e) => { Refresh(); }));
            return menuList;
        }

        internal string GetConnectionString()
        {
            var builder = new OleDbConnectionStringBuilder
            {
                Provider = "Microsoft.ACE.OLEDB.12.0",
                DataSource = ServerAddress,
            };
            return builder.ConnectionString;
        }

        public void Reset()
        {
            Title = ""; ServerAddress = ""; Account = ""; Password = "";
            Disconnect();
            _conn = null;
        }

        public void SetCommandHost(INppDBCommandHost host)
        {
            CommandHost = host;
        }
    }

}
