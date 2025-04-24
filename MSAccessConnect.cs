using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Serialization;
using NppDB.Comm;
using NppDB.MSAccess.Properties;

namespace NppDB.MSAccess
{
    [XmlRoot]
    [ConnectAttr(Id = "MSAccessConnect", Title = "MS Access")]
    public class MsAccessConnect : TreeNode, IDbConnect, IRefreshable, IMenuProvider, IIconProvider, INppDBCommandClient
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

        public Bitmap GetIcon()
        {
            return Resources.Access;
        }

        public bool CheckLogin()
{
    ToolTipText = ServerAddress;
    var isNewFileScenario = false;
    DialogResult fileDialogResult = DialogResult.Cancel;

    if (string.IsNullOrEmpty(ServerAddress) || !File.Exists(ServerAddress))
    {
        isNewFileScenario = true; // Tentatively set, will confirm based on dlg.IsNew later

        if (!string.IsNullOrEmpty(ServerAddress) &&
            MessageBox.Show("file(" + ServerAddress + ") don't existed.\ncreate a new database?", @"Alert",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
        {
             return false;
        }

        using(var dlg = new FrmMsAccessConnect())
        {
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                Stream resource = null;
                FileDialog fdlg;
                bool creatingNewFile = dlg.IsNew;

                if (creatingNewFile)
                {
                     fdlg = new SaveFileDialog { Title = @"New Access File" };
                     resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("NppDB.MSAccess.Resources.empty.accdb");
                     if (resource == null) { /* Handle error: resource not found */ MessageBox.Show("Error: Embedded empty database resource not found.", "Resource Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return false; }
                     fdlg.Filter = @"Access Database (*.accdb)|*.accdb|Access 2000-2003 Database (*.mdb)|*.mdb";
                     fdlg.DefaultExt = "accdb";
                }
                else
                {
                     fdlg = new OpenFileDialog { Title = @"Open Access File" };
                     fdlg.Filter = @"Access Databases (*.accdb;*.mdb)|*.accdb;*.mdb|All Files(*.*)|*.*";
                }

                if (!string.IsNullOrEmpty(ServerAddress))
                    fdlg.InitialDirectory = Path.GetDirectoryName(ServerAddress);
                fdlg.AddExtension = true;

                fileDialogResult = fdlg.ShowDialog();

                if (fileDialogResult == DialogResult.OK)
                {
                    if (resource != null && creatingNewFile)
                    {
                        try
                        {
                            using (resource)
                            using (var destinationFile = File.Create(fdlg.FileName))
                            {
                                 resource.Seek(0, SeekOrigin.Begin);
                                 resource.CopyTo(destinationFile);
                            }
                        }
                        catch(Exception ex)
                        {
                             MessageBox.Show($"Failed to create new database file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                             return false;
                        }
                    }
                    ServerAddress = fdlg.FileName;
                    Text = GetDefaultTitle();
                    Password = "";
                    isNewFileScenario = creatingNewFile; // Correctly set based on user intent
                }
                else
                {
                     return false; // FileDialog cancelled
                }
            }
            else
            {
                 return false; // FrmMsAccessConnect cancelled
            }
        }
    }

    if (!isNewFileScenario) // Existing file path
    {
        var testBuilder = new OleDbConnectionStringBuilder
        {
            Provider = "Microsoft.ACE.OLEDB.12.0",
            DataSource = ServerAddress,
            OleDbServices = -4 // Keep pooling disabled
        };
        using (var testConn = new OleDbConnection(testBuilder.ConnectionString))
        {
            try
            {
                testConn.Open();
                Password = ""; // Clear password if connection succeeded without one
                return true; // Success, no password needed
            }
            catch (OleDbException oleEx)
            {
                // Check for password-related errors
                if (oleEx.ErrorCode != -2147217900 &&
                    (oleEx.Message == null || !oleEx.Message.ToLowerInvariant().Contains("password")))
                    return false; // Failed for other OLEDB reason
                // Password required, prompt user
                using (var pdlg = new FrmPassword { VisiblePassword = false })
                {
                    if (pdlg.ShowDialog() != DialogResult.OK) return false; // User cancelled password entry
                    Password = pdlg.Password; // Store entered password
                    return true; // Proceed with stored password

                }

                // Other OLE DB Error
            }
            catch (Exception ex) // Other unexpected error during trial
            {
                 MessageBox.Show($"An unexpected error occurred while testing the connection:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 return false; // Failed for other reason
            }
        } // End using testConn
    }

    // New file scenario
    Password = ""; // Assume no password for new file
    return true; // Proceed for new file
}

        public void Connect()
        {
            if (_conn != null)
            {
                 if (_conn.State != ConnectionState.Closed)
                 {
                     try { _conn.Close(); } catch { }
                 }
                 _conn.Dispose();
                 _conn = null;
            }

            _conn = new OleDbConnection();
            string connectionStringToUse = GetConnectionString(); // Calls the instrumented version above

            // DEBUG: Show the connection string (mask password for safety)
            string maskedConnectionString = connectionStringToUse;
            if(connectionStringToUse.IndexOf("Password=", StringComparison.OrdinalIgnoreCase) >= 0) // Basic check if password likely present
            {
                try {
                    // Attempt safer masking using Regex if available, otherwise basic replace
                     maskedConnectionString = System.Text.RegularExpressions.Regex.Replace(connectionStringToUse, @"(Jet\sOLEDB:Database\sPassword=)[^;]+", "$1*****", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                } catch {
                    maskedConnectionString = connectionStringToUse.Replace(Password,"*****"); // Less robust masking
                }
            }

            _conn.ConnectionString = connectionStringToUse;

            try
            {
                _conn.Open();
            }
            catch (OleDbException oleEx) // Catch specific first
            {
                // DEBUG: Show raw OleDbException details from this main connect attempt
                string errorDetails = $"Connect FAILED (OleDbException):\n\nErrorCode: {oleEx.ErrorCode}\nHResult: {oleEx.HResult}\nMessage: {oleEx.Message}\n\nConnectionString Used (masked):\n{maskedConnectionString}";
                MessageBox.Show(errorDetails, "Connect OLEDB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (_conn != null) { _conn.Dispose(); _conn = null; }
                throw; // Re-throw original exception
            }
            catch (Exception ex) // Catch other exceptions
            {
                // DEBUG: Show raw generic exception details
                string errorDetails = $"Connect FAILED (Generic Exception):\n\nType: {ex.GetType().Name}\nMessage: {ex.Message}\n\nConnectionString Used (masked):\n{maskedConnectionString}";
                MessageBox.Show(errorDetails, "Connect Generic Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (_conn != null) { _conn.Dispose(); _conn = null; }
                throw; // Re-throw original exception
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
            CommandHost.Execute(NppDBCommandType.CreateResultView, new[] { id, this, CreateSqlExecutor() });
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
        public string DatabaseSystemName => "MSAccess";
        internal INppDBCommandHost CommandHost { get; private set; }

        internal OleDbConnection GetConnection()
        {
            return new OleDbConnection(GetConnectionString());
        }

        public ISQLExecutor CreateSqlExecutor()
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
                    Nodes.Add(new MSAccessDatabase { Text = @"default" });
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
                    host.Execute(NppDBCommandType.CreateResultView, new[] { id, connect, CreateSqlExecutor() });
                }));
                if (host.Execute(NppDBCommandType.GetAttachedBufferID, null) == null)
                {
                    menuList.Items.Add(new ToolStripButton("Attach", null, (s, e) =>
                    {
                        var id = host.Execute(NppDBCommandType.GetActivatedBufferID, null);
                        host.Execute(NppDBCommandType.CreateResultView, new[] { id, connect, CreateSqlExecutor() });
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
                OleDbServices = -4 // Keep pooling disabled for now
            };

            if (!string.IsNullOrEmpty(Password))
            {
                // Standard property name
                builder.Add("Jet OLEDB:Database Password", Password);
            }

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
