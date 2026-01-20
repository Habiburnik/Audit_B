using Microsoft.Data.SqlClient;
using System.Net;
using System.Reflection;

namespace Audit_B;

public partial class LoginForm : Form
{
    private SqlConnection? dbConnection;
    private Label? lblLoginUsername;
    private Label? lblPassword;
    private Label? lblUserAuthentication;
    private TextBox? edtLoginUsername;
    private TextBox? edtLoginPassword;
    private Button? btnLogin;
    private Button? btnCancel;
    private const string UPDATE_URL = "http://10.0.0.24:8434/Soft/Audit_B.exe";
    private const string DOWNLOAD_PATH = @"C:\New Software";
    public LoginForm()
    {
        InitializeComponent();
        InitializeDatabase();

        CheckForUpdates();
        string exePath = Application.StartupPath;
        if (exePath.StartsWith(@"\\"))
        {
            MessageBox.Show(
                "Please copy this application to your PC and run it from there.",
                "Cannot Run on Network Path",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            Application.Exit(); // Close the application
        }
    }

    private void InitializeComponent()
    {
        this.Text = "User Login Authentication (Audit_B)";
        this.Size = new Size(550, 260);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        // this.Icon = new Icon("AuditB_icon.ico");
        // this.BackColor = Color.FromArgb(139, 175, 185);
        try
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("Audit_B.AuditB_icon.ico"))
            {
                if (stream != null)
                {
                    this.Icon = new Icon(stream);
                }
            }
        }
        catch { }

        // Logo
        Label lblLogoArea = new Label();
        lblLogoArea.Size = new Size(180, 100);
        lblLogoArea.Location = new Point(15, 50);
        lblLogoArea.BorderStyle = BorderStyle.None;
        // lblLogoArea.BackColor = Color.FromArgb(139, 175, 185);
        lblLogoArea.Text = "GHIT";
        lblLogoArea.Font = new Font("Arial", 42, FontStyle.Bold);
        lblLogoArea.ForeColor = Color.FromArgb(100, 120, 140);
        lblLogoArea.TextAlign = ContentAlignment.MiddleCenter;
        this.Controls.Add(lblLogoArea);

        // Title
        Label lblTitle = new Label();
        lblTitle.Text = "User Authentication";
        lblTitle.Location = new Point(200, 25);
        lblTitle.Font = new Font("Arial", 18, FontStyle.Bold);
        lblTitle.AutoSize = true;
        lblTitle.ForeColor = Color.Black;

        // Username Label
        lblLoginUsername = new Label();
        lblLoginUsername.Text = "Username :";
        lblLoginUsername.Location = new Point(200, 87);
        lblLoginUsername.AutoSize = true;
        lblLoginUsername.Font = new Font("Arial", 11);

        // Username TextBox
        edtLoginUsername = new TextBox();
        edtLoginUsername.Location = new Point(285, 85);
        edtLoginUsername.Size = new Size(200, 25);
        edtLoginUsername.Font = new Font("Arial", 11);

        // Password Label
        lblPassword = new Label();
        lblPassword.Text = "Password :";
        lblPassword.Location = new Point(200, 127);
        lblPassword.AutoSize = true;
        lblPassword.Font = new Font("Arial", 11);

        // Password TextBox
        edtLoginPassword = new TextBox();
        edtLoginPassword.Location = new Point(285, 125);
        edtLoginPassword.Size = new Size(200, 25);
        edtLoginPassword.PasswordChar = '*';
        edtLoginPassword.Font = new Font("Arial", 11);
        edtLoginPassword.KeyPress += EdtLoginPassword_KeyPress;

        // Authentication Message Label
        lblUserAuthentication = new Label();
        lblUserAuthentication.Text = "";
        lblUserAuthentication.Location = new Point(285, 150);
        lblUserAuthentication.Size = new Size(320, 20);
        lblUserAuthentication.ForeColor = Color.Red;
        lblUserAuthentication.Font = new Font("Arial", 9);

        // Login Button
        btnLogin = new Button();
        btnLogin.Text = "✓ Login";
        btnLogin.Location = new Point(280, 170);
        btnLogin.Size = new Size(80, 35);
        btnLogin.Font = new Font("Arial", 11, FontStyle.Bold);
        btnLogin.BackColor = Color.FromArgb(106, 153, 78);
        btnLogin.ForeColor = Color.White;
        btnLogin.Click += BtnLogin_Click;
        btnLogin.Cursor = Cursors.Hand;

        // Cancel Button
        btnCancel = new Button();
        btnCancel.Text = "✕ Cancel";
        btnCancel.Location = new Point(380, 170);
        btnCancel.Size = new Size(80, 35);
        btnCancel.Font = new Font("Arial", 11, FontStyle.Bold);
        btnCancel.BackColor = Color.FromArgb(195, 71, 73);
        btnCancel.ForeColor = Color.White;
        btnCancel.Click += BtnCancel_Click;
        btnCancel.Cursor = Cursors.Hand;


        this.Controls.AddRange(new Control[] {
            lblTitle, lblLoginUsername, edtLoginUsername,
            lblPassword, edtLoginPassword, lblUserAuthentication,
            btnLogin, btnCancel
        });
    }

    private void InitializeDatabase()
    {
        try
        {
            string connectionString = ConfigurationHelper.GetConnectionString();
            dbConnection = new SqlConnection(connectionString);
        }
        catch (Exception)
        {
            MessageBox.Show($"Database initialization error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void EdtLoginPassword_KeyPress(object? sender, KeyPressEventArgs e)
    {
        if (e.KeyChar == (char)Keys.Return)
        {
            e.Handled = true;
            BtnLogin_Click(sender, EventArgs.Empty);
        }
    }

    private void BtnLogin_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(edtLoginUsername?.Text) || string.IsNullOrWhiteSpace(edtLoginPassword?.Text))
        {
            lblUserAuthentication?.Text = "Please enter username and password";
            return;
        }

        if (ValidateLogin(edtLoginUsername.Text, edtLoginPassword.Text))
        {
            // Open main menu form
            MainMenuForm mainForm = new MainMenuForm(edtLoginUsername.Text);
            mainForm.Show();
            this.Hide();
        }
        else
        {
            lblUserAuthentication?.Text = "Invalid credentials";
            edtLoginPassword?.Clear();
            edtLoginPassword?.Focus();
        }
    }

    private void BtnCancel_Click(object? sender, EventArgs e)
    {
        Application.Exit();
    }

    private async void CheckForUpdates()
    {
        try
        {
            await Task.Run(() =>
            {
                DateTime localBuildDate = GetLocalBuildDate();
                DateTime? remoteBuildDate = GetRemoteFileModifiedDate(UPDATE_URL);

                if (remoteBuildDate.HasValue && remoteBuildDate.Value > localBuildDate)
                {
                    this.Invoke((Action)delegate
                    {
                        DownloadUpdate();
                    });
                }
            });
        }
        catch
        {
            // Silently fail - don't block login
        }
    }

    private DateTime GetLocalBuildDate()
    {
        string exePath = Assembly.GetExecutingAssembly().Location;
        return File.GetLastWriteTime(exePath);
    }

    private DateTime? GetRemoteFileModifiedDate(string url)
    {
        try
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "HEAD";
            request.Timeout = 5000;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                return response.LastModified;
            }
        }
        catch
        {
            return null;
        }
    }

    private async void DownloadUpdate()
    {
        try
        {
            // Create download directory if it doesn't exist
            if (!Directory.Exists(DOWNLOAD_PATH))
            {
                Directory.CreateDirectory(DOWNLOAD_PATH);
            }

            string fileName = Path.GetFileName(UPDATE_URL);
            string downloadFilePath = Path.Combine(DOWNLOAD_PATH, fileName);

            // Show progress
            Form progressForm = new Form
            {
                Text = "Downloading Update",
                Size = new Size(400, 150),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            Label lblProgress = new Label
            {
                Text = "Downloading...",
                Location = new Point(20, 20),
                AutoSize = true
            };

            ProgressBar progressBar = new ProgressBar
            {
                Location = new Point(20, 50),
                Size = new Size(340, 30),
                Style = ProgressBarStyle.Continuous
            };

            progressForm.Controls.Add(lblProgress);
            progressForm.Controls.Add(progressBar);
            progressForm.Show();

            using (WebClient client = new WebClient())
            {
                client.DownloadProgressChanged += (s, e) =>
                {
                    progressBar.Value = e.ProgressPercentage;
                    lblProgress.Text = $"Downloading... {e.ProgressPercentage}%";
                };

                client.DownloadFileCompleted += (s, e) =>
                {
                    progressForm.Close();

                    if (e.Error != null)
                    {
                        MessageBox.Show($"Download failed:", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        MessageBox.Show(
                            $"Update downloaded successfully!. Please take the update software from {downloadFilePath}",
                            "Download Complete",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        Application.Exit();
                    }
                };

                await client.DownloadFileTaskAsync(new Uri(UPDATE_URL), downloadFilePath);
            }
        }
        catch (Exception)
        {
            MessageBox.Show($"Download error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }


    private bool ValidateLogin(string username, string password)
    {
        try
        {
            if (dbConnection?.State != System.Data.ConnectionState.Open)
                dbConnection?.Open();

            string query = "SELECT user_Password, AuditB FROM Employee_INFO WHERE User_name = @Username";
            SqlCommand cmd = new SqlCommand(query, dbConnection);
            cmd.Parameters.AddWithValue("@Username", username);
            SqlDataReader reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                reader.Read();
                string userPassword = reader["user_Password"].ToString() ?? "";
                int auditB = reader.IsDBNull(reader.GetOrdinal("AuditB")) ? 0 : reader.GetInt32(reader.GetOrdinal("AuditB"));
                reader.Close();

                // Check password first
                if (userPassword != password)
                {
                    MessageBox.Show("Invalid password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Check access permission
                if (auditB == -1 || auditB == 1 || auditB == 2 || auditB == 3)
                {
                    return true;
                }
                else
                {
                    MessageBox.Show("You do not have access for this software", "Access Denied",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            reader.Close();
            MessageBox.Show("Invalid username", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Login error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        finally
        {
            if (dbConnection?.State == System.Data.ConnectionState.Open)
                dbConnection?.Close();
        }
    }
}