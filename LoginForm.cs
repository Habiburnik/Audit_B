
using Microsoft.Data.SqlClient;
using System.Reflection;
using MethodInvoker = System.Windows.Forms.MethodInvoker;
using Timer = System.Windows.Forms.Timer;
using System.Net.Http;
using System.IO;
using System.Threading.Tasks;

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
    public LoginForm()
    {
        InitializeComponent();
        InitializeDatabase();

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
        lblUserAuthentication.Font = new Font("Arial", 10);

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
        catch (Exception ex)
        {
            MessageBox.Show($"Database initialization error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

    private bool ValidateSoftwareName()
    {
        try
        {
            string? assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            return assemblyName == "Audit_B";        
        }
        catch
        {
            return false;
        }
    }

    private bool ValidateLogin(string username, string password)
    {
        try
        {
            if (dbConnection?.State != System.Data.ConnectionState.Open)
                dbConnection?.Open();

            // // Check if login is numeric (Operator ID)
            // if (int.TryParse(username, out int operatorId))
            // {
            //     string query = "SELECT passwd, status FROM Operators WHERE operatorId = @Username";
            //     SqlCommand cmd = new SqlCommand(query, dbConnection);
            //     cmd.Parameters.AddWithValue("@Username", username);

            //     SqlDataReader reader = cmd.ExecuteReader();
            //     if (reader.HasRows)
            //     {
            //         reader.Read();
            //         string passwd = reader["passwd"].ToString() ?? "";
            //         int status = (int)reader["status"];
            //         reader.Close();

            //         return passwd == password && status == 0;
            //     }
            //     reader.Close();
            // }
            // else
            {
                // Employee login
                string query = "SELECT user_Password FROM Employee_INFO WHERE User_name = @Username";
                SqlCommand cmd = new SqlCommand(query, dbConnection);
                cmd.Parameters.AddWithValue("@Username", username);

                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    string userPassword = reader["user_Password"].ToString() ?? "";
                    reader.Close();

                    return userPassword == password;
                }
                reader.Close();
            }

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