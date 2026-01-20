using System.Data;
using Microsoft.Data.SqlClient;
namespace Audit_B;

public partial class AddNewForm : Form
{
    private SqlConnection? dbConnection;
    private MainMenuForm? mainMenuForm;
    // Left: Add Project
    private GroupBox? grpAddProject;
    private Label? lblProjectCode;
    private TextBox? edtProjectCode;
    private Label? lblTotalBatch;
    private TextBox? edtTotalBatch;
    private Label? lblLanguage;
    private TextBox? edtLanguage;
    private Label? lblDeadline;
    private TextBox? edtDeadline;
    private Button? btnAddProject;

    // Middle: Add Batches
    private GroupBox? grpAddBatches;
    private RadioButton? rb1, rb2, rb3, rb4, rb5;
    private Label? lblSelectProjectCode;
    private ComboBox? cbProjectCode;
    private TextBox? txtBatchList;
    private Button? btnAddBatches;

    // Right: User Permissions
    private GroupBox? grpUserPermissions;
    private Label? lblCurrentUsers;
    private ListBox? lstCurrentUsers;
    private Button? btnShowCurrentUsers;
    private Button? btnRemoveAccess;
    private Label? lblSearchUser;
    private TextBox? txtSearchUser;
    private Button? btnSearchUser;
    private Label? lblFoundUsers;
    private ListBox? lstFoundUsers;
    private Button? btnGiveAccess;

    public AddNewForm(MainMenuForm parentForm)
    {
        mainMenuForm = parentForm;
        InitializeComponent();
        InitializeDatabase();
    }

    private void InitializeComponent()
    {
        this.Text = "Add Projects / Batches / Manage Users";
        this.Size = new Size(1350, 500);
        this.StartPosition = FormStartPosition.CenterParent;

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



        // ===== GROUP 1: ADD PROJECT (LEFT) =====
        grpAddProject = new GroupBox();
        grpAddProject.Text = "Add Projects";
        grpAddProject.Location = new Point(10, 10);
        grpAddProject.Size = new Size(420, 280);

        lblProjectCode = new Label() { Text = "Project_Code :", Location = new Point(15, 30), AutoSize = true };
        edtProjectCode = new TextBox() { Location = new Point(140, 25), Size = new Size(250, 24) };

        lblTotalBatch = new Label() { Text = "Total_Batch :", Location = new Point(15, 70), AutoSize = true };
        edtTotalBatch = new TextBox() { Location = new Point(140, 65), Size = new Size(250, 24) };

        lblLanguage = new Label() { Text = "Language :", Location = new Point(15, 110), AutoSize = true };
        edtLanguage = new TextBox() { Location = new Point(140, 105), Size = new Size(250, 24) };

        lblDeadline = new Label() { Text = "Deadline :", Location = new Point(15, 150), AutoSize = true };
        edtDeadline = new TextBox() { Location = new Point(140, 145), Size = new Size(250, 24) };

        btnAddProject = new Button() { Text = "Add Project", Location = new Point(150, 200), Size = new Size(120, 30) };
        btnAddProject.Click += BtnAddProject_Click;

        grpAddProject.Controls.AddRange(new Control[] {
            lblProjectCode, edtProjectCode,
            lblTotalBatch, edtTotalBatch,
            lblLanguage, edtLanguage,
            lblDeadline, edtDeadline,
            btnAddProject
        });

        // ===== GROUP 2: ADD BATCHES (MIDDLE) =====
        grpAddBatches = new GroupBox();
        grpAddBatches.Text = "Add Batches";
        grpAddBatches.Location = new Point(440, 10);
        grpAddBatches.Size = new Size(420, 280);

        rb1 = new RadioButton() { Text = "1st", Location = new Point(15, 20), AutoSize = true };
        rb2 = new RadioButton() { Text = "2nd", Location = new Point(70, 20), AutoSize = true };
        rb3 = new RadioButton() { Text = "3rd", Location = new Point(125, 20), AutoSize = true };
        rb4 = new RadioButton() { Text = "4th", Location = new Point(180, 20), AutoSize = true };
        rb5 = new RadioButton() { Text = "5th", Location = new Point(235, 20), AutoSize = true };
        rb1.Checked = true;

        lblSelectProjectCode = new Label() { Text = "Select_Project_Code :", Location = new Point(15, 50), AutoSize = true };
        cbProjectCode = new ComboBox() { Location = new Point(160, 47), Size = new Size(240, 24), DropDownStyle = ComboBoxStyle.DropDownList };

        txtBatchList = new TextBox() { Location = new Point(15, 85), Size = new Size(385, 120), Multiline = true, ScrollBars = ScrollBars.Vertical };

        btnAddBatches = new Button() { Text = "Add Batches", Location = new Point(130, 220), Size = new Size(120, 30) };
        btnAddBatches.Click += BtnAddBatches_Click;

        grpAddBatches.Controls.AddRange(new Control[] {
            rb1, rb2, rb3, rb4, rb5,
            lblSelectProjectCode, cbProjectCode, txtBatchList, btnAddBatches
        });

        // ===== GROUP 3: USER PERMISSIONS (RIGHT) =====
        grpUserPermissions = new GroupBox();
        grpUserPermissions.Text = "User Permissions";
        grpUserPermissions.Location = new Point(870, 10);
        grpUserPermissions.Size = new Size(450, 440);

        if (mainMenuForm != null && mainMenuForm.edtPermission?.Text == "Yes" && mainMenuForm.mmReport?.Visible == true)
        {
            grpUserPermissions?.Visible = true;
            this.Size = new Size(1350, 500);
        }
        else
        {
            grpUserPermissions?.Visible = false;
            this.Size = new Size(890, 340);
        }
        // Current Users Section
        lblCurrentUsers = new Label() { Text = "Current Users with Access:", Location = new Point(15, 25), AutoSize = true, Font = new Font("Arial", 9, FontStyle.Bold) };

        lstCurrentUsers = new ListBox();
        lstCurrentUsers.Location = new Point(15, 50);
        lstCurrentUsers.Size = new Size(410, 100);
        lstCurrentUsers.SelectionMode = SelectionMode.One;

        btnShowCurrentUsers = new Button() { Text = "Show Current Users", Location = new Point(15, 160), Size = new Size(140, 30) };
        btnShowCurrentUsers.Click += BtnShowCurrentUsers_Click;

        btnRemoveAccess = new Button() { Text = "Remove Access", Location = new Point(165, 160), Size = new Size(140, 30) };
        btnRemoveAccess.BackColor = Color.FromArgb(195, 71, 73);
        btnRemoveAccess.ForeColor = Color.White;
        btnRemoveAccess.Click += BtnRemoveAccess_Click;

        // Separator line
        Label separator = new Label() { Text = "─────────────────────────────────────────────", Location = new Point(15, 200), AutoSize = true, ForeColor = Color.Gray };

        // Give Access Section
        lblSearchUser = new Label() { Text = "Search User by Username:", Location = new Point(15, 225), AutoSize = true, Font = new Font("Arial", 9, FontStyle.Bold) };

        txtSearchUser = new TextBox();
        txtSearchUser.Location = new Point(15, 250);
        txtSearchUser.Size = new Size(300, 24);
        txtSearchUser.PlaceholderText = "Enter username to search...";

        btnSearchUser = new Button() { Text = "Search", Location = new Point(325, 248), Size = new Size(100, 28) };
        btnSearchUser.Click += BtnSearchUser_Click;

        lblFoundUsers = new Label() { Text = "Found Users:", Location = new Point(15, 285), AutoSize = true };

        lstFoundUsers = new ListBox();
        lstFoundUsers.Location = new Point(15, 310);
        lstFoundUsers.Size = new Size(410, 80);
        lstFoundUsers.SelectionMode = SelectionMode.One;

        btnGiveAccess = new Button() { Text = "Give Access", Location = new Point(150, 400), Size = new Size(140, 30) };
        btnGiveAccess.BackColor = Color.FromArgb(106, 153, 78);
        btnGiveAccess.Click += BtnGiveAccess_Click;
        btnGiveAccess.ForeColor = Color.White;

        grpUserPermissions.Controls.AddRange(new Control[] {
            lblCurrentUsers, lstCurrentUsers, btnShowCurrentUsers, btnRemoveAccess,
            separator,
            lblSearchUser, txtSearchUser, btnSearchUser, lblFoundUsers, lstFoundUsers, btnGiveAccess
        });

        this.Controls.AddRange(new Control[] { grpAddProject, grpAddBatches, grpUserPermissions });

        this.Shown += AddNewForm_Shown;
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

    private void AddNewForm_Shown(object? sender, EventArgs e)
    {
        LoadProjectsIntoCombo();
        LoadCurrentUsers(); // Load users on form show
    }

    private void LoadProjectsIntoCombo()
    {
        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string sql = @"SELECT project_code FROM audit_b_projects WHERE status IN (0, -1)
                           ORDER BY CASE WHEN status = -1 THEN 1 WHEN status = 0 THEN 2 WHEN status = -2 THEN 3 WHEN status = 1 THEN 4 ELSE 5 END, deadline";

            SqlCommand cmd = new SqlCommand(sql, dbConnection);
            SqlDataReader reader = cmd.ExecuteReader();
            cbProjectCode?.Items.Clear();
            while (reader.Read())
            {
                cbProjectCode?.Items.Add(reader["project_code"].ToString() ?? "");
            }
            reader.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading projects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    // ===== USER PERMISSION METHODS =====

    private void BtnShowCurrentUsers_Click(object? sender, EventArgs e)
    {
        LoadCurrentUsers();
    }

    private void LoadCurrentUsers()
    {
        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string sql = "SELECT USER_NAME, EMP_NAME FROM EMPLOYEE_INFO WHERE AuditB = -1 ORDER BY EMP_NAME";
            SqlCommand cmd = new SqlCommand(sql, dbConnection);
            SqlDataReader reader = cmd.ExecuteReader();

            lstCurrentUsers?.Items.Clear();
            while (reader.Read())
            {
                string userName = reader["USER_NAME"].ToString() ?? "";
                string empName = reader["EMP_NAME"].ToString() ?? "";
                lstCurrentUsers?.Items.Add($"{userName} - {empName}");
            }
            reader.Close();

            if (lstCurrentUsers?.Items.Count == 0)
            {
                lstCurrentUsers?.Items.Add("No users with access found");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading current users: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    private void BtnRemoveAccess_Click(object? sender, EventArgs e)
    {
        if (lstCurrentUsers?.SelectedIndex == -1 || lstCurrentUsers?.SelectedItem?.ToString() == "No users with access found")
        {
            MessageBox.Show("Please select a user to remove access", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string selectedItem = lstCurrentUsers?.SelectedItem?.ToString() ?? "";
        string userName = selectedItem.Split('-')[0].Trim();

        var confirm = MessageBox.Show($"Are you sure you want to remove access for user '{userName}'?",
            "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
            return;

        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string sql = "UPDATE EMPLOYEE_INFO SET AuditB = 0 WHERE USER_NAME = @username";
            SqlCommand cmd = new SqlCommand(sql, dbConnection);
            cmd.Parameters.AddWithValue("@username", userName);
            cmd.ExecuteNonQuery();

            MessageBox.Show("Access removed successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadCurrentUsers();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error removing access: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    private void BtnSearchUser_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSearchUser?.Text))
        {
            MessageBox.Show("Please enter a username to search", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string sql = "SELECT USER_NAME, EMP_NAME, card_no FROM EMPLOYEE_INFO WHERE USER_NAME LIKE @search AND AuditB = 0 ORDER BY EMP_NAME";
            SqlCommand cmd = new SqlCommand(sql, dbConnection);
            cmd.Parameters.AddWithValue("@search", "%" + txtSearchUser.Text + "%");
            SqlDataReader reader = cmd.ExecuteReader();

            lstFoundUsers?.Items.Clear();
            while (reader.Read())
            {
                string userName = reader["USER_NAME"].ToString() ?? "";
                string empName = reader["EMP_NAME"].ToString() ?? "";
                string cardNo = reader["CARD_NO"].ToString() ?? "";
                lstFoundUsers?.Items.Add($"{userName} - {empName} - {cardNo}");
            }
            reader.Close();

            if (lstFoundUsers?.Items.Count == 0)
            {
                lstFoundUsers?.Items.Add("No users found");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error searching users: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    private void BtnGiveAccess_Click(object? sender, EventArgs e)
    {
        if (lstFoundUsers?.SelectedIndex == -1 || lstFoundUsers?.SelectedItem?.ToString() == "No users found")
        {
            MessageBox.Show("Please select a user to give access", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string selectedItem = lstFoundUsers?.SelectedItem?.ToString() ?? "";
        string userName = selectedItem.Split('-')[0].Trim();

        var confirm = MessageBox.Show($"Are you sure you want to give access to user '{userName}'?",
            "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (confirm != DialogResult.Yes)
            return;

        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string sql = "UPDATE EMPLOYEE_INFO SET AuditB = -1 WHERE USER_NAME = @username";
            SqlCommand cmd = new SqlCommand(sql, dbConnection);
            cmd.Parameters.AddWithValue("@username", userName);
            cmd.ExecuteNonQuery();

            MessageBox.Show("Access granted successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Clear search and reload
            txtSearchUser?.Text = "";
            lstFoundUsers?.Items.Clear();
            LoadCurrentUsers();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error giving access: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    // ===== EXISTING METHODS =====

    private void BtnAddBatches_Click(object? sender, EventArgs e)
    {
        if (cbProjectCode?.SelectedIndex == -1)
        {
            MessageBox.Show("Select Project Code", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtBatchList?.Text))
        {
            MessageBox.Show("Add Batches", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var confirm = MessageBox.Show("Are you sure you want to Add Batches?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes)
            return;

        int lineValue = 1;
        if (rb2?.Checked == true) lineValue = 2;
        else if (rb3?.Checked == true) lineValue = 3;
        else if (rb4?.Checked == true) lineValue = 4;
        else if (rb5?.Checked == true) lineValue = 5;

        var batches = txtBatchList.Lines.Select(l => l.Trim()).Where(l => !string.IsNullOrEmpty(l)).ToList();
        List<string> successfulBatches = new List<string>();
        List<string> failedBatches = new List<string>();

        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            foreach (var batch in batches)
            {
                string insertSql = "INSERT INTO audit_b (project_code, batch_name, line) VALUES (@project_code, @batch_name, @line)";
                SqlCommand insCmd = new SqlCommand(insertSql, dbConnection);
                insCmd.Parameters.AddWithValue("@project_code", cbProjectCode?.Text);
                insCmd.Parameters.AddWithValue("@batch_name", batch);
                insCmd.Parameters.AddWithValue("@line", lineValue);
                try
                {
                    insCmd.ExecuteNonQuery();
                    successfulBatches.Add(batch);
                    RemoveBatchFromTextbox(batch);
                }
                catch
                {
                    failedBatches.Add(batch);
                }
            }

            // Show results
            string message = "";
            if (successfulBatches.Count > 0)
            {
                message = $"{successfulBatches.Count} Batch(es) Added Successfully.";
            }

            if (failedBatches.Count > 0)
            {
                if (!string.IsNullOrEmpty(message))
                    message += "\n\n";
                message += $"{failedBatches.Count} Duplicate batch add fail";
            }

            if (!string.IsNullOrEmpty(message))
            {
                MessageBox.Show(message, "Batch Add Result", MessageBoxButtons.OK,
                    failedBatches.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }

            if (successfulBatches.Count == batches.Count)
            {
                cbProjectCode?.SelectedIndex = -1;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding batches: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }
    private void RemoveBatchFromTextbox(string batchToRemove)
    {
        if (txtBatchList == null) return;

        var remainingLines = txtBatchList.Lines
            .Where(line => line.Trim() != batchToRemove.Trim())
            .ToArray();

        txtBatchList.Lines = remainingLines;
    }

    private void BtnAddProject_Click(object? sender, EventArgs e)
    {
        string? projectCode = edtProjectCode?.Text.Trim();
        string? totalBatch = edtTotalBatch?.Text.Trim();
        string? language = edtLanguage?.Text.Trim();
        string? deadline = edtDeadline?.Text.Trim();

        if (string.IsNullOrEmpty(projectCode) || string.IsNullOrEmpty(totalBatch) || string.IsNullOrEmpty(language) || string.IsNullOrEmpty(deadline))
        {
            MessageBox.Show("Please fill all fields", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string checkSql = "SELECT project_code FROM audit_b_projects WHERE project_code = @project_code";
            SqlCommand checkCmd = new SqlCommand(checkSql, dbConnection);
            checkCmd.Parameters.AddWithValue("@project_code", projectCode);
            var reader = checkCmd.ExecuteReader();
            bool exists = reader.HasRows;
            reader.Close();

            if (exists)
            {
                MessageBox.Show("Project Already Exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var confirm = MessageBox.Show("Are you sure you want to Add Project?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes)
                return;

            string maxSql = "SELECT ISNULL(MAX(project_id),0) FROM audit_b_projects";
            SqlCommand maxCmd = new SqlCommand(maxSql, dbConnection);
            int projectId = Convert.ToInt32(maxCmd.ExecuteScalar()) + 1;

            string insertSql = @"INSERT INTO audit_b_projects(project_id, project_code, total_batch, language, deadline)
                                 VALUES (@project_id, @project_code, @total_batch, @language, @deadline)";
            SqlCommand insertCmd = new SqlCommand(insertSql, dbConnection);
            insertCmd.Parameters.AddWithValue("@project_id", projectId);
            insertCmd.Parameters.AddWithValue("@project_code", projectCode);
            insertCmd.Parameters.AddWithValue("@total_batch", totalBatch);
            insertCmd.Parameters.AddWithValue("@language", language);
            insertCmd.Parameters.AddWithValue("@deadline", deadline);
            insertCmd.ExecuteNonQuery();

            MessageBox.Show("Project Added Successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            edtProjectCode?.Clear();
            edtLanguage?.Clear();
            edtTotalBatch?.Clear();
            edtDeadline?.Clear();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding project: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }
}