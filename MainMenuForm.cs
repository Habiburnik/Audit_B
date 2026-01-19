using System.Data;
using Microsoft.Data.SqlClient;

namespace Audit_B;

public partial class MainMenuForm : Form
{
    private SqlConnection? dbConnection;
    // UI Components
    private MenuStrip? mainMenu;
    private Panel? pnlInfo;
    private ListBox? lbBatchList;
    public TextBox? edtUsername;
    private Label? lblUsername;
    private TextBox? edtBatchName;
    private Label? lblBatchName;
    private Label? lblTime;
    private TextBox? edtTime;
    private Button? btnStart;
    private Button? btnComplete;
    private Button? btnAuditSuspend;
    private Label? lblSampleCount;
    private TextBox? edtSampleCount;
    private TextBox? edtUserId;
    private Label? label1;
    private ListBox? lbProjectList;
    private Label? label2;
    private Label? lblBatchName2;
    private Button? btnAuditCancel;
    private TextBox? edtEmpType;
    private TextBox? edtLineNO;
    public TextBox? edtPermission;
    private CheckBox? checkBoxSuspended;
    private Label? label3;
    private TextBox? edtComments;
    private string currentUsername;
    public ToolStripMenuItem? mmAdd;
    public ToolStripMenuItem? mmReport;  // ADD THIS - make it a class field
    // ... rest of your fields
    public MainMenuForm(string username = "")
    {
        currentUsername = username;
        InitializeDatabase();
        InitializeComponent();

        // this.mdicontainer = true;
    }

    private void InitializeComponent()
    {
        // DPI Scaling
        this.Text = "Audit B - Main Menu";
        // this.WindowState = FormWindowState.Maximized;
        this.StartPosition = FormStartPosition.Manual;
        this.BackColor = Color.FromArgb(230, 230, 230);
        this.FormClosing += MainMenuFormClosing;
        this.IsMdiContainer = true;
        this.Height = 380;  
        this.Width = 1030;
        this.Location = new Point(170, 170);
        // this.Icon = new Icon("AuditB_icon.ico");

        // Main Menu
        mainMenu = new MenuStrip();
        var mmFile = new ToolStripMenuItem("&File");
        mmAdd = new ToolStripMenuItem("&Add New", null, MMAdd_Click);  // Remove 'var'
        mmReport = new ToolStripMenuItem("&Reports");
        var reportByAuditor = new ToolStripMenuItem("Report By &Auditor", null, ReportByAuditor_Click);
        var reportByBatch = new ToolStripMenuItem("Report By &Batch", null, ReportByBatch_Click);
        var refresh = new ToolStripMenuItem("&Refresh", null, Refresh_Click);
        var mmExit = new ToolStripMenuItem("E&xit", null, MMExit_Click);

        // File menu items
        // mmFile.DropDownItems.Add(mmAdd);
        // mmFile.DropDownItems.Add(new ToolStripSeparator());
        // mmFile.DropDownItems.Add(mmReport);
        // mmFile.DropDownItems.Add(new ToolStripSeparator());
        // mmFile.DropDownItems.Add(auditForm);
        // mmFile.DropDownItems.Add(new ToolStripSeparator());
        // mmFile.DropDownItems.Add(refresh);
        mmFile.DropDownItems.Add(new ToolStripSeparator());
        mmFile.DropDownItems.Add(mmExit);
        mainMenu.Items.Clear();

        mainMenu.Items.Add(mmFile);
        mainMenu.Items.Add(mmAdd);
        mainMenu.Items.Add(mmReport);
        mainMenu.Items.Add(refresh);

        // Reports submenu
        mmReport.DropDownItems.Add(reportByAuditor);
        mmReport.DropDownItems.Add(reportByBatch);
        // mainMenu.Items.Add(mmExit);
        this.MainMenuStrip = mainMenu;
        this.Controls.Add(mainMenu);

        // Info Panel (Main working area)
        pnlInfo = new Panel();
        pnlInfo.Size = new Size(this.ClientSize.Width, this.ClientSize.Height - mainMenu.Height);
        pnlInfo.Location = new Point(0, mainMenu.Height);
        pnlInfo.BorderStyle = BorderStyle.None;
        pnlInfo.Visible = true;
        pnlInfo.Dock = DockStyle.Fill;
        pnlInfo.BackColor = Color.FromArgb(230, 230, 230);
        pnlInfo.SendToBack();

        // Login Panel (Hidden - login done in separate form)

        // ===== LEFT SECTION: PROJECT LIST =====
        // Project List Label
        label2 = new Label();
        label2.Text = "Project Code";
        label2.Location = new Point(10, 30);
        label2.AutoSize = true;
        label2.Font = new Font("Arial", 11, FontStyle.Bold);

        // Project List
        lbProjectList = new ListBox();
        lbProjectList.Location = new Point(10, 55);
        lbProjectList.Size = new Size(180, 280);
        lbProjectList.SelectedIndexChanged += LbProjectList_SelectedIndexChanged;
        lbProjectList.BackColor = Color.White;

        // ===== MIDDLE SECTION: BATCH LIST =====
        // Batch List Label
        label1 = new Label();
        label1.Text = "Batch_Name";
        label1.Location = new Point(210, 30);
        label1.AutoSize = true;
        label1.Font = new Font("Arial", 11, FontStyle.Bold);

        // Checkbox for Suspended
        checkBoxSuspended = new CheckBox();
        checkBoxSuspended.Text = "Show Suspended";
        checkBoxSuspended.Location = new Point(385, 30);
        checkBoxSuspended.AutoSize = true;
        checkBoxSuspended.CheckedChanged += CheckBoxSuspended_CheckedChanged;

        // Batch List
        lbBatchList = new ListBox();
        lbBatchList.Location = new Point(210, 55);
        lbBatchList.Size = new Size(320, 280);
        lbBatchList.SelectedIndexChanged += LbBatchList_SelectedIndexChanged;
        lbBatchList.BackColor = Color.White;

        // ===== RIGHT SECTION: USER INFO =====
        // User ID
        Label lblUserID = new Label();
        lblUserID.Text = "User ID :";
        lblUserID.Location = new Point(550, 60);
        lblUserID.AutoSize = true;
        lblUserID.Font = new Font("Arial", 10);

        edtUserId = new TextBox();
        edtUserId.Location = new Point(700, 60);
        edtUserId.Size = new Size(295, 25);
        edtUserId.ReadOnly = true;
        edtUserId.Visible = true;

        edtLineNO = new TextBox();
        edtLineNO.Visible = false;
        edtLineNO.Location = new Point(950, 60);
        edtLineNO.Size = new Size(100, 25);


        // User Name
        lblUsername = new Label();
        lblUsername.Text = "User Name :";
        lblUsername.Location = new Point(550, 95);
        lblUsername.AutoSize = true;
        lblUsername.Font = new Font("Arial", 10);

        edtUsername = new TextBox();
        edtUsername.Location = new Point(700, 95);
        edtUsername.Size = new Size(295, 25);
        edtUsername.ReadOnly = true;
        edtUsername.Text = currentUsername;

        // Batch Name
        lblBatchName = new Label();
        lblBatchName.Text = "Batch Name :";
        lblBatchName.Location = new Point(550, 130);
        lblBatchName.AutoSize = true;
        lblBatchName.Font = new Font("Arial", 10);

        edtBatchName = new TextBox();
        edtBatchName.Location = new Point(700, 130);
        edtBatchName.Size = new Size(295, 25);
        edtBatchName.ReadOnly = true;

        // Sample Count
        lblSampleCount = new Label();
        lblSampleCount.Text = "Sample Count :";
        lblSampleCount.Location = new Point(550, 165);
        lblSampleCount.AutoSize = true;
        lblSampleCount.Font = new Font("Arial", 10);

        edtSampleCount = new TextBox();
        edtSampleCount.Location = new Point(700, 165);
        edtSampleCount.Size = new Size(295, 25);
        edtSampleCount.ReadOnly = false;

        // Time
        lblTime = new Label();
        lblTime.Text = "Time :";
        lblTime.Location = new Point(550, 200);
        lblTime.AutoSize = true;
        lblTime.Font = new Font("Arial", 10);

        edtTime = new TextBox();
        edtTime.Location = new Point(700, 200);
        edtTime.Size = new Size(295, 25);
        edtTime.ReadOnly = true;

        // Comments
        label3 = new Label();
        label3.Text = "Comments :";
        label3.Location = new Point(550, 235);
        label3.AutoSize = true;
        label3.Font = new Font("Arial", 10);

        edtComments = new TextBox();
        edtComments.Location = new Point(700, 235);
        edtComments.Size = new Size(295, 25);
        edtComments.ReadOnly = false;

        // ===== BUTTONS SECTION =====
        btnStart = new Button();
        btnStart.Text = "Start";
        btnStart.Location = new Point(550, 275);
        btnStart.Size = new Size(100, 40);
        btnStart.Font = new Font("Arial", 11, FontStyle.Bold);
        btnStart.BackColor = Color.FromArgb(150, 150, 150);
        btnStart.ForeColor = Color.White;
        btnStart.Click += BtnStart_Click;
        if (btnStart.Enabled)
        {
            btnStart.Cursor = Cursors.Hand;
        }

        btnComplete = new Button();
        btnComplete.Text = "Complete";
        btnComplete.Location = new Point(665, 275);
        btnComplete.Size = new Size(100, 40);
        btnComplete.Font = new Font("Arial", 11, FontStyle.Bold);
        btnComplete.BackColor = Color.FromArgb(150, 150, 150);
        btnComplete.ForeColor = Color.White;
        btnComplete.Enabled = false;
        btnComplete.Click += BtnComplete_Click;
        if (btnComplete.Enabled)
        {
            btnComplete.Cursor = Cursors.Hand;
        }

        btnAuditSuspend = new Button();
        btnAuditSuspend.Text = "Suspend";
        btnAuditSuspend.Location = new Point(780, 275);
        btnAuditSuspend.Size = new Size(100, 40);
        btnAuditSuspend.Font = new Font("Arial", 11, FontStyle.Bold);
        btnAuditSuspend.BackColor = Color.FromArgb(150, 150, 150);
        btnAuditSuspend.ForeColor = Color.White;
        btnAuditSuspend.Enabled = false;
        btnAuditSuspend.Click += BtnAuditSuspend_Click;
        if (btnAuditSuspend.Enabled)
        {
            btnAuditSuspend.Cursor = Cursors.Hand;
        }

        btnAuditCancel = new Button();
        btnAuditCancel.Text = "Cancel";
        btnAuditCancel.Location = new Point(895, 275);
        btnAuditCancel.Size = new Size(100, 40);
        btnAuditCancel.Font = new Font("Arial", 11, FontStyle.Bold);
        btnAuditCancel.BackColor = Color.FromArgb(150, 150, 150);
        btnAuditCancel.ForeColor = Color.White;
        btnAuditCancel.Enabled = false;
        btnAuditCancel.Visible = false;
        btnAuditCancel.Click += BtnAuditCancel_Click;
        if (btnAuditCancel.Enabled)
        {
            btnAuditCancel.Cursor = Cursors.Hand;
        }

        // Hidden fields
        edtEmpType = new TextBox();
        edtEmpType.Text = "emp_id"; // or "emp_id2" based on your logic
        edtEmpType.Visible = false;

        edtPermission = new TextBox();
        edtPermission.Visible = false;

        lblBatchName2 = new Label();
        lblBatchName2.Visible = false;

        // Add all controls to panel
        pnlInfo.Controls.AddRange(new Control[] {
            label2, lbProjectList,
            label1, checkBoxSuspended, lbBatchList,
            lblUserID, edtUserId, lblUsername, edtUsername,
            lblBatchName, edtBatchName, lblSampleCount, edtSampleCount,
            lblTime, edtTime, label3, edtComments,
            btnStart, btnComplete, btnAuditSuspend, btnAuditCancel,
            edtEmpType, edtLineNO, edtPermission, lblBatchName2
        });

        this.Controls.Add(pnlInfo);

        // Load user info and projects on startup
        if (!string.IsNullOrEmpty(currentUsername))
        {
            LoadUserInfo(currentUsername, 0);
        }
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

    private void BtnCancel_Click(object sender, EventArgs e)
    {
        Application.Exit();
    }

    private void FormShow(object sender, EventArgs e)
    {
        // Load initial data
        LoadProjects();
    }

    private void FormClose(object sender, FormClosedEventArgs e)
    {
        Application.Exit();
    }

    // Login is now handled in LoginForm

    private void FormCreate(object? sender, EventArgs e)
    {
        // Initialization logic
    }

    // Menu handlers
    private void MMAdd_Click(object? sender, EventArgs e)
    {
        try
        {
            var addForm = new AddNewForm();
            addForm.StartPosition = FormStartPosition.CenterParent;
            addForm.ShowDialog(this);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening Add New Form: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ReportByAuditor_Click(object? sender, EventArgs e)
    {
        try
        {
            var reportForm = new ReportByAuditorForm();
            reportForm.StartPosition = FormStartPosition.CenterParent;
            reportForm.WindowState = FormWindowState.Maximized;
            reportForm.ShowDialog(this);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening Report By Auditor Form: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ReportByBatch_Click(object? sender, EventArgs e)
    {
        var reportForm = new ReportByBatchForm(this);  // Pass 'this' (MainMenuForm)
        reportForm.Owner = this;
        reportForm.StartPosition = FormStartPosition.CenterParent;
        reportForm.WindowState = FormWindowState.Maximized;
        reportForm.ShowDialog();
    }

    private void Refresh_Click(object? sender, EventArgs e)
    {
        LoadProjects();
    }

    private void MMExit_Click(object? sender, EventArgs e)
    {
        Application.Exit();
    }

    // Button handlers
    private void BtnStart_Click(object? sender, EventArgs e)
    {
        string batchName = edtBatchName?.Text ?? "";
        string userName = edtUsername?.Text ?? "";
        string projectCode = "";
        int userId, line, minToAdd;
        DateTime startTimeMin, endTime;

        // Validate batch name
        if (string.IsNullOrWhiteSpace(edtBatchName?.Text))
        {
            MessageBox.Show("Select batch no", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Validate project selection
        if (lbProjectList?.SelectedIndex < 0)
        {
            MessageBox.Show("Please select a project", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Validate user ID
        if (string.IsNullOrWhiteSpace(edtUserId?.Text) || !int.TryParse(edtUserId?.Text, out userId))
        {
            MessageBox.Show("User ID is not valid. Please login again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Validate line number
        if (string.IsNullOrWhiteSpace(edtLineNO?.Text) || !int.TryParse(edtLineNO?.Text, out line))
        {
            MessageBox.Show("Line number is not valid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Get selected project
        projectCode = lbProjectList?.SelectedItem?.ToString() ?? "";

        // Confirm start
        var result = MessageBox.Show("Are you sure you want to start Batch?", "Confirmation",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (result != DialogResult.Yes)
        {
            return;
        }

        try
        {
            batchName = edtBatchName?.Text ?? "";

            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            // Check current status
            string checkQuery = @"SELECT status, start_datetime2, start_datetime3, end_datetime3 
                             FROM audit_b 
                             WHERE batch_name = @batch_name 
                             AND project_code = @project_code 
                             AND line = @line";

            SqlCommand checkCmd = new SqlCommand(checkQuery, dbConnection);
            checkCmd.Parameters.AddWithValue("@batch_name", batchName);
            checkCmd.Parameters.AddWithValue("@project_code", projectCode);
            checkCmd.Parameters.AddWithValue("@line", line);

            SqlDataReader reader = checkCmd.ExecuteReader();

            if (!reader.HasRows)
            {
                reader.Close();
                MessageBox.Show("Batch not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            reader.Read();

            // Safely read status (should never be NULL)
            int status = 0;
            if (!reader.IsDBNull(reader.GetOrdinal("status")))
            {
                status = reader.GetInt32(reader.GetOrdinal("status"));
            }

            // Check for NULL values
            bool startDateTime2IsNull = reader.IsDBNull(reader.GetOrdinal("start_datetime2"));
            bool startDateTime3IsNull = reader.IsDBNull(reader.GetOrdinal("start_datetime3"));
            bool endDateTime3IsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime3"));

            // Store datetime values only if not NULL
            DateTime? startDateTime3Value = null;
            DateTime? endDateTime3Value = null;

            if (!startDateTime3IsNull)
            {
                startDateTime3Value = reader.GetDateTime(reader.GetOrdinal("start_datetime3"));
            }

            if (!endDateTime3IsNull)
            {
                endDateTime3Value = reader.GetDateTime(reader.GetOrdinal("end_datetime3"));
            }

            reader.Close();

            // Check if batch already started
            if (status == 1)
            {
                MessageBox.Show("Batch Already Started", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoadProjects();
                return;
            }
            else if (status == 2)
            {
                MessageBox.Show("Batch Already Started", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoadProjects();
                return;
            }

            // Get user details
            userName = edtUsername?.Text ?? "";

            // Update UI state
            StartTime = DateTime.Now;
            edtBatchName?.Enabled = true;
            btnStart?.Enabled = false;
            btnComplete?.Enabled = true;
            btnAuditSuspend?.Enabled = true;
            btnAuditCancel?.Enabled = true;
            edtTime?.Text = "";
            lbBatchList?.Enabled = false;
            lbProjectList?.Enabled = false;

            string updateQuery = "";

            // Handle suspended batch (status = -1)
            if (status == -1)
            {
                if (startDateTime2IsNull)
                {
                    // First resume - update start_datetime2
                    updateQuery = @"UPDATE audit_b 
                               SET status = 1, 
                                   username = @UserName, 
                                   " + edtEmpType?.Text + @" = @userid, 
                                   start_dateTime2 = GETDATE()
                               WHERE Batch_Name = @Batch_Name 
                               AND project_code = @Project_code 
                               AND line = @line";
                }
                else if (startDateTime3IsNull)
                {
                    // Second resume - update start_datetime3
                    updateQuery = @"UPDATE audit_b 
                               SET status = 1, 
                                   username = @UserName, 
                                   " + edtEmpType?.Text + @" = @userid, 
                                   start_dateTime3 = GETDATE()
                               WHERE Batch_Name = @Batch_Name 
                               AND project_code = @Project_code 
                               AND line = @line";
                }
                else
                {
                    // Third+ resume - adjust end_datetime2 and update start_datetime3
                    // Only calculate minutes if both values are available
                    if (startDateTime3Value.HasValue && endDateTime3Value.HasValue)
                    {
                        startTimeMin = startDateTime3Value.Value;
                        endTime = endDateTime3Value.Value;

                        TimeSpan difference = endTime - startTimeMin;
                        minToAdd = (int)difference.TotalMinutes;

                        // First update end_datetime2
                        string adjustQuery = @"UPDATE audit_b 
                                          SET end_datetime2 = DATEADD(MINUTE, @minToAdd, end_datetime2)
                                          WHERE Batch_Name = @Batch_Name 
                                          AND project_code = @project_code 
                                          AND line = @line";

                        SqlCommand adjustCmd = new SqlCommand(adjustQuery, dbConnection);
                        adjustCmd.Parameters.AddWithValue("@minToAdd", minToAdd);
                        adjustCmd.Parameters.AddWithValue("@Batch_Name", batchName);
                        adjustCmd.Parameters.AddWithValue("@project_code", projectCode);
                        adjustCmd.Parameters.AddWithValue("@line", line);
                        adjustCmd.ExecuteNonQuery();
                    }

                    // Then update start_datetime3 and clear end_datetime3
                    updateQuery = @"UPDATE audit_b 
                               SET status = 1, 
                                   username = @UserName, 
                                   " + edtEmpType?.Text + @" = @userid, 
                                   start_dateTime3 = GETDATE(),
                                   end_dateTime3 = NULL
                               WHERE Batch_Name = @Batch_Name 
                               AND project_code = @Project_code 
                               AND line = @line";
                }
            }
            else
            {
                // Normal start (status = 0)
                updateQuery = @"UPDATE audit_b 
                           SET status = 1, 
                               UserName = @UserName, 
                               " + edtEmpType?.Text + @" = @userid, 
                               start_dateTime = GETDATE()
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @Project_code 
                           AND line = @line";
            }

            // Execute the update
            SqlCommand updateCmd = new SqlCommand(updateQuery, dbConnection);
            updateCmd.Parameters.AddWithValue("@UserName", userName);
            updateCmd.Parameters.AddWithValue("@Batch_Name", batchName);
            updateCmd.Parameters.AddWithValue("@UserID", userId);
            updateCmd.Parameters.AddWithValue("@Project_code", projectCode);
            updateCmd.Parameters.AddWithValue("@line", line);
            updateCmd.ExecuteNonQuery();

            // Check and update project queue count
            string queueCheckQuery = @"SELECT que_count 
                                  FROM audit_b_projects 
                                  WHERE project_code = @project_code";

            SqlCommand queueCheckCmd = new SqlCommand(queueCheckQuery, dbConnection);
            queueCheckCmd.Parameters.AddWithValue("@project_code", projectCode);

            object queueResult = queueCheckCmd.ExecuteScalar();

            if (queueResult != null && queueResult != DBNull.Value && Convert.ToInt32(queueResult) > 0)
            {
                string queueUpdateQuery = @"UPDATE audit_b_projects 
                                       SET que_count = que_count - 1 
                                       WHERE project_code = @project_code";

                SqlCommand queueUpdateCmd = new SqlCommand(queueUpdateQuery, dbConnection);
                queueUpdateCmd.Parameters.AddWithValue("@Project_code", projectCode);
                queueUpdateCmd.ExecuteNonQuery();
            }

            MessageBox.Show("Batch started successfully!", "Success",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error starting batch: {ex.Message}\n\nStack Trace: {ex.StackTrace}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

            // Reset UI state on error
            btnStart?.Enabled = true;
            btnComplete?.Enabled = false;
            btnAuditSuspend?.Enabled = false;
            btnAuditCancel?.Enabled = false;
            lbBatchList?.Enabled = true;
            lbProjectList?.Enabled = true;
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }

    // Add this field to your class if not already present
    private DateTime StartTime;

    private void BtnComplete_Click(object? sender, EventArgs e)
    {
        string batchName, sampleCount, projectCode, userName;
        int line, userId;
        DateTime endTime;

        if (string.IsNullOrWhiteSpace(edtSampleCount?.Text))
        {
            MessageBox.Show("Please add sample count.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        endTime = DateTime.Now;
        TimeSpan timeDiff = endTime - StartTime;
        edtTime?.Text = $"{(int)timeDiff.TotalHours:D2}:{timeDiff.Minutes:D2}:{timeDiff.Seconds:D2}";

        var result = MessageBox.Show("Are you sure you want to Submit?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result != DialogResult.Yes)
            return;

        try
        {
            batchName = edtBatchName?.Text ?? "";
            sampleCount = edtSampleCount?.Text ?? "";
            projectCode = lbProjectList?.SelectedItem?.ToString() ?? "";
            line = int.Parse(edtLineNO?.Text ?? "");
            userName = edtUsername?.Text ?? "";
            userId = int.Parse(edtUserId?.Text ?? "");

            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            // Check current end datetime status
            string checkQuery = @"SELECT end_datetime, end_datetime2, end_datetime3 
                                FROM audit_b 
                                WHERE batch_name = @batch_name 
                                AND project_code = @project_code 
                                AND line = @line";
            SqlCommand checkCmd = new SqlCommand(checkQuery, dbConnection);
            checkCmd.Parameters.AddWithValue("@batch_name", batchName);
            checkCmd.Parameters.AddWithValue("@project_code", projectCode);
            checkCmd.Parameters.AddWithValue("@line", line);

            SqlDataReader reader = checkCmd.ExecuteReader();
            if (!reader.HasRows)
            {
                reader.Close();
                MessageBox.Show("Batch not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            reader.Read();
            bool endDateTimeIsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime"));
            bool endDateTime2IsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime2"));
            bool endDateTime3IsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime3"));
            reader.Close();

            string updateQuery;

            if (endDateTimeIsNull)
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = 2, comments = null, end_datetime = GETDATE(), 
                               username = @UserName, " + edtEmpType?.Text + @" = @userid, 
                               sample_count = @sample_count 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }
            else if (endDateTime2IsNull)
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = 2, comments = null, end_datetime2 = GETDATE(), 
                               username = @UserName, " + edtEmpType?.Text + @" = @userid, 
                               sample_count = @sample_count 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }
            else
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = 2, comments = null, end_datetime3 = GETDATE(), 
                               username = @UserName, " + edtEmpType?.Text + @" = @userid, 
                               sample_count = @sample_count 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }

            SqlCommand updateCmd = new SqlCommand(updateQuery, dbConnection);
            updateCmd.Parameters.AddWithValue("@Batch_Name", batchName);
            updateCmd.Parameters.AddWithValue("@Sample_Count", sampleCount);
            updateCmd.Parameters.AddWithValue("@project_code", projectCode);
            updateCmd.Parameters.AddWithValue("@line", line);
            updateCmd.Parameters.AddWithValue("@UserName", userName);
            updateCmd.Parameters.AddWithValue("@UserID", userId);
            updateCmd.ExecuteNonQuery();

            // Reset UI state
            btnStart?.Enabled = true;
            btnComplete?.Enabled = false;
            btnAuditSuspend?.Enabled = false;
            btnAuditCancel?.Enabled = false;

            edtBatchName?.Text = "";
            edtSampleCount?.Text = "";
            edtTime?.Text = "";

            edtBatchName?.Enabled = false;
            lbBatchList?.Enabled = true;
            lbProjectList?.Enabled = true;

            lbBatchList?.Items.Clear();

            MessageBox.Show("Batch completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error completing batch: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }

    private void BtnAuditSuspend_Click(object? sender, EventArgs e)
    {
        string batchName, projectCode, comments;
        int line;

        if (string.IsNullOrWhiteSpace(edtComments?.Text))
        {
            MessageBox.Show("Please enter a reason for suspending the batch in the Comments field.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            edtComments?.Focus();
            return;
        }

        var result = MessageBox.Show("Are you sure you want to Suspend?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result != DialogResult.Yes)
            return;

        try
        {
            batchName = edtBatchName?.Text ?? "";
            comments = edtComments?.Text ?? "";
            projectCode = lbProjectList?.SelectedItem?.ToString() ?? "";
            line = int.Parse(edtLineNO?.Text ?? "");

            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            // Check current end datetime status
            string checkQuery = @"SELECT end_datetime, end_datetime2, start_datetime3, end_datetime3 
                                FROM audit_b 
                                WHERE batch_name = @batch_name 
                                AND project_code = @project_code 
                                AND line = @line";
            SqlCommand checkCmd = new SqlCommand(checkQuery, dbConnection);
            checkCmd.Parameters.AddWithValue("@batch_name", batchName);
            checkCmd.Parameters.AddWithValue("@project_code", projectCode);
            checkCmd.Parameters.AddWithValue("@line", line);

            SqlDataReader reader = checkCmd.ExecuteReader();
            if (!reader.HasRows)
            {
                reader.Close();
                MessageBox.Show("Batch not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            reader.Read();
            bool endDateTimeIsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime"));
            bool endDateTime2IsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime2"));
            bool endDateTime3IsNull = reader.IsDBNull(reader.GetOrdinal("end_datetime3"));
            reader.Close();

            string updateQuery;

            if (endDateTimeIsNull)
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = -1, comments = @comments, end_datetime = GETDATE() 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }
            else if (endDateTime2IsNull)
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = -1, comments = @comments, end_datetime2 = GETDATE() 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }
            else if (endDateTime3IsNull)
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = -1, comments = @comments, end_datetime3 = GETDATE() 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }
            else
            {
                updateQuery = @"UPDATE audit_b 
                           SET status = -1, comments = @comments 
                           WHERE Batch_Name = @Batch_Name 
                           AND project_code = @project_code 
                           AND line = @line";
            }

            SqlCommand updateCmd = new SqlCommand(updateQuery, dbConnection);
            updateCmd.Parameters.AddWithValue("@Batch_Name", batchName);
            updateCmd.Parameters.AddWithValue("@project_code", projectCode);
            updateCmd.Parameters.AddWithValue("@line", line);
            updateCmd.Parameters.AddWithValue("@comments", comments);
            updateCmd.ExecuteNonQuery();

            // Reset UI state
            btnAuditCancel?.Enabled = false;
            btnAuditSuspend?.Enabled = false;
            btnComplete?.Enabled = false;
            btnStart?.Enabled = true;

            edtBatchName?.Text = "";
            edtBatchName?.Enabled = false;
            edtComments?.Text = "";

            lbBatchList?.Enabled = true;
            lbProjectList?.Enabled = true;

            // Reload batches
            if (lbProjectList?.SelectedIndex >= 0)
            {
                string selectedProject = lbProjectList.SelectedItem?.ToString() ?? "";
                LoadBatchesForProject(selectedProject);
            }

            MessageBox.Show("Batch suspended successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error suspending batch: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }

    private void BtnAuditCancel_Click(object? sender, EventArgs e)
    {
        string batchName, projectCode;
        int line;

        var result = MessageBox.Show("Are you sure you want to Cancel?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result != DialogResult.Yes)
            return;

        try
        {
            batchName = edtBatchName?.Text ?? "";
            projectCode = lbProjectList?.SelectedItem?.ToString() ?? "";
            line = int.Parse(edtLineNO?.Text ?? "");

            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            // Insert into Audit_B_Cancel table
            string insertQuery = @"INSERT INTO Audit_B_Cancel 
                                (Project_code, Batch_Name, Sample_Count, Status, UserName, Operatorid, Emp_id, 
                                 Start_Datetime_, End_Datetime, Start_dateTime2, End_dateTime2, 
                                 start_datetime3, end_datetime3, Cancel_Datetime, line) 
                                SELECT 
                                    Project_code, 
                                    Batch_Name, 
                                    Sample_Count, 
                                    Status, 
                                    UserName, 
                                    operatorid, 
                                    emp_id, 
                                    Start_Datetime, 
                                    End_Datetime, 
                                    Start_dateTime2, 
                                    End_dateTime2, 
                                    Start_dateTime3, 
                                    End_dateTime3, 
                                    GETDATE() AS Cancel_Datetime, 
                                    line 
                                FROM Audit_B 
                                WHERE Batch_Name = @Batch_Name 
                                AND Project_code = @Project_code 
                                AND line = @line";
            SqlCommand insertCmd = new SqlCommand(insertQuery, dbConnection);
            insertCmd.Parameters.AddWithValue("@Batch_Name", batchName);
            insertCmd.Parameters.AddWithValue("@Project_code", projectCode);
            insertCmd.Parameters.AddWithValue("@line", line);
            insertCmd.ExecuteNonQuery();

            // Update audit_b to set status = -2 and clear all datetimes
            string updateQuery = @"UPDATE audit_b 
                               SET status = -2, 
                                   start_datetime = NULL, 
                                   end_datetime = NULL, 
                                   start_datetime2 = NULL, 
                                   end_datetime2 = NULL, 
                                   start_datetime3 = NULL, 
                                   end_datetime3 = NULL 
                               WHERE Batch_Name = @Batch_Name 
                               AND project_code = @project_code 
                               AND line = @line";
            SqlCommand updateCmd = new SqlCommand(updateQuery, dbConnection);
            updateCmd.Parameters.AddWithValue("@Batch_Name", batchName);
            updateCmd.Parameters.AddWithValue("@project_code", projectCode);
            updateCmd.Parameters.AddWithValue("@line", line);
            updateCmd.ExecuteNonQuery();

            // Reset UI state
            btnAuditCancel?.Enabled = false;
            btnAuditSuspend?.Enabled = false;
            btnComplete?.Enabled = false;
            btnStart?.Enabled = true;

            edtBatchName?.Text = "";
            edtBatchName?.Enabled = false;
            edtComments?.Text = "";

            lbBatchList?.Enabled = true;
            lbProjectList?.Enabled = true;

            // Reload batches
            if (lbProjectList?.SelectedIndex >= 0)
            {
                string selectedProject = lbProjectList?.SelectedItem?.ToString() ?? "";
                LoadBatchesForProject(selectedProject);
            }

            MessageBox.Show("Batch cancelled successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error cancelling batch: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }

    // Helper methods
    private void SetMenuVisibility(int auditB)
    {
        // Menu visibility is now managed by the menu structure
    }

    private void LoadUserInfo(string username, int auditB)
    {
        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string query = "SELECT EMP_NAME, emp_id FROM EMPLOYEE_INFO WHERE USER_NAME = @username";
            SqlCommand cmd = new SqlCommand(query, dbConnection);
            cmd.Parameters.AddWithValue("@username", username);

            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                reader.Read();
                edtUsername?.Text = reader["EMP_NAME"].ToString() ?? "";
                edtUserId?.Text = reader["emp_id"].ToString() ?? "notfound";
            }
            reader.Close();

            LoadProjects();
            LoadCurrentBatch(int.Parse(edtUserId?.Text ?? "0"));

            // ADD THIS LINE - Check permissions after loading user info
            CheckUserPermissions(username);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading user info: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }
    private void CheckUserPermissions(string username)
    {
        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();
            string query = "SELECT user_Password, Audit_B FROM Employee_Info WHERE User_name = @Username";
            SqlCommand cmd = new SqlCommand(query, dbConnection);
            cmd.Parameters.AddWithValue("@Username", username);

            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                reader.Read();
                int auditB = reader.IsDBNull(reader.GetOrdinal("Audit_B")) ? 0 : reader.GetInt32(reader.GetOrdinal("Audit_B"));
                reader.Close();

                // Set visibility based on Audit_B value
                if (auditB == 1)
                {
                    mmReport?.Visible = false;
                    mmAdd?.Visible = true;
                    edtPermission?.Text = "Yes";
                }
                else if (auditB == 2)
                {
                    mmReport?.Visible = true;
                    mmAdd?.Visible = false;
                    edtPermission?.Text = "No";
                }
                else if (auditB == 3)
                {
                    mmReport?.Visible = true;
                    mmAdd?.Visible = true;
                    edtPermission?.Text = "Yes";
                }
                else
                {
                    mmReport?.Visible = false;
                    mmAdd?.Visible = false;
                    edtPermission?.Text = "No";
                }
            }
            else
            {
                reader.Close();
                // Default: hide everything if user not found
                mmReport?.Visible = false;
                mmAdd?.Visible = false;
                edtPermission?.Text = "No";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error checking user permissions: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }
    private void LoadUserData()
    {
        // User data loading is done in ValidateLogin
    }

    private void LoadProjects()
    {
        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string query = "SELECT project_code FROM audit_b_projects WHERE status = -1 ORDER BY status DESC, deadline";
            SqlCommand cmd = new SqlCommand(query, dbConnection);
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            lbProjectList?.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                lbProjectList?.Items.Add(row["project_code"].ToString() ?? "");
            }

            if (lbProjectList?.Items.Count > 0)
            {
                lbProjectList?.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading projects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection.Close();
        }
    }

    private void LoadBatchesForProject(string projectCode)
    {
        if (string.IsNullOrWhiteSpace(projectCode))
            return;

        int userId;

        // Validate user ID
        if (string.IsNullOrWhiteSpace(edtUserId?.Text) || !int.TryParse(edtUserId?.Text, out userId))
        {
            MessageBox.Show("User ID is not valid. Please login again.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string query;

            if (checkBoxSuspended!.Checked)
            {
                // Show only suspended batches for current user
                query = @"SELECT project_code, batch_name, line 
                     FROM audit_b 
                     WHERE project_code = @project_code 
                     AND (status = -1 AND " + edtEmpType?.Text + @" = @userid)";
            }
            else
            {
                // Show all available batches with priority ordering
                query = @"SELECT project_code, batch_name, line 
                     FROM audit_b 
                     WHERE project_code = @project_code 
                     AND (
                         (status IN (0, -2))
                         OR
                         (status IN ( 1) AND " + edtEmpType?.Text + @" = @userid)
                     )
                     ORDER BY 
                         CASE status 
                             WHEN 1 THEN 1
                             WHEN 0 THEN 2
                             WHEN -2 THEN 3
                             WHEN -1 THEN 4
                             ELSE 5
                         END, batch_name";
            }

            SqlCommand cmd = new SqlCommand(query, dbConnection);
            cmd.Parameters.AddWithValue("@project_code", projectCode);
            cmd.Parameters.AddWithValue("@userid", userId);

            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            // Update label with count
            label1?.Text = $"Batch_Name ({dt.Rows.Count})";

            // Clear and populate batch list
            lbBatchList?.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                lbBatchList?.Items.Add(row["batch_name"].ToString() ?? "");
            }

            // Clear batch name field
            edtBatchName?.Text = "";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading batches: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }
    private void LoadCurrentBatch(int userId)
    {
        try
        {
            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            string query = @"SELECT b.project_code, b.status, b.batch_name FROM audit_B b 
                           JOIN employee_info e ON e.emp_id = b.emp_id 
                           WHERE b.emp_id = @userid AND b.status = 1";
            SqlCommand cmd = new SqlCommand(query, dbConnection);
            cmd.Parameters.AddWithValue("@userid", userId);

            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                reader.Read();
                string projectCode = reader["project_code"].ToString() ?? "";
                string batchName = reader["batch_name"].ToString() ?? "";

                int? projectIndex = lbProjectList?.Items.IndexOf(projectCode);
                if (projectIndex >= 0)
                {
                    lbProjectList?.SelectedIndex = (int)projectIndex;
                }
    
                btnStart?.Enabled = false;
                btnAuditSuspend?.Enabled = true;
                btnComplete?.Enabled = true;
                lbProjectList?.Enabled = false;
                lbBatchList?.Enabled = false;
            }
            reader.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading current batch: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    private void LbProjectList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (lbProjectList?.SelectedIndex >= 0)
        {
            string? selectedProject = lbProjectList?.SelectedItem?.ToString() ?? "";
            LoadBatchesForProject(selectedProject);
        }
    }
    private void LbBatchList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (lbBatchList?.SelectedIndex == -1)
            return;

        string batchName;
        int userId;

        try
        {
            // Check if suspended checkbox is not checked
            if (!checkBoxSuspended!.Checked)
            {
                // Force selection to first item if not already
                if (lbBatchList?.SelectedIndex != 0)
                {
                    lbBatchList?.SelectedIndex = 0;
                    return; // Event will fire again with correct index
                }
            }

            batchName = lbBatchList?.SelectedItem?.ToString() ?? "";

            // Validate user ID
            if (string.IsNullOrWhiteSpace(edtUserId?.Text) || !int.TryParse(edtUserId.Text, out userId))
            {
                MessageBox.Show("User ID is not valid. Please login again.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dbConnection?.State != ConnectionState.Open)
                dbConnection?.Open();

            // Query to get batch details
            string query = @"SELECT project_code, batch_name, line, comments 
                        FROM audit_b 
                        WHERE batch_name = @batch_name 
                        AND (
                            (status IN (0, -2))
                            OR
                            (status IN (-1, 1) AND " + edtEmpType?.Text + @" = @userid)
                        )";

            SqlCommand cmd = new SqlCommand(query, dbConnection);
            cmd.Parameters.AddWithValue("@Batch_name", batchName);
            cmd.Parameters.AddWithValue("@userid", userId);

            SqlDataReader reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {
                reader.Read();

                // Update UI fields
                edtBatchName?.Text = batchName;
                edtComments?.Text = reader["comments"].ToString() ?? "";
                edtLineNO?.Text = reader["line"].ToString() ?? "";

                // You can also store project_code if needed
                // string projectCode = reader["project_code"].ToString() ?? "";
            }
            else
            {
                // Clear fields if no matching record found
                edtBatchName?.Text = "";
                edtComments?.Text = "";
                edtLineNO?.Text = "";
            }

            reader.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading batch details: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (dbConnection?.State == ConnectionState.Open)
                dbConnection?.Close();
        }
    }

    private void MainMenuFormClosing(object? sender, FormClosingEventArgs e)
    {
        Application.Exit();
    }

    private void CheckBoxSuspended_CheckedChanged(object? sender, EventArgs e)
    {
        edtComments?.Text = "";
        // Exit if batch list is empty
        if (lbBatchList?.Items.Count == 0)
        {
            return;
        }

        // Reload batches based on checkbox state
        if (lbProjectList?.SelectedIndex >= 0)
        {
            string selectedProject = lbProjectList.SelectedItem?.ToString() ?? "";
            LoadBatchesForProject(selectedProject);
        }
    }
}