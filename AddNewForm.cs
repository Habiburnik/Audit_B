using System.Data;
using Microsoft.Data.SqlClient;
namespace Audit_B;

public partial class AddNewForm : Form
{
    private SqlConnection? dbConnection;

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

    // Right: Add Batches
    private GroupBox? grpAddBatches;
    private RadioButton? rb1, rb2, rb3, rb4, rb5;
    private Label? lblSelectProjectCode;
    private ComboBox? cbProjectCode;
    private TextBox? txtBatchList; // multiline for batches
    private Button? btnAddBatches;

    public AddNewForm()
    {
        InitializeComponent();
        InitializeDatabase();
    }

    private void InitializeComponent()
    {
        this.Text = "Add Projects / Batches";
        this.Size = new Size(900, 450);
        this.StartPosition = FormStartPosition.CenterParent;

        // Group Add Project (left)
        grpAddProject = new GroupBox();
        grpAddProject.Text = "Add Projects";
        grpAddProject.Location = new Point(10, 10);
        grpAddProject.Size = new Size(420, 380);

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

        // Group Add Batches (right)
        grpAddBatches = new GroupBox();
        grpAddBatches.Text = "Add Batches";
        grpAddBatches.Location = new Point(440, 10);
        grpAddBatches.Size = new Size(420, 380);

        rb1 = new RadioButton() { Text = "1st", Location = new Point(15, 20), AutoSize = true };
        rb2 = new RadioButton() { Text = "2nd", Location = new Point(70, 20), AutoSize = true };
        rb3 = new RadioButton() { Text = "3rd", Location = new Point(125, 20), AutoSize = true };
        rb4 = new RadioButton() { Text = "4th", Location = new Point(180, 20), AutoSize = true };
        rb5 = new RadioButton() { Text = "5th", Location = new Point(235, 20), AutoSize = true };
        rb1.Checked = true;

        lblSelectProjectCode = new Label() { Text = "Select_Project_Code :", Location = new Point(15, 50), AutoSize = true };
        cbProjectCode = new ComboBox() { Location = new Point(160, 47), Size = new Size(240, 24), DropDownStyle = ComboBoxStyle.DropDownList };

        txtBatchList = new TextBox() { Location = new Point(15, 85), Size = new Size(385, 220), Multiline = true, ScrollBars = ScrollBars.Vertical }; 

        btnAddBatches = new Button() { Text = "Add Batches", Location = new Point(130, 320), Size = new Size(120, 30) };
        btnAddBatches.Click += BtnAddBatches_Click;


        grpAddBatches.Controls.AddRange(new Control[] {
            rb1, rb2, rb3, rb4, rb5,
            lblSelectProjectCode, cbProjectCode, txtBatchList, btnAddBatches
        });

        this.Controls.AddRange(new Control[] { grpAddProject, grpAddBatches });

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
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            MessageBox.Show("Batch Added Successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            txtBatchList.Clear();
            cbProjectCode?.SelectedIndex = -1;
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

            // Check if project exists
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

            // Get next project_id
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