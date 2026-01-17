using System.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Diagnostics;


namespace Audit_B
{
    public partial class ReportByAuditorForm : Form
    {
        private SqlConnection? dbConnection;
        private SqlDataAdapter? dataAdapter;
        private DataTable? dataTable;

        // Top Panel Controls
        private Panel? pnlTop;
        private Label? lblProjectCode;
        private ComboBox? cbProjectList;
        private CheckBox? checkBoxShowComplete;
        private Label? lblFrom;
        private DateTimePicker? dtpFromDate;
        private Label? lblTo;
        private DateTimePicker? dtpToDate;
        private Label? lblAuditorList;
        private TextBox? txtAuditorList;
        private Button? btnReportByProject;
        private Button? btnReportByDate;
        private Button? btnReportByAuditor;
        private Button? btnAuditorReport;
        private Button? btnAttendance;

        // DataGridView
        private DataGridView? dataGridView;

        // Bottom Panel Controls
        private Panel? pnlBottom;
        private Label? lblTotalAuditor;
        private Label? lblBatchComplete;
        private Label? lblBatchRunning;
        private Label? lblBatchPending;
        private Label? lblTotalBatch;
        private Label? lblTotalSample;

        // Find Panel
        private Panel? pnlFind;
        private TextBox? txtFindText;
        private Button? btnSearch;

        // Hidden memo for processing
        private TextBox? txtAuditorList2;

        public ReportByAuditorForm()
        {
            InitializeDatabase();
            InitializeComponent();
            LoadProjects();

            ContextMenuStrip cms = new ContextMenuStrip();
            ToolStripMenuItem exportItem = new ToolStripMenuItem("Export to Excel");
            dataGridView?.ContextMenuStrip = cms!;
            exportItem.Click += ExportToExcel_Click;

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
                MessageBox.Show($"Database initialization error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void InitializeComponent()
        {
            this.Text = "Report By Auditor";
            this.Size = new Size(1400, 800);
            this.StartPosition = FormStartPosition.CenterParent;
            // this.BackColor = Color.FromArgb(139, 175, 185);
            this.KeyPreview = true;
            this.KeyDown += Form_KeyDown;
            // this.Icon = new Icon("AuditB_icon.ico");

            // ===== TOP PANEL =====
            pnlTop = new Panel();
            pnlTop.Dock = DockStyle.Top;
            pnlTop.Height = 150;
            // pnlTop.BackColor = Color.FromArgb(139, 175, 185);
            pnlTop.Padding = new Padding(10);

            // Project Code Label
            lblProjectCode = new Label();
            lblProjectCode.Text = "Project_Code";
            lblProjectCode.Location = new Point(10, 10);
            lblProjectCode.AutoSize = true;
            lblProjectCode.Font = new Font("Arial", 10, FontStyle.Regular);

            // Project ComboBox
            cbProjectList = new ComboBox();
            cbProjectList.Location = new Point(10, 35);
            cbProjectList.Size = new Size(180, 25);
            cbProjectList.DropDownStyle = ComboBoxStyle.DropDownList;

            // Show Complete Checkbox
            checkBoxShowComplete = new CheckBox();
            checkBoxShowComplete.Text = "Show Complete Project";
            checkBoxShowComplete.Location = new Point(10, 65);
            checkBoxShowComplete.AutoSize = true;
            checkBoxShowComplete.CheckedChanged += CheckBoxShowComplete_CheckedChanged;

            // From Date Label
            lblFrom = new Label();
            lblFrom.Text = "From";
            lblFrom.Location = new Point(10, 95);
            lblFrom.AutoSize = true;
            lblFrom.Font = new Font("Arial", 10, FontStyle.Regular);

            // From DateTimePicker
            dtpFromDate = new DateTimePicker();
            dtpFromDate.Location = new Point(10, 120);
            dtpFromDate.Size = new Size(180, 25);
            dtpFromDate.Format = DateTimePickerFormat.Short;
            dtpFromDate.Value = DateTime.Now;

            // To Label
            lblTo = new Label();
            lblTo.Text = "To";
            lblTo.Location = new Point(200, 95);
            lblTo.AutoSize = true;
            lblTo.Font = new Font("Arial", 10, FontStyle.Regular);

            // To DateTimePicker
            dtpToDate = new DateTimePicker();
            dtpToDate.Location = new Point(200, 120);
            dtpToDate.Size = new Size(180, 25);
            dtpToDate.Format = DateTimePickerFormat.Short;
            dtpToDate.Value = DateTime.Now;

            // Auditor List Label
            lblAuditorList = new Label();
            lblAuditorList.Text = "Auditor_List";
            lblAuditorList.Location = new Point(400, 10);
            lblAuditorList.AutoSize = true;
            lblAuditorList.Font = new Font("Arial", 10, FontStyle.Regular);

            // Auditor List TextBox (Multiline)
            txtAuditorList = new TextBox();
            txtAuditorList.Location = new Point(400, 35);
            txtAuditorList.Size = new Size(240, 110);
            txtAuditorList.Multiline = true;
            txtAuditorList.ScrollBars = ScrollBars.Vertical;
            txtAuditorList.Font = new Font("Arial", 9);

            // Hidden Auditor List 2
            txtAuditorList2 = new TextBox();
            txtAuditorList2.Visible = false;
            txtAuditorList2.Multiline = true;

            // Buttons
            btnReportByProject = new Button();
            btnReportByProject.Text = "Report By Project";
            btnReportByProject.Location = new Point(660, 10);
            btnReportByProject.Size = new Size(140, 30);
            btnReportByProject.BackColor = Color.White;
            btnReportByProject.Click += BtnReportByProject_Click;

            btnReportByDate = new Button();
            btnReportByDate.Text = "Report By Date";
            btnReportByDate.Location = new Point(660, 45);
            btnReportByDate.Size = new Size(140, 30);
            btnReportByDate.BackColor = Color.White;
            btnReportByDate.Click += BtnReportByDate_Click;

            btnReportByAuditor = new Button();
            btnReportByAuditor.Text = "Report By Auditor";
            btnReportByAuditor.Location = new Point(660, 80);
            btnReportByAuditor.Size = new Size(140, 30);
            btnReportByAuditor.BackColor = Color.White;
            btnReportByAuditor.Click += BtnReportByAuditor_Click;

            btnAuditorReport = new Button();
            btnAuditorReport.Text = "Auditor_Report";
            btnAuditorReport.Location = new Point(660, 115);
            btnAuditorReport.Size = new Size(140, 30);
            btnAuditorReport.BackColor = Color.White;
            btnAuditorReport.Click += BtnAuditorReport_Click;

            btnAttendance = new Button();
            btnAttendance.Text = "Attendance";
            btnAttendance.Location = new Point(810, 10);
            btnAttendance.Size = new Size(140, 30);
            btnAttendance.BackColor = Color.White;
            btnAttendance.Click += BtnAttendance_Click;

            // Add controls to top panel
            pnlTop.Controls.AddRange(new Control[] {
                lblProjectCode, cbProjectList, checkBoxShowComplete,
                lblFrom, dtpFromDate, lblTo, dtpToDate,
                lblAuditorList, txtAuditorList, txtAuditorList2,
                btnReportByProject, btnReportByDate, btnReportByAuditor,
                btnAuditorReport, btnAttendance
            });

            // ===== DATAGRIDVIEW =====
            dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.ReadOnly = true;
            dataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView.BackgroundColor = Color.White;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.DoubleClick += DataGridView_DoubleClick;

            // ===== BOTTOM PANEL =====
            pnlBottom = new Panel();
            pnlBottom.Dock = DockStyle.Bottom;
            pnlBottom.Height = 50;
            // pnlBottom.BackColor = Color.FromArgb(139, 175, 185);
            pnlBottom.Padding = new Padding(10);

            lblTotalAuditor = new Label();
            lblTotalAuditor.Text = "Total Auditor :";
            lblTotalAuditor.Location = new Point(10, 15);
            lblTotalAuditor.AutoSize = true;
            lblTotalAuditor.Font = new Font("Arial", 10, FontStyle.Bold);

            lblBatchComplete = new Label();
            lblBatchComplete.Text = "Batch Complete :";
            lblBatchComplete.Location = new Point(200, 15);
            lblBatchComplete.AutoSize = true;
            lblBatchComplete.Font = new Font("Arial", 10, FontStyle.Bold);

            lblBatchRunning = new Label();
            lblBatchRunning.Text = "Batch Running :";
            lblBatchRunning.Location = new Point(390, 15);
            lblBatchRunning.AutoSize = true;
            lblBatchRunning.Font = new Font("Arial", 10, FontStyle.Bold);

            lblBatchPending = new Label();
            lblBatchPending.Text = "Batch Pending :";
            lblBatchPending.Location = new Point(580, 15);
            lblBatchPending.AutoSize = true;
            lblBatchPending.Font = new Font("Arial", 10, FontStyle.Bold);

            lblTotalBatch = new Label();
            lblTotalBatch.Text = "Total Batch :";
            lblTotalBatch.Location = new Point(770, 15);
            lblTotalBatch.AutoSize = true;
            lblTotalBatch.Font = new Font("Arial", 10, FontStyle.Bold);

            lblTotalSample = new Label();
            lblTotalSample.Text = "Total Sample :";
            lblTotalSample.Location = new Point(960, 15);
            lblTotalSample.AutoSize = true;
            lblTotalSample.Font = new Font("Arial", 10, FontStyle.Bold);

            pnlBottom.Controls.AddRange(new Control[] {
                lblTotalAuditor, lblBatchComplete, lblBatchRunning,
                lblBatchPending, lblTotalBatch, lblTotalSample
            });

            // ===== FIND PANEL (Hidden by default) =====
            pnlFind = new Panel();
            pnlFind.Size = new Size(300, 50);
            pnlFind.Location = new Point(this.ClientSize.Width - 320, pnlTop.Height + 10);
            // pnlFind.BackColor = Color.FromArgb(139, 175, 185);
            pnlFind.BorderStyle = BorderStyle.FixedSingle;
            pnlFind.Visible = false;

            txtFindText = new TextBox();
            txtFindText.Location = new Point(10, 12);
            txtFindText.Size = new Size(200, 25);

            btnSearch = new Button();
            btnSearch.Text = "Search";
            btnSearch.Location = new Point(220, 10);
            btnSearch.Size = new Size(70, 25);
            btnSearch.Click += BtnSearch_Click;

            pnlFind.Controls.AddRange(new Control[] { txtFindText, btnSearch });

            // Add all main controls to form
            this.Controls.Add(dataGridView);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);
            this.Controls.Add(pnlFind);

            // Bring find panel to front
            pnlFind.BringToFront();
        }

        private void LoadProjects()
        {
            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                cbProjectList?.Items.Clear();
                cbProjectList?.Items.Add("All Project");
                string query = @"SELECT project_code FROM audit_b_projects 
                                WHERE status IN (0, -1, -2)
                                ORDER BY CASE 
                                    WHEN Status = -1 THEN 1
                                    WHEN Status = 0 THEN 2
                                    WHEN Status = -2 THEN 3
                                    WHEN Status = 1 THEN 4
                                END, Deadline";

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    cbProjectList?.Items.Add(reader["project_code"].ToString() ?? "");
                }
                reader.Close();

                if (cbProjectList?.Items.Count > 0)
                    cbProjectList.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading projects: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection.Close();
            }
        }

        private void CheckBoxShowComplete_CheckedChanged(object? sender, EventArgs e)
        {
            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                cbProjectList?.Items.Clear();
                cbProjectList?.Items.Add("All Project");
                string query;
                if (checkBoxShowComplete!.Checked)
                {
                    query = "SELECT DISTINCT project_code FROM audit_b_projects WHERE status IN (1)";
                }
                else
                {
                    query = "SELECT DISTINCT project_code FROM audit_b_projects WHERE status IN (0, -1, -2)";
                }

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    cbProjectList?.Items.Add(reader["project_code"].ToString() ?? "");
                }
                reader.Close();

                if (cbProjectList?.Items.Count > 0)
                    cbProjectList.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading projects: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection.Close();
            }
        }

        private string ProcessAuditorList()
        {
            txtAuditorList2?.Clear();
            txtAuditorList2?.Lines = new string[] { "(" };

            var lines = txtAuditorList?.Lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();

            if (lines != null)
            {
                foreach (var line in lines)
                {
                    txtAuditorList2?.AppendText($"'{line.Trim()}',\r\n");
                }
            }

            txtAuditorList2?.AppendText(")");

            string result = txtAuditorList2?.Text.Replace(",\r\n)", "\r\n)") ?? "";
            return result;
        }

        private void CalculateSummary()
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
                return;

            double totalSample = 0;
            int totalBatch = 0;
            int completeCount = 0;
            int runningCount = 0;
            HashSet<int> auditorSet = new HashSet<int>();

            foreach (DataRow row in dataTable.Rows)
            {
                if (row["Sample_Count"] != DBNull.Value)
                    totalSample += Convert.ToDouble(row["Sample_Count"]);

                totalBatch++;

                if (row["Auditor_id"] != DBNull.Value)
                    auditorSet.Add(Convert.ToInt32(row["Auditor_id"]));

                if (row["Status"] != DBNull.Value)
                {
                    string? status = row["Status"].ToString();
                    if (status == "Complete")
                        completeCount++;
                    else if (status == "Running")
                        runningCount++;
                }
            }

            lblTotalSample?.Text = $"Total Sample : {totalSample}";
            lblTotalBatch?.Text = $"Total Batch : {totalBatch}";
            lblTotalAuditor?.Text = $"Total Auditors : {auditorSet.Count}";
            lblBatchComplete?.Text = $"Batch Complete : {completeCount}";
            lblBatchRunning?.Text = $"Batch Running : {runningCount}";
            lblBatchPending?.Text = $"Batch Pending : {totalBatch - completeCount - runningCount}";
        }

        private void CalculateSummaryForAuditor()
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
                return;

            double totalSample = 0;
            double totalBatch = 0;
            int auditorCount = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                if (row["Sample_count"] != DBNull.Value)
                    totalSample += Convert.ToDouble(row["Sample_count"]);

                if (row["Batch_Count"] != DBNull.Value)
                    totalBatch += Convert.ToDouble(row["Batch_Count"]);
            }

            lblTotalSample?.Text = $"Total Sample : {totalSample}";
            lblTotalBatch?.Text = $"Total Batch : {totalBatch}";
            lblTotalAuditor?.Text = $"Total Auditors : {auditorCount}";
            lblBatchRunning?.Text = "Batch_Running :";
            lblBatchPending?.Text = "Batch_Pending :";
            lblBatchComplete?.Text = "Batch_Complete :";
        }

        private void CalculateSummaryForDate()
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
                return;

            double totalSample = 0;
            double totalBatch = 0;

            foreach (DataRow row in dataTable.Rows)
            {
                if (row["Sample_count"] != DBNull.Value)
                    totalSample += Convert.ToDouble(row["Sample_count"]);

                if (row["Batch_Count"] != DBNull.Value)
                    totalBatch += Convert.ToDouble(row["Batch_Count"]);
            }

            lblTotalSample?.Text = $"Total Sample : {totalSample}";
            lblTotalBatch?.Text = $"Total Batch : {totalBatch}";
            lblTotalAuditor?.Text = "Total Auditors : 0";
            lblBatchRunning?.Text = "Batch_Running :";
            lblBatchPending?.Text = "Batch_Pending :";
            lblBatchComplete?.Text = "Batch_Complete :";
        }

        private void BtnSearch_Click(object? sender, EventArgs e)
        {
            string? searchWord = txtFindText?.Text.Trim();
            if (string.IsNullOrEmpty(searchWord))
            {
                MessageBox.Show("Add Search Value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dataGridView?.DataSource == null)
            {
                MessageBox.Show("No data available to search", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool found = false;
            int startRow = dataGridView.CurrentCell?.RowIndex + 1 ?? 0;

            // Search from current position to end
            for (int i = startRow; i < dataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    var cellValue = dataGridView.Rows[i].Cells[j].Value?.ToString() ?? "";
                    if (cellValue.IndexOf(searchWord, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        dataGridView.CurrentCell = dataGridView.Rows[i].Cells[j];
                        dataGridView.FirstDisplayedScrollingRowIndex = i;
                        found = true;
                        break;
                    }
                }
                if (found) break;
            }

            // If not found, search from beginning
            if (!found)
            {
                for (int i = 0; i < startRow; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        var cellValue = dataGridView.Rows[i].Cells[j].Value?.ToString() ?? "";
                        if (cellValue.IndexOf(searchWord, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            dataGridView.CurrentCell = dataGridView.Rows[i].Cells[j];
                            dataGridView.FirstDisplayedScrollingRowIndex = i;
                            found = true;
                            break;
                        }
                    }
                    if (found) break;
                }
            }

            if (!found)
            {
                MessageBox.Show("Value not found.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void DataGridView_DoubleClick(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentCell != null && dataGridView.CurrentCell.Value != null)
            {
                Clipboard.SetText(dataGridView.CurrentCell.Value.ToString() ?? string.Empty);
            }
        }


        private void Form_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                if (pnlFind?.Visible == true)
                {
                    pnlFind.Visible = false;
                    txtFindText?.Text = "";
                }
            }

            if (e.Control && e.KeyCode == Keys.F)
            {
                pnlFind?.Visible = !pnlFind.Visible;
                if (pnlFind?.Visible == true)
                    txtFindText?.Focus();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (dbConnection != null && dbConnection.State == ConnectionState.Open)
            {
                dbConnection.Close();
            }
        }

        // ===== BUTTON CLICK METHODS - ADD THESE AT THE END =====
        // Button click implementations will be provided separately
        private void BtnAuditorReport_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtAuditorList?.Text))
            {
                MessageBox.Show("Please add auditor Id", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string auditorListProcessed = ProcessAuditorList();
                string? fromDateStr = dtpFromDate?.Value.ToString("yyyy-MM-dd");
                string? toDateStr = dtpToDate?.Value.ToString("yyyy-MM-dd");
                string? projectCode = cbProjectList?.Text;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query;

                if (cbProjectList?.SelectedIndex == 0) // All Projects
                {
                    query = $@"
WITH AuditCTE AS (
    SELECT
        Project_Code,
        Batch_Name,
        line as Audit_Type,
        Sample_Count,
        CASE
            WHEN Status = 2 THEN 'Complete'
            WHEN Status = 1 THEN 'Running'
            WHEN Status = -2 THEN 'Canceled'
            WHEN Status = 0 THEN 'Untouched'
            WHEN Status = -1 THEN 'Suspended'
        END AS Status,
        CASE
            WHEN operatorid IS NOT NULL THEN operatorid
            ELSE emp_id
        END AS Auditor_id,
        UserName,
        CONVERT(VARCHAR(8),
            DATEADD(SECOND,
                ISNULL(DATEDIFF(SECOND, Start_DateTime, End_DateTime), 0) +
                ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
                ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0),
            '1900-01-01 00:00:00'),
        108) AS Audit_Time,
        Start_datetime as Start_Date,
        CASE
            WHEN End_DateTime3 IS NOT NULL THEN End_DateTime3
            WHEN End_DateTime2 IS NOT NULL THEN End_DateTime2
            WHEN End_DateTime IS NOT NULL THEN End_DateTime
            ELSE NULL
        END AS Complete_date, comments
    FROM Audit_B
)
SELECT
    Project_Code,
    Batch_Name,
    Audit_Type,
    Sample_Count,
    Status,
    Auditor_id,
    UserName,
    Audit_Time,
    Start_date,
    Complete_date,
    comments
FROM AuditCTE
WHERE Auditor_id in {auditorListProcessed}
    AND CAST(Complete_date as date) >= '{fromDateStr}'
    AND CAST(Complete_date as date) <= '{toDateStr}'";
                }
                else
                {
                    query = $@"
WITH AuditCTE AS (
    SELECT
        Project_Code,
        Batch_Name,
        line as Audit_Type,
        Sample_Count,
        CASE
            WHEN Status = 2 THEN 'Complete'
            WHEN Status = 1 THEN 'Running'
            WHEN Status = -2 THEN 'Canceled'
            WHEN Status = 0 THEN 'Untouched'
            WHEN Status = -1 THEN 'Suspended'
        END AS Status,
        CASE
            WHEN operatorid IS NOT NULL THEN operatorid
            ELSE emp_id
        END AS Auditor_id,
        UserName,
        CONVERT(VARCHAR(8),
            DATEADD(SECOND,
                ISNULL(DATEDIFF(SECOND, Start_DateTime, End_DateTime), 0) +
                ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
                ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0),
            '1900-01-01 00:00:00'),
        108) AS Audit_Time,
        Start_Datetime as Start_date,
        CASE
            WHEN End_DateTime3 IS NOT NULL THEN End_DateTime3
            WHEN End_DateTime2 IS NOT NULL THEN End_DateTime2
            WHEN End_DateTime IS NOT NULL THEN End_DateTime
            ELSE NULL
        END AS Complete_date, comments
    FROM Audit_B
)
SELECT
    Project_Code,
    Batch_Name,
    Audit_Type,
    Sample_Count,
    Status,
    Auditor_id,
    UserName,
    Audit_Time,
    Start_date,
    Complete_date,
    comments
FROM AuditCTE
WHERE Project_Code = @project_code
    AND Auditor_id in {auditorListProcessed}
    AND CAST(Complete_date as date) >= '{fromDateStr}'
    AND CAST(Complete_date as date) <= '{toDateStr}'";
                }

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                if (cbProjectList?.SelectedIndex != 0)
                {
                    cmd.Parameters.AddWithValue("@project_code", projectCode);
                }

                dataAdapter = new SqlDataAdapter(cmd);
                dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridView?.DataSource = dataTable;

                // Set column widths
                if (dataGridView?.Columns.Count >= 11)
                {
                    dataGridView.Columns[0].Width = 150;
                    dataGridView.Columns[1].Width = 360;
                    dataGridView.Columns[2].Width = 70;
                    dataGridView.Columns[3].Width = 90;
                    dataGridView.Columns[4].Width = 100;
                    dataGridView.Columns[5].Width = 70;
                    dataGridView.Columns[6].Width = 250;
                    dataGridView.Columns[7].Width = 120;
                    dataGridView.Columns[8].Width = 200;
                    dataGridView.Columns[9].Width = 200;
                    dataGridView.Columns[10].Width = 200;
                }

                CalculateSummary();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating report: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection.Close();
            }
        }
        private void BtnReportByAuditor_Click(object? sender, EventArgs e)
        {
            try
            {
                string? fromDateStr = dtpFromDate?.Value.ToString("yyyy-MM-dd");
                string? toDateStr = dtpToDate?.Value.ToString("yyyy-MM-dd");
                string? projectCode = cbProjectList?.Text;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query;

                if (cbProjectList?.SelectedIndex == 0) // All Projects
                {
                    query = $@"
WITH ProjectCodes AS (
    SELECT DISTINCT
        Auditor_id,
        project_code
    FROM (
        SELECT
            COALESCE(operatorid, emp_id) AS Auditor_id,
            project_code,
            CASE
                WHEN end_Datetime3 IS NOT NULL THEN end_Datetime3
                WHEN end_Datetime2 IS NOT NULL THEN end_Datetime2
                WHEN end_Datetime IS NOT NULL THEN end_Datetime
            END AS end_date
        FROM audit_b
        WHERE status = 2
    ) AS DistinctProjects
    WHERE
        CAST(end_date AS date) >= '{fromDateStr}' AND
        CAST(end_date AS date) <= '{toDateStr}'
)

SELECT
    Auditor_id,
    MAX(UserName) AS Name,
    COUNT(Batch_name) AS Batch_Count,
    SUM(sample_count) AS Sample_count,
    CONCAT(
        RIGHT('0' + CAST(SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) / 3600 AS VARCHAR(10)), 2),
        ':',
        RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) % 3600) / 60 AS VARCHAR(2)), 2),
        ':',
        RIGHT('0' + CAST(SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) % 60 AS VARCHAR(2)), 2)
    ) AS TotalAuditTime,
    CASE WHEN COUNT(Batch_name) > 0 THEN
        CONCAT(
            RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) / 3600 AS VARCHAR(10)), 2),
            ':',
            RIGHT('0' + CAST(((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) % 3600) / 60 AS VARCHAR(2)), 2),
            ':',
            RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) % 60 AS VARCHAR(2)), 2)
        )
    ELSE
        '00:00:00'
    END AS AvgAuditTimePerBatch,
    (
        SELECT STRING_AGG(project_code, ', ')
        FROM ProjectCodes
        WHERE Auditor_id = MainQuery.Auditor_id
    ) AS Project_Codes
FROM (
    SELECT
        Batch_name,
        sample_count,
        project_code,
        Username,
        COALESCE(operatorid, emp_id) AS Auditor_id,
        CASE
            WHEN end_Datetime3 IS NOT NULL THEN end_Datetime3
            WHEN end_Datetime2 IS NOT NULL THEN end_Datetime2
            WHEN end_Datetime IS NOT NULL THEN end_Datetime
        END AS end_date,
        DATEADD(SECOND,
            ISNULL(DATEDIFF(SECOND, Start_DateTime, End_DateTime), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0),
            '1900-01-01 00:00:00'
        ) AS Audit_Time
    FROM audit_b
    WHERE status = 2
) AS MainQuery
WHERE
    CAST(end_date AS date) >= '{fromDateStr}' AND
    CAST(end_date AS date) <= '{toDateStr}'
GROUP BY
    Auditor_id";
                }
                else
                {
                    string auditorFilter = "";
                    if (!string.IsNullOrWhiteSpace(txtAuditorList?.Text))
                    {
                        string auditorListProcessed = ProcessAuditorList();
                        auditorFilter = $"AND auditor_id in {auditorListProcessed}";
                    }

                    query = $@"
SELECT
    Auditor_id,
    MAX(UserName) AS Name,
    COUNT(Batch_name) AS Batch_Count,
    SUM(sample_count) AS Sample_count,
    CONCAT(
        RIGHT('0' + CAST(SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) / 3600 AS VARCHAR(10)), 2),
        ':',
        RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) % 3600) / 60 AS VARCHAR(2)), 2),
        ':',
        RIGHT('0' + CAST(SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) % 60 AS VARCHAR(2)), 2)
    ) AS TotalAuditTime,
    CASE WHEN COUNT(Batch_name) > 0 THEN
        CONCAT(
            RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) / 3600 AS VARCHAR(10)), 2),
            ':',
            RIGHT('0' + CAST(((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) % 3600) / 60 AS VARCHAR(2)), 2),
            ':',
            RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) % 60 AS VARCHAR(2)), 2)
        )
    ELSE
        '00:00:00'
    END AS AvgAuditTimePerBatch
FROM (
    SELECT
        Batch_name,
        sample_count,
        project_code,
        Username,
        CASE
            WHEN end_Datetime3 IS NOT NULL THEN end_Datetime3
            WHEN end_Datetime2 IS NOT NULL THEN end_Datetime2
            WHEN end_Datetime IS NOT NULL THEN end_Datetime
        END AS end_date,
        CASE
            WHEN operatorid IS NOT NULL THEN operatorid
            ELSE emp_id
        END AS Auditor_id,
        DATEADD(SECOND,
            ISNULL(DATEDIFF(SECOND, Start_DateTime, End_DateTime), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0),
            '1900-01-01 00:00:00'
        ) AS Audit_Time
    FROM audit_b
    WHERE status = 2
) AS subquery
WHERE
    CAST(end_date AS date) >= '{fromDateStr}' AND
    CAST(end_date AS date) <= '{toDateStr}' AND
    project_code = @project_code
    {auditorFilter}
GROUP BY
    Auditor_id";
                }

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                if (cbProjectList?.SelectedIndex != 0)
                {
                    cmd.Parameters.AddWithValue("@project_code", projectCode);
                }

                dataAdapter = new SqlDataAdapter(cmd);
                dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridView?.DataSource = dataTable;

                dataGridView?.Columns["Auditor_id"].DisplayIndex = 0;
                dataGridView?.Columns["Name"].DisplayIndex = 1;
                dataGridView?.Columns["Batch_Count"].DisplayIndex = 2;
                dataGridView?.Columns["Sample_count"].DisplayIndex = 3;

                // Set column widths
                if (dataGridView?.Columns.Count >= 6)
                {
                    dataGridView.Columns[0].Width = 90;
                    dataGridView.Columns[1].Width = 260;
                    dataGridView.Columns[2].Width = 100;
                    dataGridView.Columns[3].Width = 150;
                    dataGridView.Columns[4].Width = 150;
                    dataGridView.Columns[5].Width = 150;
                    if (dataGridView.Columns.Count > 6)
                        dataGridView.Columns[6].Width = 600;
                }

                CalculateSummaryForAuditor();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating report: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }
        private void BtnReportByDate_Click(object? sender, EventArgs e)
        {
            try
            {
                string? fromDateStr = dtpFromDate?.Value.ToString("yyyy-MM-dd");
                string? toDateStr = dtpToDate?.Value.ToString("yyyy-MM-dd");
                string? projectCode = cbProjectList?.Text;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query = $@"
SELECT
    CAST(end_date AS date) AS Complete_Date,
    COUNT(Batch_name) AS Batch_Count,
    SUM(sample_count) AS Sample_count,
    SUM(DATEPART(SECOND, Audit_Time) +
        DATEPART(MINUTE, Audit_Time) * 60 +
        DATEPART(HOUR, Audit_Time) * 3600) AS TotalSeconds,
    CONCAT(
        RIGHT('0' + CAST(SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) / 3600 AS VARCHAR(10)), 2),
        ':',
        RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) % 3600) / 60 AS VARCHAR(2)), 2),
        ':',
        RIGHT('0' + CAST(SUM(DATEPART(SECOND, Audit_Time) +
            DATEPART(MINUTE, Audit_Time) * 60 +
            DATEPART(HOUR, Audit_Time) * 3600) % 60 AS VARCHAR(2)), 2)
    ) AS TotalAuditTime,
    CASE WHEN COUNT(Batch_name) > 0 THEN
        CONCAT(
            RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) / 3600 AS VARCHAR(10)), 2),
            ':',
            RIGHT('0' + CAST(((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) % 3600) / 60 AS VARCHAR(2)), 2),
            ':',
            RIGHT('0' + CAST((SUM(DATEPART(SECOND, Audit_Time) +
                DATEPART(MINUTE, Audit_Time) * 60 +
                DATEPART(HOUR, Audit_Time) * 3600) / COUNT(Batch_name)) % 60 AS VARCHAR(2)), 2)
        )
    ELSE
        '00:00:00'
    END AS AvgAuditTimePerBatch
FROM (
    SELECT
        Batch_name,
        sample_count,
        project_code,
        CASE
            WHEN end_Datetime3 IS NOT NULL THEN end_Datetime3
            WHEN end_Datetime2 IS NOT NULL THEN end_Datetime2
            WHEN end_Datetime IS NOT NULL THEN end_Datetime
        END AS end_date,
        DATEADD(SECOND,
            ISNULL(DATEDIFF(SECOND, Start_DateTime, End_DateTime), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0),
            '1900-01-01 00:00:00'
        ) AS Audit_Time
    FROM audit_b
    WHERE status = 2
) AS subquery
WHERE
    CAST(end_date AS date) >= '{fromDateStr}' AND
    CAST(end_date AS date) <= '{toDateStr}'";

                if (cbProjectList?.SelectedIndex != 0)
                {
                    query += " AND Project_code = @Project_code";
                }

                query += @"
GROUP BY
    CAST(end_date AS date)";

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                if (cbProjectList?.SelectedIndex != 0)
                {
                    cmd.Parameters.AddWithValue("@Project_code", projectCode);
                }

                dataAdapter = new SqlDataAdapter(cmd);
                dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridView?.DataSource = dataTable;

                // Set column widths
                if (dataGridView?.Columns.Count >= 6)
                {
                    dataGridView.Columns[0].Width = 255;
                    dataGridView.Columns[1].Width = 155;
                    dataGridView.Columns[2].Width = 155;
                    dataGridView.Columns[3].Width = 155;
                    dataGridView.Columns[4].Width = 155;
                    dataGridView.Columns[5].Width = 155;
                }

                CalculateSummaryForDate();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating report: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }
        private void BtnReportByProject_Click(object? sender, EventArgs e)
        {
            try
            {
                string? fromDateStr = dtpFromDate?.Value.ToString("yyyy-MM-dd");
                string? toDateStr = dtpToDate?.Value.ToString("yyyy-MM-dd");

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query = $@"
WITH TimeData AS (
    SELECT
        Project_Code,
        COUNT(Batch_name) AS Batch_Count,
        SUM(sample_count) AS Sample_count,
        SUM(
            ISNULL(DATEDIFF(SECOND, Start_DateTime, End_DateTime), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0)
        ) AS TotalSeconds
    FROM audit_b
    WHERE status = 2
    AND CAST(COALESCE(End_DateTime3, End_DateTime2, End_DateTime) AS date)
        BETWEEN '{fromDateStr}' AND '{toDateStr}'
    GROUP BY Project_Code
)

SELECT
    Project_Code,
    Batch_Count,
    Sample_count,
    CONCAT(
        RIGHT('0' + CAST(TotalSeconds / 3600 AS VARCHAR(10)), 2), ':',
        RIGHT('0' + CAST((TotalSeconds % 3600) / 60 AS VARCHAR(2)), 2), ':',
        RIGHT('0' + CAST(TotalSeconds % 60 AS VARCHAR(2)), 2)
    ) AS TotalAuditTime,
    CASE WHEN Batch_Count > 0 THEN
        CONCAT(
            RIGHT('0' + CAST((TotalSeconds / Batch_Count) / 3600 AS VARCHAR(10)), 2), ':',
            RIGHT('0' + CAST(((TotalSeconds / Batch_Count) % 3600) / 60 AS VARCHAR(2)), 2), ':',
            RIGHT('0' + CAST((TotalSeconds / Batch_Count) % 60 AS VARCHAR(2)), 2)
        )
    ELSE
        '00:00:00'
    END AS AvgAuditTimePerBatch
FROM TimeData";

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                dataAdapter = new SqlDataAdapter(cmd);
                dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridView?.DataSource = dataTable;

                // Set column widths
                if (dataGridView?.Columns.Count >= 5)
                {
                    dataGridView.Columns[0].Width = 255;
                    dataGridView.Columns[1].Width = 155;
                    dataGridView.Columns[2].Width = 155;
                    dataGridView.Columns[3].Width = 200;
                    dataGridView.Columns[4].Width = 200;
                }

                CalculateSummaryForDate();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating report: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }

        private void ExportToExcel_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "Report.xlsx"
            };

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Report");

            int colIndex = 1;

            // Headers
            if (dataGridView != null)
            {
                foreach (DataGridViewColumn col in dataGridView.Columns)
                {
                    if (!col.Visible) continue;

                    worksheet.Cell(1, colIndex).Value = col.HeaderText;
                    worksheet.Cell(1, colIndex).Style.Font.Bold = true;
                    colIndex++;
                }
            }

            // Data
            if (dataGridView != null)
            {
                for (int r = 0; r < dataGridView.Rows.Count; r++)
                {
                    colIndex = 1;

                    foreach (DataGridViewColumn col in dataGridView.Columns)
                    {
                        if (!col.Visible) continue;

                        var value = dataGridView.Rows[r].Cells[col.Index].Value;

                        if (value is DateTime dt)
                            worksheet.Cell(r + 2, colIndex).Value = dt;
                        else if (value is int or long or float or double or decimal)
                            worksheet.Cell(r + 2, colIndex).Value = Convert.ToDouble(value);
                        else
                            worksheet.Cell(r + 2, colIndex).Value = value?.ToString() ?? "";

                        colIndex++;
                    }
                }
            }
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(sfd.FileName);

            MessageBox.Show("Export completed successfully!", "Success",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = sfd.FileName,
                    UseShellExecute = true // important!
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open Excel file: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAttendance_Click(object? sender, EventArgs e)
        {
            try
            {
                string? fromDateStr = dtpFromDate?.Value.ToString("yyyy-MM-dd");
                string? toDateStr = dtpToDate?.Value.ToString("yyyy-MM-dd");

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();
                string query = $@"
DECLARE @StartDate DATE = '{fromDateStr}';
DECLARE @EndDate DATE = '{toDateStr}';

WITH DaysInMonth AS (
    SELECT @StartDate AS WorkDay
    UNION ALL
    SELECT DATEADD(DAY, 1, WorkDay)
    FROM DaysInMonth
    WHERE WorkDay < @EndDate
),

Employee_Names AS (
    SELECT DISTINCT COALESCE(Emp_ID, OperatorID) AS Auditor_ID, UserName
    FROM Audit_B
    WHERE (Start_DateTime BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate))
       OR (Start_DateTime2 BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate))
       OR (Start_DateTime3 BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate))
),

Employee_Workdays AS (
    SELECT COALESCE(Emp_ID, OperatorID) AS Auditor_ID,
           CAST(Start_DateTime AS DATE) AS WorkDate
    FROM Audit_B
    WHERE Start_DateTime BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate)
    UNION
    SELECT COALESCE(Emp_ID, OperatorID) AS Auditor_ID,
           CAST(Start_DateTime2 AS DATE) AS WorkDate
    FROM Audit_B
    WHERE Start_DateTime2 BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate)
    UNION
    SELECT COALESCE(Emp_ID, OperatorID) AS Auditor_ID,
           CAST(Start_DateTime3 AS DATE) AS WorkDate
    FROM Audit_B
    WHERE Start_DateTime3 BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate)
),

HolidayDates AS (
    SELECT CAST(HOLIDAY_START_DATE AS DATE) AS HolidayDate
    FROM Holiday_info
    WHERE HOLIDAY_START_DATE BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate)
       OR HOLIDAY_END_DATE BETWEEN @StartDate AND DATEADD(DAY, 1, @EndDate)
),

Attendance AS (
    SELECT D.WorkDay, E.Auditor_ID, E.UserName,
           CASE
               WHEN EW.WorkDate IS NOT NULL THEN 'P'
               WHEN HD.HolidayDate IS NOT NULL THEN 'H'
               WHEN DATENAME(WEEKDAY, D.WorkDay) = 'Friday' THEN 'H'
               ELSE 'A'
           END AS Status
    FROM DaysInMonth D
    CROSS JOIN Employee_Names E
    LEFT JOIN Employee_Workdays EW
        ON E.Auditor_ID = EW.Auditor_ID AND D.WorkDay = EW.WorkDate
    LEFT JOIN HolidayDates HD
        ON D.WorkDay = HD.HolidayDate
)

SELECT
    A.Auditor_ID,
    A.UserName,
    MAX(CASE WHEN DAY(A.WorkDay) = 1 THEN A.Status END) AS [1],
    MAX(CASE WHEN DAY(A.WorkDay) = 2 THEN A.Status END) AS [2],
    MAX(CASE WHEN DAY(A.WorkDay) = 3 THEN A.Status END) AS [3],
    MAX(CASE WHEN DAY(A.WorkDay) = 4 THEN A.Status END) AS [4],
    MAX(CASE WHEN DAY(A.WorkDay) = 5 THEN A.Status END) AS [5],
    MAX(CASE WHEN DAY(A.WorkDay) = 6 THEN A.Status END) AS [6],
    MAX(CASE WHEN DAY(A.WorkDay) = 7 THEN A.Status END) AS [7],
    MAX(CASE WHEN DAY(A.WorkDay) = 8 THEN A.Status END) AS [8],
    MAX(CASE WHEN DAY(A.WorkDay) = 9 THEN A.Status END) AS [9],
    MAX(CASE WHEN DAY(A.WorkDay) = 10 THEN A.Status END) AS [10],
    MAX(CASE WHEN DAY(A.WorkDay) = 11 THEN A.Status END) AS [11],
    MAX(CASE WHEN DAY(A.WorkDay) = 12 THEN A.Status END) AS [12],
    MAX(CASE WHEN DAY(A.WorkDay) = 13 THEN A.Status END) AS [13],
    MAX(CASE WHEN DAY(A.WorkDay) = 14 THEN A.Status END) AS [14],
    MAX(CASE WHEN DAY(A.WorkDay) = 15 THEN A.Status END) AS [15],
    MAX(CASE WHEN DAY(A.WorkDay) = 16 THEN A.Status END) AS [16],
    MAX(CASE WHEN DAY(A.WorkDay) = 17 THEN A.Status END) AS [17],
    MAX(CASE WHEN DAY(A.WorkDay) = 18 THEN A.Status END) AS [18],
    MAX(CASE WHEN DAY(A.WorkDay) = 19 THEN A.Status END) AS [19],
    MAX(CASE WHEN DAY(A.WorkDay) = 20 THEN A.Status END) AS [20],
    MAX(CASE WHEN DAY(A.WorkDay) = 21 THEN A.Status END) AS [21],
    MAX(CASE WHEN DAY(A.WorkDay) = 22 THEN A.Status END) AS [22],
    MAX(CASE WHEN DAY(A.WorkDay) = 23 THEN A.Status END) AS [23],
    MAX(CASE WHEN DAY(A.WorkDay) = 24 THEN A.Status END) AS [24],
    MAX(CASE WHEN DAY(A.WorkDay) = 25 THEN A.Status END) AS [25],
    MAX(CASE WHEN DAY(A.WorkDay) = 26 THEN A.Status END) AS [26],
    MAX(CASE WHEN DAY(A.WorkDay) = 27 THEN A.Status END) AS [27],
    MAX(CASE WHEN DAY(A.WorkDay) = 28 THEN A.Status END) AS [28],
    MAX(CASE WHEN DAY(A.WorkDay) = 29 THEN A.Status END) AS [29],
    MAX(CASE WHEN DAY(A.WorkDay) = 30 THEN A.Status END) AS [30],
    MAX(CASE WHEN DAY(A.WorkDay) = 31 THEN A.Status END) AS [31],
    SUM(CASE WHEN A.Status = 'P' THEN 1 ELSE 0 END) AS Total_Present
FROM Attendance A
GROUP BY A.Auditor_ID, A.UserName
ORDER BY A.Auditor_ID
OPTION (MAXRECURSION 0)";

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                dataAdapter = new SqlDataAdapter(cmd);
                dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridView?.DataSource = dataTable;

                // Set column widths for attendance report
                dataGridView?.Columns[0].Width = 100; // Auditor_ID
                dataGridView?.Columns[1].Width = 300; // UserName
                if (dataGridView?.Columns.Count > 3)
                {
                    for (int i = 2; i < dataGridView.Columns.Count - 1; i++)
                    {
                        dataGridView.Columns[i].Width = 30;
                    }
                }

                lblTotalAuditor?.Text = "Total Auditor :" + dataTable.Rows.Count.ToString();
                lblBatchComplete?.Text = "Batch Complete :";
                lblBatchRunning?.Text = "Batch Running :";
                lblBatchPending?.Text = "Batch Pending :";
                lblTotalBatch?.Text = "Total Batch :";
                lblTotalSample?.Text = "Total Sample :";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating attendance report: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }
    }
}