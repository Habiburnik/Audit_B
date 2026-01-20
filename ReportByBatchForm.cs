using System.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Diagnostics;

namespace Audit_B
{
    public partial class ReportByBatchForm : Form
    {
        private SqlConnection? dbConnection;
        private SqlDataAdapter? dataAdapter;
        private DataTable? dataTable;
        private SqlCommand? queryUpdate;

        // Top Panel Controls
        private Panel? pnlTop;
        private Label? lblProjectCode;
        private ComboBox? cbProjectList;
        private CheckBox? checkBoxShowComplete;
        private Label? lblFrom;
        private DateTimePicker? dtpFromDate;
        private Label? lblTo;
        private DateTimePicker? dtpToDate;
        private Label? lblBatchList;
        private TextBox? txtBatchList;
        private Button? btnAllProjectStatus;
        private Button? btnReportAll;
        private Button? btnReportByDate;
        private Button? btnDeleteBatches;
        private Button? btnBatchReset;

        // GroupBox - Update Queue
        private GroupBox? grpUpdateQue;
        private Label? lblSelectProject;
        private ComboBox? cbQueProjectList;
        private Label? lblQueCount;
        private TextBox? txtQueCount;
        private Button? btnQueUpdate;

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
        private Label? lblTotalQue;

        // Find Panel
        private Panel? pnlFind;
        private TextBox? txtFindText;
        private Button? btnSearch;

        // Project Update Panel
        private Panel? pnlProjectUpdate;
        private Label? lblProjectCode2;
        private TextBox? txtProjectCode;
        private Label? lblTotalBatch2;
        private TextBox? txtTotalBatch;
        private Label? lblLanguage;
        private TextBox? txtLanguage;
        private Label? lblDeadline;
        private TextBox? txtDeadline;
        private Label? lblComments;
        private TextBox? txtComments;
        private Button? btnProjectUpdate;

        // Auditor List Panel
        private Panel? pnlAuditorList;
        private Label? lblAuditorListTitle;
        private ListBox? listBoxAuditors;

        // Context Menu
        private ContextMenuStrip? contextMenu;
        private ToolStripMenuItem? menuProjectComplete;
        private ToolStripMenuItem? menuProjectStop;
        private ToolStripMenuItem? menuProjectRunning;
        private ToolStripMenuItem? menuUpdateProjectInfo;
        private ToolStripMenuItem? menuUpdateBlank;
        private ToolStripMenuItem? menuShowAuditors;
        private ToolStripMenuItem? menuExportToExcel;

        // Hidden memo for processing
        private TextBox? txtBatchList2;

        // For permission check
        private MainMenuForm mainMenuForm;

        public ReportByBatchForm(MainMenuForm parentForm)
        {
            mainMenuForm = parentForm;
            InitializeDatabase();
            InitializeComponent();
            LoadProjects();
        }

        private void InitializeDatabase()
        {
            try
            {
                string connectionString = ConfigurationHelper.GetConnectionString();
                dbConnection = new SqlConnection(connectionString);
                queryUpdate = new SqlCommand();
                queryUpdate.Connection = dbConnection;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database initialization error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "Report By Batch";
            this.Size = new Size(1400, 800);
            this.StartPosition = FormStartPosition.CenterParent;
            // this.BackColor = Color.FromArgb(139, 175, 185);
            this.KeyPreview = true;
            this.KeyDown += Form_KeyDown;
            // this.Icon = new Icon("AuditB_icon.ico");

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

            // ===== TOP PANEL =====
            pnlTop = new Panel();
            pnlTop.Dock = DockStyle.Top;
            pnlTop.Height = 195;
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
            cbProjectList.DropDownStyle = ComboBoxStyle.DropDown;

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

            // Buttons Row 1
            btnAllProjectStatus = new Button();
            btnAllProjectStatus.Text = "All Project Status";
            btnAllProjectStatus.Location = new Point(10, 155);
            btnAllProjectStatus.Size = new Size(120, 30);
            btnAllProjectStatus.BackColor = Color.White;
            btnAllProjectStatus.Click += BtnAllProjectStatus_Click;

            btnReportAll = new Button();
            btnReportAll.Text = "Report All";
            btnReportAll.Location = new Point(135, 155);
            btnReportAll.Size = new Size(120, 30);
            btnReportAll.BackColor = Color.White;
            btnReportAll.Click += BtnReportAll_Click;

            btnReportByDate = new Button();
            btnReportByDate.Text = "Report By Date";
            btnReportByDate.Location = new Point(260, 155);
            btnReportByDate.Size = new Size(120, 30);
            btnReportByDate.BackColor = Color.White;
            btnReportByDate.Click += BtnReportByDate_Click;

            // Batch List Label
            lblBatchList = new Label();
            lblBatchList.Text = "Batch_List";
            lblBatchList.Location = new Point(400, 10);
            lblBatchList.AutoSize = true;
            lblBatchList.Font = new Font("Arial", 10, FontStyle.Regular);

            // Batch List TextBox (Multiline)
            txtBatchList = new TextBox();
            txtBatchList.Location = new Point(400, 35);
            txtBatchList.Size = new Size(350, 110);
            txtBatchList.Multiline = true;
            txtBatchList.ScrollBars = ScrollBars.Vertical;
            txtBatchList.Font = new Font("Arial", 9);

            // Delete and Reset Buttons
            btnDeleteBatches = new Button();
            btnDeleteBatches.Text = "Delete Batches";
            btnDeleteBatches.Location = new Point(400, 155);
            btnDeleteBatches.Size = new Size(120, 30);
            btnDeleteBatches.BackColor = Color.White;
            btnDeleteBatches.Visible = false;
            btnDeleteBatches.Click += BtnDeleteBatches_Click;

            btnBatchReset = new Button();
            btnBatchReset.Text = "Batch Reset";
            btnBatchReset.Location = new Point(525, 155);
            btnBatchReset.Size = new Size(120, 30);
            btnBatchReset.BackColor = Color.White;
            btnBatchReset.Visible = false;
            btnBatchReset.Click += BtnBatchReset_Click;

            // Hidden Batch List 2
            txtBatchList2 = new TextBox();
            txtBatchList2.Visible = false;
            txtBatchList2.Multiline = true;

            // ===== UPDATE QUEUE GROUPBOX =====
            grpUpdateQue = new GroupBox();
            grpUpdateQue.Text = "Update Que";
            grpUpdateQue.Location = new Point(770, 10);
            grpUpdateQue.Size = new Size(250, 175);
            // grpUpdateQue.BackColor = Color.FromArgb(139, 175, 185);

            lblSelectProject = new Label();
            lblSelectProject.Text = "Select Project :";
            lblSelectProject.Location = new Point(10, 25);
            lblSelectProject.AutoSize = true;

            cbQueProjectList = new ComboBox();
            cbQueProjectList.Location = new Point(10, 50);
            cbQueProjectList.Size = new Size(230, 25);
            cbQueProjectList.DropDownStyle = ComboBoxStyle.DropDown;

            lblQueCount = new Label();
            lblQueCount.Text = "Que Count :";
            lblQueCount.Location = new Point(10, 85);
            lblQueCount.AutoSize = true;

            txtQueCount = new TextBox();
            txtQueCount.Location = new Point(10, 110);
            txtQueCount.Size = new Size(230, 25);

            btnQueUpdate = new Button();
            btnQueUpdate.Text = "Update Que";
            btnQueUpdate.Location = new Point(10, 140);
            btnQueUpdate.Size = new Size(100, 25);
            btnQueUpdate.BackColor = Color.White;
            btnQueUpdate.Click += BtnQueUpdate_Click;

            grpUpdateQue.Controls.AddRange(new Control[] {
                lblSelectProject, cbQueProjectList, lblQueCount, txtQueCount, btnQueUpdate
            });

            // Add controls to top panel
            pnlTop.Controls.AddRange(new Control[] {
                lblProjectCode, cbProjectList, checkBoxShowComplete,
                lblFrom, dtpFromDate, lblTo, dtpToDate,
                btnAllProjectStatus, btnReportAll, btnReportByDate,
                lblBatchList, txtBatchList, txtBatchList2,
                btnDeleteBatches, btnBatchReset,
                grpUpdateQue
            });

            // ===== DATAGRIDVIEW =====
            dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.ReadOnly = true;
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.BackgroundColor = Color.White;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.DoubleClick += DataGridView_DoubleClick;
            dataGridView.ColumnHeaderMouseClick += DataGridView_ColumnHeaderMouseClick;
            dataGridView.DataBindingComplete += dataGridView_DataBindingComplete;

            // Context Menu for DataGridView
            contextMenu = new ContextMenuStrip();

            menuProjectComplete = new ToolStripMenuItem("Project Complete");
            menuProjectComplete.Click += MenuProjectComplete_Click;
            menuProjectComplete.Visible = false;

            menuProjectStop = new ToolStripMenuItem("Project Stop");
            menuProjectStop.Click += MenuProjectStop_Click;
            menuProjectStop.Visible = false;

            menuProjectRunning = new ToolStripMenuItem("Project Running");
            menuProjectRunning.Click += MenuProjectRunning_Click;
            menuProjectRunning.Visible = false;

            menuUpdateProjectInfo = new ToolStripMenuItem("Update Project Info");
            menuUpdateProjectInfo.Click += MenuUpdateProjectInfo_Click;
            menuUpdateProjectInfo.Visible = false;

            menuUpdateBlank = new ToolStripMenuItem("Update Blank");
            menuUpdateBlank.Click += MenuUpdateBlank_Click;
            menuUpdateBlank.Visible = false;

            menuShowAuditors = new ToolStripMenuItem("Show Auditors");
            menuShowAuditors.Click += MenuShowAuditors_Click;

            menuExportToExcel = new ToolStripMenuItem("Export to Excel");
            menuExportToExcel.Click += ExportToExcel_Click;

            contextMenu.Items.AddRange(new ToolStripItem[] {
                menuProjectComplete, menuProjectStop, menuProjectRunning,
                new ToolStripSeparator(),
                menuUpdateProjectInfo, menuUpdateBlank,
                new ToolStripSeparator(),
                menuShowAuditors, menuExportToExcel
            });

            dataGridView.ContextMenuStrip = contextMenu;

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
            lblBatchComplete.Location = new Point(180, 15);
            lblBatchComplete.AutoSize = true;
            lblBatchComplete.Font = new Font("Arial", 10, FontStyle.Bold);

            lblBatchRunning = new Label();
            lblBatchRunning.Text = "Batch Running :";
            lblBatchRunning.Location = new Point(350, 15);
            lblBatchRunning.AutoSize = true;
            lblBatchRunning.Font = new Font("Arial", 10, FontStyle.Bold);

            lblBatchPending = new Label();
            lblBatchPending.Text = "Batch Pending :";
            lblBatchPending.Location = new Point(520, 15);
            lblBatchPending.AutoSize = true;
            lblBatchPending.Font = new Font("Arial", 10, FontStyle.Bold);

            lblTotalBatch = new Label();
            lblTotalBatch.Text = "Total Batch :";
            lblTotalBatch.Location = new Point(690, 15);
            lblTotalBatch.AutoSize = true;
            lblTotalBatch.Font = new Font("Arial", 10, FontStyle.Bold);

            lblTotalSample = new Label();
            lblTotalSample.Text = "Total Sample :";
            lblTotalSample.Location = new Point(860, 15);
            lblTotalSample.AutoSize = true;
            lblTotalSample.Font = new Font("Arial", 10, FontStyle.Bold);

            lblTotalQue = new Label();
            lblTotalQue.Text = "Total Que : 0";
            lblTotalQue.Location = new Point(1030, 15);
            lblTotalQue.AutoSize = true;
            lblTotalQue.Font = new Font("Arial", 10, FontStyle.Bold);

            pnlBottom.Controls.AddRange(new Control[] {
                lblTotalAuditor, lblBatchComplete, lblBatchRunning,
                lblBatchPending, lblTotalBatch, lblTotalSample, lblTotalQue
            });

            // ===== FIND PANEL (Hidden by default) =====
            pnlFind = new Panel();
            pnlFind.Size = new Size(300, 50);
            pnlFind.Location = new Point(this.ClientSize.Width - 320, pnlTop.Height - 60);
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

            // ===== PROJECT UPDATE PANEL (Hidden by default) =====
            pnlProjectUpdate = new Panel();
            pnlProjectUpdate.Size = new Size(400, 350);
            pnlProjectUpdate.Location = new Point(
                (this.ClientSize.Width - 400) / 2,
                (this.ClientSize.Height - 350) / 2
            );
            pnlProjectUpdate.BackColor = Color.FromArgb(200, 200, 200);
            pnlProjectUpdate.BorderStyle = BorderStyle.FixedSingle;
            pnlProjectUpdate.Visible = false;

            lblProjectCode2 = new Label();
            lblProjectCode2.Text = "Project Code :";
            lblProjectCode2.Location = new Point(10, 20);
            lblProjectCode2.AutoSize = true;

            txtProjectCode = new TextBox();
            txtProjectCode.Location = new Point(120, 20);
            txtProjectCode.Size = new Size(260, 25);
            txtProjectCode.ReadOnly = true;

            lblTotalBatch2 = new Label();
            lblTotalBatch2.Text = "Total Batch :";
            lblTotalBatch2.Location = new Point(10, 60);
            lblTotalBatch2.AutoSize = true;

            txtTotalBatch = new TextBox();
            txtTotalBatch.Location = new Point(120, 60);
            txtTotalBatch.Size = new Size(260, 25);

            lblLanguage = new Label();
            lblLanguage.Text = "Language :";
            lblLanguage.Location = new Point(10, 100);
            lblLanguage.AutoSize = true;

            txtLanguage = new TextBox();
            txtLanguage.Location = new Point(120, 100);
            txtLanguage.Size = new Size(260, 25);

            lblDeadline = new Label();
            lblDeadline.Text = "Deadline :";
            lblDeadline.Location = new Point(10, 140);
            lblDeadline.AutoSize = true;

            txtDeadline = new TextBox();
            txtDeadline.Location = new Point(120, 140);
            txtDeadline.Size = new Size(260, 25);

            lblComments = new Label();
            lblComments.Text = "Comments :";
            lblComments.Location = new Point(10, 180);
            lblComments.AutoSize = true;

            txtComments = new TextBox();
            txtComments.Location = new Point(120, 180);
            txtComments.Size = new Size(260, 100);
            txtComments.Multiline = true;
            txtComments.ScrollBars = ScrollBars.Vertical;

            btnProjectUpdate = new Button();
            btnProjectUpdate.Text = "Update";
            btnProjectUpdate.Location = new Point(150, 300);
            btnProjectUpdate.Size = new Size(100, 30);
            btnProjectUpdate.BackColor = Color.White;
            btnProjectUpdate.Click += BtnProjectUpdate_Click;

            pnlProjectUpdate.Controls.AddRange(new Control[] {
                lblProjectCode2, txtProjectCode, lblTotalBatch2, txtTotalBatch,
                lblLanguage, txtLanguage, lblDeadline, txtDeadline,
                lblComments, txtComments, btnProjectUpdate
            });

            // ===== AUDITOR LIST PANEL (Hidden by default) =====
            pnlAuditorList = new Panel();
            pnlAuditorList.Size = new Size(300, 400);
            pnlAuditorList.Location = new Point(
                (this.ClientSize.Width - 300) / 2,
                (this.ClientSize.Height - 400) / 2
            );
            pnlAuditorList.BackColor = Color.FromArgb(200, 200, 200);
            pnlAuditorList.BorderStyle = BorderStyle.FixedSingle;
            pnlAuditorList.Visible = false;

            lblAuditorListTitle = new Label();
            lblAuditorListTitle.Text = "Running Auditors";
            lblAuditorListTitle.Location = new Point(10, 10);
            lblAuditorListTitle.AutoSize = true;
            lblAuditorListTitle.Font = new Font("Arial", 12, FontStyle.Bold);

            listBoxAuditors = new ListBox();
            listBoxAuditors.Location = new Point(10, 40);
            listBoxAuditors.Size = new Size(280, 350);

            pnlAuditorList.Controls.AddRange(new Control[] {
                lblAuditorListTitle, listBoxAuditors
            });

            // Add all main controls to form
            this.Controls.Add(dataGridView);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);
            this.Controls.Add(pnlFind);
            this.Controls.Add(pnlProjectUpdate);
            this.Controls.Add(pnlAuditorList);

            // Bring panels to front
            pnlFind.BringToFront();
            pnlProjectUpdate.BringToFront();
            pnlAuditorList.BringToFront();

            // Check permissions
            CheckPermissions();
        }

        private void CheckPermissions()
        {
            if (mainMenuForm != null && mainMenuForm.edtPermission?.Text == "Yes")
            {
                btnDeleteBatches?.Visible = true;
                btnBatchReset?.Visible = true;
            }
            else
            {
                btnDeleteBatches?.Visible = false;
                btnBatchReset?.Visible = false;
            }
        }

        private void LoadProjects()
        {
            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                cbProjectList?.Items.Clear();
                cbQueProjectList?.Items.Clear();
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
                    string? projectCode = reader["project_code"].ToString();
                    cbProjectList?.Items.Add(projectCode ?? "");
                    cbQueProjectList?.Items.Add(projectCode ?? "");
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
                    query =
                    "SELECT DISTINCT project_code FROM audit_b_projects WHERE status IN (1)";

                }
                else
                {
                    query = @"SELECT project_code FROM audit_b_projects 
                                WHERE status IN (0, -1, -2)
                                ORDER BY CASE 
                                    WHEN Status = -1 THEN 1
                                    WHEN Status = 0 THEN 2
                                    WHEN Status = -2 THEN 3
                                    WHEN Status = 1 THEN 4
                                END, Deadline";
                }

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string? projectCode = Convert.ToString(reader["project_code"]);

                    if (!string.IsNullOrWhiteSpace(projectCode))
                    {
                        cbProjectList?.Items.Add(projectCode);
                    }
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

        private string ProcessBatchList()
        {
            txtBatchList2?.Clear();
            txtBatchList2?.Lines = new string[] { "(" };

            var lines = txtBatchList?.Lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();

            if (lines != null)
            {
                foreach (var line in lines)
                {
                    txtBatchList2?.AppendText($"'{line.Trim()}',\r\n");
                }
            }
            txtBatchList2?.AppendText(")");

            string result = txtBatchList2?.Text.Replace(",\r\n)", "\r\n)") ?? "";
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
                if (pnlProjectUpdate?.Visible == true)
                {
                    pnlProjectUpdate.Visible = false;
                }
                if (pnlFind?.Visible == true)
                {
                    pnlFind.Visible = false;
                    txtFindText?.Text = "";
                }
                if (pnlAuditorList?.Visible == true)
                {
                    pnlAuditorList.Visible = false;
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
        private void BtnReportAll_Click(object? sender, EventArgs e)
        {
            lblTotalQue?.Text = "Total Que : 0";
            menuProjectComplete?.Visible = false;
            menuProjectStop?.Visible = false;
            menuProjectRunning?.Visible = false;
            menuUpdateBlank?.Visible = false;
            menuUpdateProjectInfo?.Visible = false;
            string? projectCode = cbProjectList?.Text;

            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query;

                if (string.IsNullOrWhiteSpace(txtBatchList?.Text))
                {
                    // Query without batch filter
                    query = @"
SELECT
    Project_Code,
    Batch_Name,
    Line AS Audit_Type,
    Sample_Count,
    CASE Status
        WHEN  2 THEN 'Complete'
        WHEN  1 THEN 'Running'
        WHEN -2 THEN 'Canceled'
        WHEN  0 THEN 'Untouched'
        WHEN -1 THEN 'Suspended'
    END AS Status,
    COALESCE(operatorid, emp_id) AS Auditor_id,
    UserName,
    CONCAT(
        CAST(t.total_seconds / 3600 AS VARCHAR(10)),
        ':',
        RIGHT('0' + CAST((t.total_seconds % 3600) / 60 AS VARCHAR(2)), 2),
        ':',
        RIGHT('0' + CAST(t.total_seconds % 60 AS VARCHAR(2)), 2)
    ) AS Audit_Time,
    Start_DateTime AS Start_date,
    COALESCE(End_DateTime3, End_DateTime2, End_DateTime) AS Complete_date,
    comments
FROM Audit_B
CROSS APPLY (
    SELECT
        ISNULL(DATEDIFF(SECOND, Start_DateTime,  End_DateTime),  0) +
        ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
        ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0)
        AS total_seconds
) t
WHERE
    project_code = @project_code
ORDER BY
    Start_date DESC";
                }
                else
                {
                    // Query with batch filter
                    string batchListProcessed = ProcessBatchList();
                    query = $@"
SELECT
    Project_Code,
    Batch_Name,
    Line AS Audit_Type,
    Sample_Count,
    CASE Status
        WHEN  2 THEN 'Complete'
        WHEN  1 THEN 'Running'
        WHEN -2 THEN 'Canceled'
        WHEN  0 THEN 'Untouched'
        WHEN -1 THEN 'Suspended'
    END AS Status,
    COALESCE(operatorid, emp_id) AS Auditor_id,
    UserName,
    CONCAT(
        CAST(t.total_seconds / 3600 AS VARCHAR(10)),
        ':',
        RIGHT('0' + CAST((t.total_seconds % 3600) / 60 AS VARCHAR(2)), 2),
        ':',
        RIGHT('0' + CAST(t.total_seconds % 60 AS VARCHAR(2)), 2)
    ) AS Audit_Time,
    Start_DateTime AS Start_date,
    COALESCE(End_DateTime3, End_DateTime2, End_DateTime) AS Complete_date,
    comments
FROM Audit_B
CROSS APPLY (
    SELECT
        ISNULL(DATEDIFF(SECOND, Start_DateTime,  End_DateTime),  0) +
        ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
        ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0)
        AS total_seconds
) t
WHERE
    project_code = @project_code AND batch_name in {batchListProcessed}
ORDER BY
    Start_date DESC";
                }

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                cmd.Parameters.AddWithValue("@project_code", projectCode);

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
                    dataGridView.Columns[8].Width = 180;
                    dataGridView.Columns[9].Width = 180;
                    dataGridView.Columns[10].Width = 200;
                }

                CalculateSummary();

                if (dataGridView?.Rows.Count > 0)
                {
                    dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                }
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
        private void BtnReportByDate_Click(object? sender, EventArgs e)
        {
            lblTotalQue?.Text = "Total Que : 0";
            menuProjectComplete?.Visible = false;
            menuProjectStop?.Visible = false;
            menuProjectRunning?.Visible = false;
            menuUpdateBlank?.Visible = false;
            menuUpdateProjectInfo?.Visible = false;

            string fromDateStr = dtpFromDate?.Value.ToString("yyyy-MM-dd") ?? "";
            string toDateStr = dtpToDate?.Value.ToString("yyyy-MM-dd") ?? "";
            string? projectCode = cbProjectList?.Text;

            try
            {
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
        Line AS Audit_Type,
        Sample_Count,
        CASE Status
            WHEN  2 THEN 'Complete'
            WHEN  1 THEN 'Running'
            WHEN -2 THEN 'Canceled'
            WHEN  0 THEN 'Untouched'
            WHEN -1 THEN 'Suspended'
        END AS Status,
        COALESCE(operatorid, emp_id) AS Auditor_id,
        UserName,
        CONCAT(
            CAST(t.total_seconds / 3600 AS VARCHAR(10)),
            ':',
            RIGHT('0' + CAST((t.total_seconds % 3600) / 60 AS VARCHAR(2)), 2),
            ':',
            RIGHT('0' + CAST(t.total_seconds % 60 AS VARCHAR(2)), 2)
        ) AS Audit_Time,
        Start_DateTime AS Start_date,
        COALESCE(End_DateTime3, End_DateTime2, End_DateTime) AS Complete_date,
        comments
    FROM Audit_B
    CROSS APPLY (
        SELECT
            ISNULL(DATEDIFF(SECOND, Start_DateTime,  End_DateTime),  0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0)
            AS total_seconds
    ) t
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
WHERE
    CAST(Complete_date AS date) >= '{fromDateStr}'
    AND CAST(Complete_date AS date) <= '{toDateStr}' ";
                }
                else if (string.IsNullOrWhiteSpace(txtBatchList?.Text))
                {
                    // Specific project without batch filter
                    query = $@"
WITH AuditCTE AS (
    SELECT
        Project_Code,
        Batch_Name,
        Line AS Audit_Type,
        Sample_Count,
        CASE Status
            WHEN  2 THEN 'Complete'
            WHEN  1 THEN 'Running'
            WHEN -2 THEN 'Canceled'
            WHEN  0 THEN 'Untouched'
            WHEN -1 THEN 'Suspended'
        END AS Status,
        COALESCE(operatorid, emp_id) AS Auditor_id,
        UserName,
        CONCAT(
            CAST(t.total_seconds / 3600 AS VARCHAR(10)),
            ':',
            RIGHT('0' + CAST((t.total_seconds % 3600) / 60 AS VARCHAR(2)), 2),
            ':',
            RIGHT('0' + CAST(t.total_seconds % 60 AS VARCHAR(2)), 2)
        ) AS Audit_Time,
        Start_DateTime AS Start_date,
        COALESCE(End_DateTime3, End_DateTime2, End_DateTime) AS Complete_date,
        comments
    FROM Audit_B
    CROSS APPLY (
        SELECT
            ISNULL(DATEDIFF(SECOND, Start_DateTime,  End_DateTime),  0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0)
            AS total_seconds
    ) t
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
WHERE
    Project_Code = @project_code
    AND CAST(Complete_date as date) >= '{fromDateStr}'
    AND CAST(Complete_date as date) <= '{toDateStr}'";
                }
                else
                {
                    // Specific project with batch filter
                    string batchListProcessed = ProcessBatchList();
                    query = $@"
WITH AuditCTE AS (
    SELECT
        Project_Code,
        Batch_Name,
        Line AS Audit_Type,
        Sample_Count,
        CASE Status
            WHEN  2 THEN 'Complete'
            WHEN  1 THEN 'Running'
            WHEN -2 THEN 'Canceled'
            WHEN  0 THEN 'Untouched'
            WHEN -1 THEN 'Suspended'
        END AS Status,
        COALESCE(operatorid, emp_id) AS Auditor_id,
        UserName,
        CONCAT(
            CAST(t.total_seconds / 3600 AS VARCHAR(10)),
            ':',
            RIGHT('0' + CAST((t.total_seconds % 3600) / 60 AS VARCHAR(2)), 2),
            ':',
            RIGHT('0' + CAST(t.total_seconds % 60 AS VARCHAR(2)), 2)
        ) AS Audit_Time,
        Start_DateTime AS Start_date,
        COALESCE(End_DateTime3, End_DateTime2, End_DateTime) AS Complete_date,
        comments
    FROM Audit_B
    CROSS APPLY (
        SELECT
            ISNULL(DATEDIFF(SECOND, Start_DateTime,  End_DateTime),  0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime2, End_DateTime2), 0) +
            ISNULL(DATEDIFF(SECOND, Start_DateTime3, End_DateTime3), 0)
            AS total_seconds
    ) t
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
WHERE
    Project_Code = @project_code
    AND batch_name in {batchListProcessed}
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

                if (dataGridView?.Rows.Count > 0)
                {
                    dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                }
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
        private void BtnAllProjectStatus_Click(object? sender, EventArgs e)
        {
            // Show menu items if user has permission (check if MMAdd is visible in main menu)
            if (mainMenuForm != null && mainMenuForm.mmAdd != null && mainMenuForm.mmAdd.Visible)
            {
                menuProjectComplete?.Visible = true;
                menuProjectStop?.Visible = true;
                menuProjectRunning?.Visible = true;
                menuUpdateProjectInfo?.Visible = true;
                menuUpdateBlank?.Visible = true;
            }

            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query = @"
SELECT
    p.Project_code,
    p.Total_Batch,
    COUNT(CASE WHEN b.Status = 2 THEN 1 END) AS Submit_Complete,
    p.Total_Batch - COUNT(CASE WHEN b.Status = 2 THEN 1 END) AS Submit_Due,
    COUNT(CASE WHEN b.Status = 1 THEN 1 END) AS Submit_Running,
    p.Language,
    p.Deadline,
    CASE
        WHEN p.Status = -2 THEN 'Stop'
        WHEN p.Status = -1 THEN 'Running'
        WHEN p.Status =  0 THEN ''
        WHEN p.Status =  1 THEN 'Complete'
    END AS Status,
    p.Que_Count,
    p.Que_Update_Date,
    p.Comments,
    STUFF((
        SELECT ', ' + b2.Username
        FROM Audit_B b2
        WHERE b2.Project_code = p.Project_code AND b2.Status = 1
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(1000)'), 1, 2, '') AS Submit_Running_Auditors
FROM
    Audit_B_Projects p
LEFT JOIN
    Audit_B b ON p.Project_code = b.Project_code
GROUP BY
    p.Project_code,
    p.Total_Batch,
    p.Language,
    p.Deadline,
    p.Status,
    p.Que_Count,
    p.Que_Update_Date,
    p.Comments
ORDER BY
    CASE
        WHEN p.Status = -1 THEN 1
        WHEN p.Status =  0 THEN 2
        WHEN p.Status = -2 THEN 3
        WHEN p.Status =  1 THEN 4
    END,
    p.Deadline";

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                dataAdapter = new SqlDataAdapter(cmd);
                dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                dataGridView?.DataSource = dataTable;

                // Clear summary labels
                lblTotalSample?.Text = "Total Sample :";
                lblTotalBatch?.Text = "Total Batch :";
                lblTotalAuditor?.Text = "Total Auditors :";
                lblBatchComplete?.Text = "Batch Complete :";
                lblBatchRunning?.Text = "Batch Running :";
                lblBatchPending?.Text = "Batch Pending :";

                // Set column widths
                if (dataGridView?.Columns.Count >= 12)
                {
                    dataGridView.Columns[0].Width = 250;
                    dataGridView.Columns[1].Width = 100;
                    dataGridView.Columns[2].Width = 100;
                    dataGridView.Columns[3].Width = 100;
                    dataGridView.Columns[4].Width = 100;
                    dataGridView.Columns[5].Width = 200;
                    dataGridView.Columns[6].Width = 130;
                    dataGridView.Columns[7].Width = 100;
                    dataGridView.Columns[8].Width = 100;
                    dataGridView.Columns[9].Width = 130;
                    dataGridView.Columns[10].Width = 150;
                    dataGridView.Columns[11].Width = 250;
                }

                // Calculate total queue count
                double sum = 0;
                foreach (DataRow row in dataTable.Rows)
                {
                    if (row["Que_Count"] != DBNull.Value)
                        sum += Convert.ToDouble(row["Que_Count"]);
                }
                lblTotalQue?.Text = $"Total Que : {sum}";

                try
                {
                    string countQuery = @"
        SELECT 
            COUNT(DISTINCT COALESCE(operatorid, emp_id)) AS Total_Running_Auditors,
            COUNT(*) AS Total_Running_Batches
        FROM Audit_B
        WHERE Status = 1";

                    SqlCommand countCmd = new SqlCommand(countQuery, dbConnection);
                    SqlDataReader reader = countCmd.ExecuteReader();

                    if (reader.Read())
                    {
                        int runningAuditors = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                        int runningBatches = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);

                        lblTotalAuditor?.Text = $"Total Auditors : {runningAuditors}";
                        lblBatchRunning?.Text = $"Batch Running : {runningBatches}";
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    // If count query fails, keep default text
                    System.Diagnostics.Debug.WriteLine($"Error getting running counts: {ex.Message}");
                }


                if (dataGridView?.Rows.Count > 0)
                {
                    dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading project status: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection.Close();
            }
        }
        private void BtnQueUpdate_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtQueCount?.Text) || string.IsNullOrWhiteSpace(cbQueProjectList?.Text))
            {
                MessageBox.Show("Please select a project and enter queue count", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show("Are you sure you want to add que count to this project?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                int queCount = int.Parse(txtQueCount.Text);
                string projectCode = cbQueProjectList.Text;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                queryUpdate?.CommandText = @"UPDATE audit_b_projects 
                                   SET que_count = @que_count, 
                                       que_update_date = @que_update_date 
                                   WHERE project_code = @project_code";

                queryUpdate?.Parameters.Clear();
                queryUpdate?.Parameters.AddWithValue("@que_count", queCount);
                queryUpdate?.Parameters.AddWithValue("@project_code", projectCode);
                queryUpdate?.Parameters.AddWithValue("@que_update_date", DateTime.Today);
                queryUpdate?.ExecuteNonQuery();
                // Clear inputs
                txtQueCount.Text = "";
                cbQueProjectList.SelectedIndex = -1;

                // Refresh the project status view
                BtnAllProjectStatus_Click(sender, e);

                MessageBox.Show("Queue count updated successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (FormatException)
            {
                MessageBox.Show("Please enter a valid number for queue count", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating queue count: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }
        private void BtnDeleteBatches_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBatchList?.Text))
            {
                MessageBox.Show("Please enter batch list in format: BatchName[TAB]LineNumber",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validate format
            bool isValid = true;
            var lines = txtBatchList.Lines;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split('\t');
                if (parts.Length != 2)
                {
                    isValid = false;
                    break;
                }
            }

            if (!isValid)
            {
                MessageBox.Show("Format Error. Each line must be: BatchName[TAB]LineNumber",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show("Do you want to proceed with deletion?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
                return;

            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var parts = line.Split('\t');
                    string batchName = parts[0].Trim();
                    int lineNo = int.Parse(parts[1].Trim());

                    // Insert into cancel log
                    queryUpdate?.CommandText = @"INSERT INTO audit_b_cancel
                                       (Batch_name, Line, Username, cancel_dateTime) 
                                       VALUES(@BatchName, @Status, @Username, @current_date)";

                    queryUpdate?.Parameters.Clear();
                    queryUpdate?.Parameters.AddWithValue("@BatchName", batchName);
                    queryUpdate?.Parameters.AddWithValue("@Status", lineNo);
                    queryUpdate?.Parameters.AddWithValue("@Username", mainMenuForm?.edtUsername?.Text ?? "");
                    queryUpdate?.Parameters.AddWithValue("@current_date", DateTime.Now);
                    queryUpdate?.ExecuteNonQuery();

                    // Delete from audit_b
                    queryUpdate?.CommandText = @"DELETE FROM audit_b 
                                       WHERE batch_name = @BatchName 
                                       AND line = @Status";

                    queryUpdate?.Parameters.Clear();
                    queryUpdate?.Parameters.AddWithValue("@BatchName", batchName);
                    queryUpdate?.Parameters.AddWithValue("@Status", lineNo);
                    queryUpdate?.ExecuteNonQuery();
                }

                MessageBox.Show("Batches Deleted", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtBatchList.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting batches: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }
        private void BtnBatchReset_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBatchList?.Text))
            {
                MessageBox.Show("Please enter batch list in format: BatchName[TAB]LineNumber",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validate format
            bool isValid = true;
            var lines = txtBatchList.Lines;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split('\t');
                if (parts.Length != 2)
                {
                    isValid = false;
                    break;
                }
            }

            if (!isValid)
            {
                MessageBox.Show("Format Error. Each line must be: BatchName[TAB]LineNumber",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show("Do you want to proceed with Batch Reset?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var parts = line.Split('\t');
                    string batchName = parts[0].Trim();
                    int lineNo = int.Parse(parts[1].Trim());

                    // Insert into cancel log
                    queryUpdate?.CommandText = @"INSERT INTO audit_b_cancel
                                       (Batch_name, Line, Username, cancel_dateTime) 
                                       VALUES(@BatchName, @Status, @Username, @current_date)";

                    queryUpdate?.Parameters.Clear();
                    queryUpdate?.Parameters.AddWithValue("@BatchName", batchName);
                    queryUpdate?.Parameters.AddWithValue("@Status", lineNo);
                    queryUpdate?.Parameters.AddWithValue("@Username", mainMenuForm?.edtUsername?.Text ?? "");
                    queryUpdate?.Parameters.AddWithValue("@current_date", DateTime.Now);
                    queryUpdate?.ExecuteNonQuery();

                    // Reset batch (set status to -2 and clear fields)
                    queryUpdate?.CommandText = @"UPDATE Audit_B 
                                       SET sample_count = NULL, 
                                           Status = -2, 
                                           OperatorID = NULL, 
                                           Emp_ID = NULL,
                                           UserName = NULL, 
                                           Start_DateTime = NULL, 
                                           End_DateTime = NULL, 
                                           Start_DateTime2 = NULL,
                                           End_DateTime2 = NULL, 
                                           Start_DateTime3 = NULL, 
                                           End_DateTime3 = NULL 
                                       WHERE batch_name = @batchName 
                                       AND line = @Status";

                    queryUpdate?.Parameters.Clear();
                    queryUpdate?.Parameters.AddWithValue("@BatchName", batchName);
                    queryUpdate?.Parameters.AddWithValue("@Status", lineNo);
                    queryUpdate?.ExecuteNonQuery();
                }

                MessageBox.Show("Batches reset complete", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtBatchList.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error resetting batches: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }
        // Context Menu Click Handlers

        private void MenuProjectComplete_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentRow == null)
                return;

            var result = MessageBox.Show("Are you sure you want to update the project as complete?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                string? projectCode = dataGridView.CurrentRow.Cells["Project_code"].Value?.ToString();

                if (string.IsNullOrEmpty(projectCode))
                    return;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                queryUpdate?.CommandText = @"UPDATE audit_b_Projects 
                                   SET status = 1 
                                   WHERE project_code = @project_code";

                queryUpdate?.Parameters.Clear();
                queryUpdate?.Parameters.AddWithValue("@project_code", projectCode);
                queryUpdate?.ExecuteNonQuery();

                MessageBox.Show("Project status updated to Complete", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating project status: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }

        private void MenuProjectStop_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentRow == null)
                return;

            var result = MessageBox.Show("Are you sure you want to update the project as Stop?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                string? projectCode = dataGridView.CurrentRow.Cells["Project_code"].Value?.ToString();

                if (string.IsNullOrEmpty(projectCode))
                    return;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                queryUpdate?.CommandText = @"UPDATE audit_b_Projects 
                                   SET status = -2 
                                   WHERE project_code = @project_code";

                queryUpdate?.Parameters.Clear();
                queryUpdate?.Parameters.AddWithValue("@project_code", projectCode);
                queryUpdate?.ExecuteNonQuery();

                MessageBox.Show("Project status updated to Stop", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating project status: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }

        private void MenuProjectRunning_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentRow == null)
                return;

            var result = MessageBox.Show("Are you sure you want to update the project as Running?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                string? projectCode = dataGridView.CurrentRow.Cells["Project_code"].Value?.ToString();

                if (string.IsNullOrEmpty(projectCode))
                    return;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                queryUpdate?.CommandText = @"UPDATE audit_b_Projects 
                                   SET status = -1 
                                   WHERE project_code = @project_code";

                queryUpdate?.Parameters.Clear();
                queryUpdate?.Parameters.AddWithValue("@project_code", projectCode);
                queryUpdate?.ExecuteNonQuery();

                MessageBox.Show("Project status updated to Running", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating project status: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }

        private void MenuUpdateBlank_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentRow == null)
                return;

            var result = MessageBox.Show("Are you sure you want to update the project as Blank?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                string? projectCode = dataGridView.CurrentRow.Cells["Project_code"].Value?.ToString();

                if (string.IsNullOrEmpty(projectCode))
                    return;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                queryUpdate?.CommandText = @"UPDATE audit_b_Projects 
                                   SET status = 0 
                                   WHERE project_code = @project_code";

                queryUpdate?.Parameters.Clear();
                queryUpdate?.Parameters.AddWithValue("@project_code", projectCode);
                queryUpdate?.ExecuteNonQuery();

                MessageBox.Show("Project status updated to Blank", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating project status: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (dbConnection?.State == ConnectionState.Open)
                    dbConnection?.Close();
            }
        }

        private void MenuUpdateProjectInfo_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentRow == null)
                return;

            try
            {
                // Populate update panel with current values
                txtProjectCode?.Text = dataGridView.CurrentRow.Cells["Project_code"].Value?.ToString() ?? "";
                txtTotalBatch?.Text = dataGridView.CurrentRow.Cells["Total_Batch"].Value?.ToString() ?? "";
                txtLanguage?.Text = dataGridView.CurrentRow.Cells["Language"].Value?.ToString() ?? "";
                if (dataGridView.CurrentRow.Cells["Deadline"].Value is DateTime dt)
                {
                    txtDeadline?.Text = dt.ToString("yyyy-MM-dd"); // or "dd-MM-yyyy"
                }
                txtComments?.Text = dataGridView.CurrentRow.Cells["Comments"].Value?.ToString() ?? "";
                // Show update panel
                pnlProjectUpdate?.Visible = true;
                pnlProjectUpdate?.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading project info: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DataGridView_ColumnHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView?.SelectionMode == DataGridViewSelectionMode.FullRowSelect)
            {
                dataGridView?.SelectionMode = DataGridViewSelectionMode.CellSelect;
            }
            else
            {
                dataGridView?.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            }
        }
        private void dataGridView_DataBindingComplete(object? sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewColumn col in dataGridView?.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void MenuShowAuditors_Click(object? sender, EventArgs e)
        {
            if (dataGridView?.CurrentRow == null)
                return;

            try
            {
                string? projectCode = dataGridView.CurrentRow.Cells["Project_code"].Value?.ToString();

                if (string.IsNullOrEmpty(projectCode))
                    return;

                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                string query = @"SELECT UserName 
                        FROM Audit_B 
                        WHERE Project_Code = @Project_Code 
                        AND Status = 1 
                        ORDER BY UserName";

                SqlCommand cmd = new SqlCommand(query, dbConnection);
                cmd.Parameters.AddWithValue("@Project_Code", projectCode);

                SqlDataReader reader = cmd.ExecuteReader();

                listBoxAuditors?.Items.Clear();
                while (reader.Read())
                {
                    string? userName = Convert.ToString(reader["UserName"]);

                    if (!string.IsNullOrWhiteSpace(userName))
                    {
                        listBoxAuditors?.Items.Add(userName);
                    }
                }
                reader.Close();

                pnlAuditorList?.Visible = true;
                pnlAuditorList?.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading auditors: {ex.Message}", "Error",
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

        private void BtnProjectUpdate_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to update the project status?",
                "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                string? projectCode = txtProjectCode?.Text;
                string? language = txtLanguage?.Text;
                string? deadline = txtDeadline?.Text;
                int totalBatch = int.Parse(txtTotalBatch?.Text ?? "0");
                string? comments = txtComments?.Text;
                if (dbConnection?.State != ConnectionState.Open)
                    dbConnection?.Open();

                queryUpdate?.CommandText = @"UPDATE audit_b_Projects 
                                   SET total_batch = @total_batch, 
                                       language = @language, 
                                       deadline = @deadline,
                                       comments = @comments 
                                   WHERE project_code = @project_code";

                queryUpdate?.Parameters.Clear();
                queryUpdate?.Parameters.AddWithValue("@project_code", projectCode);
                queryUpdate?.Parameters.AddWithValue("@total_batch", totalBatch);
                queryUpdate?.Parameters.AddWithValue("@language", language);
                queryUpdate?.Parameters.AddWithValue("@deadline", deadline);
                queryUpdate?.Parameters.AddWithValue("@comments", comments);
                queryUpdate?.ExecuteNonQuery();

                pnlProjectUpdate?.Visible = false;

                MessageBox.Show("Project information updated successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating project info: {ex.Message}", "Error",
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