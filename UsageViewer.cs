using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace UsageViewer
{
    public class MainForm : Form
    {
        TextBox txtServer, txtSearch, txtClinic;
        ComboBox cboMode;
        DateTimePicker dtpFrom, dtpTo;
        Button btnToday, btnMonth, btnRun, btnExport;
        DataGridView grid;
        Label lblCount, lblStatus;
        DataTable data;
        const string DBNAME = "PM_DRUG";

        public MainForm()
        {
            Text = "조제 사용량 집계  -  PM2014";
            Width = 1180; Height = 720;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("맑은 고딕", 9.5f);

            var panel = new Panel { Dock = DockStyle.Top, Height = 88, Padding = new Padding(10) };
            Controls.Add(panel);

            var l1 = new Label { Text = "서버", Left = 10, Top = 14, Width = 34 };
            txtServer = new TextBox { Left = 46, Top = 11, Width = 100, Text = ".\\PM2014" };

            var l2 = new Label { Text = "집계단위", Left = 160, Top = 14, Width = 56 };
            cboMode = new ComboBox { Left = 216, Top = 11, Width = 90, DropDownStyle = ComboBoxStyle.DropDownList };
            cboMode.Items.AddRange(new object[] { "약품별", "시간별", "요일별", "월별" });
            cboMode.SelectedIndex = 0;

            var l3 = new Label { Text = "기간", Left = 320, Top = 14, Width = 34 };
            dtpFrom = new DateTimePicker { Left = 356, Top = 11, Width = 110, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            var l4 = new Label { Text = "~", Left = 470, Top = 14, Width = 14 };
            dtpTo = new DateTimePicker { Left = 486, Top = 11, Width = 110, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtpFrom.Value = DateTime.Today; dtpTo.Value = DateTime.Today;

            btnToday = new Button { Text = "오늘", Left = 604, Top = 10, Width = 56, Height = 25 };
            btnToday.Click += (s, e) => { dtpFrom.Value = DateTime.Today; dtpTo.Value = DateTime.Today; DoRun(); };
            btnMonth = new Button { Text = "이번달", Left = 664, Top = 10, Width = 60, Height = 25 };
            btnMonth.Click += (s, e) => { var n = DateTime.Today; dtpFrom.Value = new DateTime(n.Year, n.Month, 1); dtpTo.Value = n; DoRun(); };

            var l5 = new Label { Text = "약품명", Left = 10, Top = 52, Width = 48 };
            txtSearch = new TextBox { Left = 58, Top = 49, Width = 150 };
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { DoRun(); e.SuppressKeyPress = true; } };

            var l6 = new Label { Text = "의원명", Left = 216, Top = 52, Width = 48 };
            txtClinic = new TextBox { Left = 264, Top = 49, Width = 150 };
            txtClinic.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { DoRun(); e.SuppressKeyPress = true; } };

            btnRun = new Button { Text = "집 계", Left = 426, Top = 47, Width = 84, Height = 28, Font = new Font("맑은 고딕", 10f, FontStyle.Bold) };
            btnRun.Click += (s, e) => DoRun();
            btnExport = new Button { Text = "엑셀(CSV) 저장", Left = 518, Top = 47, Width = 118, Height = 28 };
            btnExport.Click += (s, e) => DoExport();

            lblCount = new Label { Left = 648, Top = 52, Width = 260, Text = "", ForeColor = Color.DarkBlue, Font = new Font("맑은 고딕", 10f, FontStyle.Bold) };

            panel.Controls.AddRange(new Control[] { l1, txtServer, l2, cboMode, l3, dtpFrom, l4, dtpTo, btnToday, btnMonth, l5, txtSearch, l6, txtClinic, btnRun, btnExport, lblCount });

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                RowHeadersVisible = false,
                AllowUserToResizeRows = false,
                BackgroundColor = Color.White,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 34,
                EnableHeadersVisualStyles = false
            };
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 10f, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 120, 90);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(60, 120, 90);
            grid.DataBindingComplete += (s, e) => FormatGrid();
            Controls.Add(grid);

            lblStatus = new Label { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Text = "준비됨", ForeColor = Color.DimGray };
            Controls.Add(lblStatus);

            // 도킹 순서 보정(헤더가 패널에 가려지지 않도록): Fill 그리드를 맨 앞으로
            panel.SendToBack();
            lblStatus.SendToBack();
            grid.BringToFront();

            AcceptButton = btnRun;
            Load += (s, e) => DoRun();
        }

        string ConnStr()
        {
            return "Server=" + txtServer.Text.Trim() + ";Database=" + DBNAME +
                   ";Integrated Security=True;TrustServerCertificate=True;Connect Timeout=10";
        }

        string BuildSql(string mode)
        {
            string filter = @"m.InsertTime >= @from AND m.InsertTime < @toEnd
   AND (@search = '' OR s.DrugCode IN (SELECT DrugCode FROM CD_GOODS WHERE GoodsName LIKE '%'+@search+'%'))
   AND (@clinic = '' OR m.OrderName LIKE '%'+@clinic+'%')";

            if (mode == "약품별")
                return @"
SET NOCOUNT ON;
WITH g AS (SELECT DrugCode, MAX(GoodsName) GoodsName, MAX(SplName) SplName FROM CD_GOODS GROUP BY DrugCode)
SELECT
   ISNULL(g.GoodsName, s.DrugCode) AS 약품명,
   ISNULL(NULLIF(NULLIF(p.CD_NM_sale,'X'),'0'),'') AS 약품위치,
   ISNULL(m.OrderName,'') AS 처방발행의원,
   CONVERT(int, ISNULL(t.IM_QT_sale_debit,0)) AS 현재고,
   CAST(SUM(s.INV_QUAN) AS decimal(18,1)) AS 사용량,
   COUNT(DISTINCT s.PreSerial) AS 처방건수,
   ISNULL(g.SplName,'') AS 제조사,
   s.DrugCode AS 약품코드
FROM PM_PRES.dbo.PS_sub_pharm s
JOIN PM_PRES.dbo.PS_main m ON m.PreSerial = s.PreSerial
LEFT JOIN g ON g.DrugCode = s.DrugCode
LEFT JOIN CD_item_position p ON p.DrugCode = s.DrugCode
LEFT JOIN IM_total t ON t.DrugCode = s.DrugCode
WHERE " + filter + @"
GROUP BY s.DrugCode, g.GoodsName, p.CD_NM_sale, m.OrderName, t.IM_QT_sale_debit, g.SplName
HAVING SUM(s.INV_QUAN) <> 0
ORDER BY 사용량 DESC;";

            if (mode == "시간별")
                return @"
SET NOCOUNT ON;
SELECT
   RIGHT('0'+CAST(DATEPART(hour,m.InsertTime) AS varchar),2)+'시' AS 시간대,
   COUNT(DISTINCT m.PreSerial) AS 처방수,
   CAST(SUM(s.INV_QUAN) AS decimal(18,1)) AS 사용량합
FROM PM_PRES.dbo.PS_sub_pharm s
JOIN PM_PRES.dbo.PS_main m ON m.PreSerial = s.PreSerial
WHERE " + filter + @"
GROUP BY DATEPART(hour,m.InsertTime)
ORDER BY 시간대;";

            if (mode == "요일별")
                return @"
SET NOCOUNT ON;
SELECT
   DATENAME(weekday,m.InsertTime) AS 요일,
   COUNT(DISTINCT m.PreSerial) AS 처방수,
   CAST(SUM(s.INV_QUAN) AS decimal(18,1)) AS 사용량합
FROM PM_PRES.dbo.PS_sub_pharm s
JOIN PM_PRES.dbo.PS_main m ON m.PreSerial = s.PreSerial
WHERE " + filter + @"
GROUP BY DATEPART(weekday,m.InsertTime), DATENAME(weekday,m.InsertTime)
ORDER BY DATEPART(weekday,m.InsertTime);";

            // 월별
            return @"
SET NOCOUNT ON;
SELECT
   CONVERT(char(7),m.InsertTime,120) AS 년월,
   COUNT(DISTINCT m.PreSerial) AS 처방수,
   CAST(SUM(s.INV_QUAN) AS decimal(18,1)) AS 사용량합
FROM PM_PRES.dbo.PS_sub_pharm s
JOIN PM_PRES.dbo.PS_main m ON m.PreSerial = s.PreSerial
WHERE " + filter + @"
GROUP BY CONVERT(char(7),m.InsertTime,120)
ORDER BY 년월;";
        }

        void DoRun()
        {
            string mode = cboMode.SelectedItem.ToString();
            DateTime from = dtpFrom.Value.Date;
            DateTime toEnd = dtpTo.Value.Date.AddDays(1);
            if (toEnd <= from) { toEnd = from.AddDays(1); }
            lblStatus.Text = "집계 중...";
            Cursor = Cursors.WaitCursor;
            try
            {
                var dt = new DataTable();
                using (var cn = new SqlConnection(ConnStr()))
                using (var cmd = new SqlCommand(BuildSql(mode), cn))
                {
                    cmd.CommandTimeout = 120;
                    cmd.Parameters.AddWithValue("@from", from);
                    cmd.Parameters.AddWithValue("@toEnd", toEnd);
                    cmd.Parameters.AddWithValue("@search", txtSearch.Text.Trim());
                    cmd.Parameters.AddWithValue("@clinic", txtClinic.Text.Trim());
                    cn.Open();
                    using (var da = new SqlDataAdapter(cmd)) da.Fill(dt);
                }
                data = dt;
                grid.DataSource = dt;
                lblCount.Text = "총 " + dt.Rows.Count.ToString("N0") + " 행";
                lblStatus.Text = string.Format("[{0}] {1:yyyy-MM-dd} ~ {2:yyyy-MM-dd}  집계완료 ({3:HH:mm:ss})",
                    mode, from, dtpTo.Value.Date, DateTime.Now);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "오류";
                MessageBox.Show("집계 오류:\r\n\r\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { Cursor = Cursors.Default; }
        }

        void FormatGrid()
        {
            if (grid.Columns.Count == 0) return;
            foreach (DataGridViewColumn c in grid.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.Automatic;
                string n = c.Name;
                if (n == "약품명") c.Width = 330;
                else if (n == "약품위치") { c.Width = 150; c.DefaultCellStyle.Font = new Font("맑은 고딕", 9.5f, FontStyle.Bold); }
                else if (n == "처방발행의원") c.Width = 170;
                else if (n == "현재고") { c.Width = 80; c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; c.DefaultCellStyle.Format = "N0"; }
                else if (n == "사용량" || n == "사용량합") { c.Width = 100; c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; c.DefaultCellStyle.Format = "N1"; c.DefaultCellStyle.Font = new Font("맑은 고딕", 9.5f, FontStyle.Bold); c.DefaultCellStyle.ForeColor = Color.FromArgb(0, 90, 60); }
                else if (n == "처방건수" || n == "처방수") { c.Width = 80; c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; c.DefaultCellStyle.Format = "N0"; }
                else if (n == "제조사") c.Width = 150;
                else if (n == "약품코드") c.Width = 90;
                else if (n == "시간대" || n == "요일" || n == "년월") { c.Width = 110; c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; }
            }
        }

        void DoExport()
        {
            if (data == null || data.Rows.Count == 0) { MessageBox.Show("저장할 데이터가 없습니다."); return; }
            using (var sfd = new SaveFileDialog { Filter = "CSV 파일|*.csv", FileName = "사용량집계_" + cboMode.SelectedItem + "_" + DateTime.Now.ToString("yyyyMMdd") + ".csv" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                var sb = new StringBuilder();
                for (int c = 0; c < data.Columns.Count; c++) { if (c > 0) sb.Append(','); sb.Append(Q(data.Columns[c].ColumnName)); }
                sb.AppendLine();
                foreach (DataRow row in data.Rows)
                {
                    for (int c = 0; c < data.Columns.Count; c++) { if (c > 0) sb.Append(','); sb.Append(Q(row[c].ToString())); }
                    sb.AppendLine();
                }
                File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                lblStatus.Text = "저장 완료: " + sfd.FileName;
                MessageBox.Show("저장되었습니다.\r\n" + sfd.FileName, "완료");
            }
        }
        static string Q(string s) { s = s ?? ""; if (s.Contains(",") || s.Contains("\"") || s.Contains("\n")) return "\"" + s.Replace("\"", "\"\"") + "\""; return s; }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
