using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ExpiryViewer
{
    public class MainForm : Form
    {
        TextBox txtServer, txtSearch;
        NumericUpDown numMin, numMax;
        CheckBox chkStock;
        Button btnSearch, btnExport;
        DataGridView grid;
        Label lblCount, lblStatus;
        DataTable data;

        const string DBNAME = "PM_DRUG";

        public MainForm()
        {
            Text = "조제의약품 유효기간 조회  -  PM2014";
            Width = 1180; Height = 720;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("맑은 고딕", 9.5f);

            var panel = new Panel { Dock = DockStyle.Top, Height = 84, Padding = new Padding(10) };
            Controls.Add(panel);

            var l1 = new Label { Text = "서버", Left = 10, Top = 14, Width = 34 };
            txtServer = new TextBox { Left = 46, Top = 11, Width = 110, Text = ".\\PM2014" };

            var l2 = new Label { Text = "유효기간 남은 개월:", Left = 170, Top = 14, Width = 118 };
            numMin = new NumericUpDown { Left = 290, Top = 11, Width = 50, Minimum = 0, Maximum = 120, Value = 1 };
            var l3 = new Label { Text = "~", Left = 344, Top = 14, Width = 14 };
            numMax = new NumericUpDown { Left = 360, Top = 11, Width = 50, Minimum = 0, Maximum = 120, Value = 12 };
            var l3b = new Label { Text = "개월", Left = 414, Top = 14, Width = 34 };

            chkStock = new CheckBox { Text = "현재고 0 제외", Left = 456, Top = 13, Width = 110, Checked = true };

            var l4 = new Label { Text = "검색(약품명/위치)", Left = 10, Top = 50, Width = 110 };
            txtSearch = new TextBox { Left = 120, Top = 47, Width = 260 };
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { DoSearch(); e.SuppressKeyPress = true; } };

            btnSearch = new Button { Text = "조 회", Left = 392, Top = 45, Width = 90, Height = 28 };
            btnSearch.Click += (s, e) => DoSearch();
            btnExport = new Button { Text = "엑셀(CSV) 저장", Left = 490, Top = 45, Width = 120, Height = 28 };
            btnExport.Click += (s, e) => DoExport();

            lblCount = new Label { Left = 630, Top = 50, Width = 200, Text = "", ForeColor = Color.DarkBlue, Font = new Font("맑은 고딕", 10f, FontStyle.Bold) };

            panel.Controls.AddRange(new Control[] { l1, txtServer, l2, numMin, l3, numMax, l3b, chkStock, l4, txtSearch, btnSearch, btnExport, lblCount });

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
                BackgroundColor = Color.White
            };
            grid.ColumnHeadersVisible = true;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = 34;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 10f, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(70, 110, 170);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(70, 110, 170);
            grid.DataBindingComplete += (s, e) => FormatGrid();
            Controls.Add(grid);

            lblStatus = new Label { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Text = "준비됨", ForeColor = Color.DimGray };
            Controls.Add(lblStatus);

            // 도킹 순서 보정: 이 폼은 Fill(그리드)이 맨 앞(마지막 배치)이어야
            // 패널/상태줄이 먼저 영역을 차지하고 그리드가 남은 가운데를 채움(헤더가 안 가려짐)
            panel.SendToBack();
            lblStatus.SendToBack();
            grid.BringToFront();

            AcceptButton = btnSearch;
            Load += (s, e) => DoSearch();
        }

        string ConnStr()
        {
            return "Server=" + txtServer.Text.Trim() + ";Database=" + DBNAME +
                   ";Integrated Security=True;TrustServerCertificate=True;Connect Timeout=10";
        }

        const string SQL = @"
SET NOCOUNT ON;
DECLARE @today date = CAST(GETDATE() AS date);
DECLARE @lo date = DATEADD(month, @minM, @today);
DECLARE @hi date = DATEADD(month, @maxM, @today);
WITH expd AS (
  SELECT DrugCode, MIN(d) NearExp
  FROM (SELECT DrugCode, TRY_CONVERT(date, WH_END_validity) d FROM WH_sub) x
  WHERE x.d IS NOT NULL AND x.d >= @today
  GROUP BY DrugCode
),
g AS (SELECT DrugCode, MAX(GoodsName) GoodsName, MAX(SplName) SplName FROM CD_GOODS GROUP BY DrugCode)
SELECT
   g.GoodsName AS 약품명,
   ISNULL(NULLIF(NULLIF(p.CD_NM_sale,'X'),'0'),'') AS 약품위치,
   CONVERT(char(10), e.NearExp, 23) AS 유효기간,
   DATEDIFF(month, @today, e.NearExp) AS 남은개월,
   CONVERT(int, ISNULL(t.IM_QT_sale_debit,0)) AS 현재고,
   ISNULL(lot.WH_LOT_NO,'') AS 로트,
   g.SplName AS 제조사,
   e.DrugCode AS 약품코드
FROM expd e
JOIN g ON g.DrugCode = e.DrugCode
LEFT JOIN CD_item_position p ON p.DrugCode = e.DrugCode
LEFT JOIN IM_total t ON t.DrugCode = e.DrugCode
OUTER APPLY (SELECT TOP 1 WH_LOT_NO FROM WH_sub w
             WHERE w.DrugCode=e.DrugCode AND TRY_CONVERT(date,w.WH_END_validity)=e.NearExp
               AND ISNULL(w.WH_LOT_NO,'')<>'') lot
WHERE e.NearExp BETWEEN @lo AND @hi
  AND (@excludeZero=0 OR ISNULL(t.IM_QT_sale_debit,0) > 0)
  AND (@search='' OR g.GoodsName LIKE '%'+@search+'%' OR ISNULL(p.CD_NM_sale,'') LIKE '%'+@search+'%')
ORDER BY e.NearExp, g.GoodsName;";

        void DoSearch()
        {
            int mn = (int)numMin.Value, mx = (int)numMax.Value;
            if (mn > mx) { var tmp = mn; mn = mx; mx = tmp; numMin.Value = mn; numMax.Value = mx; }
            lblStatus.Text = "조회 중...";
            Cursor = Cursors.WaitCursor;
            try
            {
                var dt = new DataTable();
                using (var cn = new SqlConnection(ConnStr()))
                using (var cmd = new SqlCommand(SQL, cn))
                {
                    cmd.CommandTimeout = 60;
                    cmd.Parameters.AddWithValue("@minM", mn);
                    cmd.Parameters.AddWithValue("@maxM", mx);
                    cmd.Parameters.AddWithValue("@excludeZero", chkStock.Checked ? 1 : 0);
                    cmd.Parameters.AddWithValue("@search", txtSearch.Text.Trim());
                    cn.Open();
                    using (var da = new SqlDataAdapter(cmd)) da.Fill(dt);
                }
                data = dt;
                grid.DataSource = dt;
                lblCount.Text = "총 " + dt.Rows.Count.ToString("N0") + " 건";
                lblStatus.Text = "조회 완료  (" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ")";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "오류";
                MessageBox.Show("조회 오류:\r\n\r\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { Cursor = Cursors.Default; }
        }

        void FormatGrid()
        {
            if (grid.Columns.Count == 0) return;
            int[] w = { 330, 150, 90, 70, 70, 110, 160, 90 };
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                if (i < w.Length) grid.Columns[i].Width = w[i];
                grid.Columns[i].SortMode = DataGridViewColumnSortMode.Automatic;
            }
            if (grid.Columns.Contains("남은개월")) grid.Columns["남은개월"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (grid.Columns.Contains("현재고")) { grid.Columns["현재고"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; grid.Columns["현재고"].DefaultCellStyle.Format = "N0"; }
            if (grid.Columns.Contains("유효기간")) grid.Columns["유효기간"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (grid.Columns.Contains("약품위치")) grid.Columns["약품위치"].DefaultCellStyle.Font = new Font("맑은 고딕", 9.5f, FontStyle.Bold);
            foreach (DataGridViewRow r in grid.Rows)
            {
                var v = r.Cells["남은개월"].Value;
                if (v != null && v != DBNull.Value)
                {
                    int m = Convert.ToInt32(v);
                    if (m <= 3) r.DefaultCellStyle.BackColor = Color.MistyRose;
                    else if (m <= 6) r.DefaultCellStyle.BackColor = Color.LightYellow;
                }
            }
        }

        void DoExport()
        {
            if (data == null || data.Rows.Count == 0) { MessageBox.Show("저장할 데이터가 없습니다."); return; }
            using (var sfd = new SaveFileDialog { Filter = "CSV 파일|*.csv", FileName = "유효기간조회_" + DateTime.Now.ToString("yyyyMMdd") + ".csv" })
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
