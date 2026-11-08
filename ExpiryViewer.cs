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
        TextBox txtSearch;
        NumericUpDown numMin, numMax;
        CheckBox chkStock;
        Button btnSearch, btnExport, btnConn;
        DataGridView grid;
        Label lblCount, lblStatus;
        DataTable data;

        void UpdateTitle() { Text = "조제의약품 유효기간 조회  -  [" + Conn.Server + (Conn.Auth == "SQL" ? " / SQL인증" : " / Windows인증") + "]"; }

        public MainForm()
        {
            Text = "조제의약품 유효기간 조회  -  PM2014";
            Width = 1180; Height = 720;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("맑은 고딕", 9.5f);

            var panel = new Panel { Dock = DockStyle.Top, Height = 84, Padding = new Padding(10) };
            Controls.Add(panel);

            btnConn = new Button { Text = "접속설정", Top = 10, Width = 90, Height = 26, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnConn.Left = 1148 - 90 - 12;
            btnConn.Click += (s, e) => { using (var f = new SettingsForm()) { if (f.ShowDialog(this) == DialogResult.OK) { UpdateTitle(); DoSearch(); } } };

            var l2 = new Label { Text = "유효기간 남은 개월:", Left = 10, Top = 14, Width = 118 };
            numMin = new NumericUpDown { Left = 130, Top = 11, Width = 50, Minimum = 0, Maximum = 120, Value = 1 };
            var l3 = new Label { Text = "~", Left = 184, Top = 14, Width = 14 };
            numMax = new NumericUpDown { Left = 200, Top = 11, Width = 50, Minimum = 0, Maximum = 120, Value = 12 };
            var l3b = new Label { Text = "개월", Left = 254, Top = 14, Width = 34 };

            chkStock = new CheckBox { Text = "현재고 0 제외", Left = 296, Top = 13, Width = 110, Checked = true };

            var l4 = new Label { Text = "검색(약품명/위치)", Left = 10, Top = 50, Width = 110 };
            txtSearch = new TextBox { Left = 120, Top = 47, Width = 260 };
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { DoSearch(); e.SuppressKeyPress = true; } };

            btnSearch = new Button { Text = "조 회", Left = 392, Top = 45, Width = 90, Height = 28 };
            btnSearch.Click += (s, e) => DoSearch();
            btnExport = new Button { Text = "엑셀(CSV) 저장", Left = 490, Top = 45, Width = 120, Height = 28 };
            btnExport.Click += (s, e) => DoExport();

            lblCount = new Label { Left = 630, Top = 50, Width = 200, Text = "", ForeColor = Color.DarkBlue, Font = new Font("맑은 고딕", 10f, FontStyle.Bold) };

            panel.Controls.AddRange(new Control[] { btnConn, l2, numMin, l3, numMax, l3b, chkStock, l4, txtSearch, btnSearch, btnExport, lblCount });

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
            UpdateTitle();
            Load += (s, e) => DoSearch();
        }

        string ConnStr() { return Conn.ConnStr(); }

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
            Conn.Load();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    static class Conn
    {
        public static string Server = ".\\PM2014";
        public static string Auth = "WIN";   // WIN or SQL
        public static string User = "";
        public static string Pw = "";
        public const string DB = "PM_DRUG";

        static string Dir { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PharmTool"); } }
        static string Cfg { get { return Path.Combine(Dir, "conn.cfg"); } }

        public static void Load()
        {
            try
            {
                if (!File.Exists(Cfg)) return;
                foreach (var line in File.ReadAllLines(Cfg))
                {
                    int i = line.IndexOf('='); if (i < 0) continue;
                    string k = line.Substring(0, i).Trim(), v = line.Substring(i + 1);
                    if (k == "Server") Server = v;
                    else if (k == "Auth") Auth = v;
                    else if (k == "User") User = v;
                    else if (k == "Pw") Pw = Unprotect(v);
                }
            }
            catch { }
        }
        public static void Save()
        {
            try
            {
                Directory.CreateDirectory(Dir);
                File.WriteAllText(Cfg, "Server=" + Server + "\r\nAuth=" + Auth + "\r\nUser=" + User + "\r\nPw=" + Protect(Pw) + "\r\n");
            }
            catch { }
        }
        public static string ConnStr()
        {
            if (Auth == "SQL")
                return "Server=" + Server + ";Database=" + DB + ";User ID=" + User + ";Password=" + Pw + ";TrustServerCertificate=True;Connect Timeout=10";
            return "Server=" + Server + ";Database=" + DB + ";Integrated Security=True;TrustServerCertificate=True;Connect Timeout=10";
        }
        static string Protect(string s)
        {
            try { if (string.IsNullOrEmpty(s)) return ""; var b = System.Security.Cryptography.ProtectedData.Protect(Encoding.UTF8.GetBytes(s), null, System.Security.Cryptography.DataProtectionScope.CurrentUser); return Convert.ToBase64String(b); }
            catch { return ""; }
        }
        static string Unprotect(string s)
        {
            try { if (string.IsNullOrEmpty(s)) return ""; var b = System.Security.Cryptography.ProtectedData.Unprotect(Convert.FromBase64String(s), null, System.Security.Cryptography.DataProtectionScope.CurrentUser); return Encoding.UTF8.GetString(b); }
            catch { return ""; }
        }
    }

    class SettingsForm : Form
    {
        TextBox tServer, tUser, tPw;
        ComboBox cAuth;
        public SettingsForm()
        {
            Text = "접속 설정"; Width = 430; Height = 280;
            FormBorderStyle = FormBorderStyle.FixedDialog; StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false; MinimizeBox = false; Font = new Font("맑은 고딕", 9.5f);

            Controls.Add(new Label { Text = "서버", Left = 20, Top = 24, Width = 70 });
            tServer = new TextBox { Left = 100, Top = 21, Width = 290, Text = Conn.Server };
            Controls.Add(tServer);
            Controls.Add(new Label { Text = "(예: DESKTOP-8VR1BS2\\PM2014)", Left = 100, Top = 46, Width = 290, ForeColor = Color.Gray });

            Controls.Add(new Label { Text = "인증방식", Left = 20, Top = 76, Width = 70 });
            cAuth = new ComboBox { Left = 100, Top = 73, Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            cAuth.Items.AddRange(new object[] { "Windows 인증", "SQL 인증(아이디/암호)" });
            cAuth.SelectedIndex = Conn.Auth == "SQL" ? 1 : 0;
            Controls.Add(cAuth);

            var lU = new Label { Text = "아이디", Left = 20, Top = 112, Width = 70 };
            tUser = new TextBox { Left = 100, Top = 109, Width = 180, Text = Conn.User };
            var lP = new Label { Text = "암호", Left = 20, Top = 148, Width = 70 };
            tPw = new TextBox { Left = 100, Top = 145, Width = 180, Text = Conn.Pw, UseSystemPasswordChar = true };
            Controls.AddRange(new Control[] { lU, tUser, lP, tPw });

            EventHandler sync = (s, e) => { bool sql = cAuth.SelectedIndex == 1; lU.Enabled = tUser.Enabled = lP.Enabled = tPw.Enabled = sql; };
            cAuth.SelectedIndexChanged += sync; sync(null, null);

            var bTest = new Button { Text = "연결 테스트", Left = 100, Top = 190, Width = 110, Height = 30 };
            bTest.Click += (s, e) => { Apply(); try { using (var cn = new SqlConnection(Conn.ConnStr())) { cn.Open(); } MessageBox.Show("연결 성공", "확인"); } catch (Exception ex) { MessageBox.Show("연결 실패:\r\n\r\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); } };
            var bOk = new Button { Text = "저장", Left = 225, Top = 190, Width = 78, Height = 30 };
            bOk.Click += (s, e) => { Apply(); Conn.Save(); DialogResult = DialogResult.OK; Close(); };
            var bCancel = new Button { Text = "취소", Left = 311, Top = 190, Width = 78, Height = 30, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { bTest, bOk, bCancel });
            AcceptButton = bOk; CancelButton = bCancel;
        }
        void Apply() { Conn.Server = tServer.Text.Trim(); Conn.Auth = cAuth.SelectedIndex == 1 ? "SQL" : "WIN"; Conn.User = tUser.Text.Trim(); Conn.Pw = tPw.Text; }
    }
}
