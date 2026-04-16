using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class OpenCloseGroupPivotForm : Form
    {
        private Button btnLoad;
        private Button btnCopy;
        private Label lblFile;
        private Label lblHint;
        private DataGridView dgv;
        private OpenFileDialog ofd;
        public OpenCloseGroupPivotForm()
        {
            BuildUi();
        }
        private void BuildUi()
        {
            Text = "APAMA / APTURA 날짜별 Open / Close 집계";
            Width = 1300;
            Height = 820;
            btnLoad = new Button
            {
                Text = "엑셀 파일 불러오기",
                Left = 15,
                Top = 15,
                Width = 180,
                Height = 32
            };
            btnLoad.Click += BtnLoad_Click;
            btnCopy = new Button
            {
                Text = "데이터 복사",
                Left = 205,
                Top = 15,
                Width = 140,
                Height = 32
            };
            btnCopy.Click += BtnCopy_Click;
            lblFile = new Label
            {
                Left = 360,
                Top = 22,
                Width = 850,
                Text = "파일: (없음)"
            };
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1200,
                AutoSize = false,
                Text = "C열(Open 날짜), D열(장비명), H열(Close 날짜)를 기준으로 APAMA / APTURA 그룹별 날짜별 개수를 집계합니다."
            };
            dgv = new DataGridView
            {
                Left = 15,
                Top = 80,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 95,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = true,
                EditMode = DataGridViewEditMode.EditProgrammatically,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                ScrollBars = ScrollBars.Both,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false
            };
            dgv.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) e.SuppressKeyPress = true;
                if ((e.Control && e.KeyCode == Keys.C) || (e.Control && e.KeyCode == Keys.Insert)) e.SuppressKeyPress = true;
            };
            ofd = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx;*.xls",
                Title = "대상 엑셀 파일 선택"
            };
            Controls.Add(btnLoad);
            Controls.Add(btnCopy);
            Controls.Add(lblFile);
            Controls.Add(lblHint);
            Controls.Add(dgv);
        }
        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;
            try
            {
                var path = ofd.FileName;
                lblFile.Text = "파일: " + System.IO.Path.GetFileName(path);
                var dt = BuildGroupPivotTable(path);
                dgv.DataSource = null;
                dgv.Columns.Clear();
                dgv.AutoGenerateColumns = true;
                dgv.DataSource = dt;
                for (int c = 1; c < dgv.Columns.Count; c++)
                    dgv.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("집계할 데이터가 없습니다.",
                        "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message,
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyGridForExcel();
        }
        private DataTable BuildGroupPivotTable(string path)
        {
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheet("Issue");
                int headerRow = FindHeaderRow(ws);
                if (headerRow <= 0)
                    throw new Exception("헤더 행을 찾지 못했습니다.");
                int dataStartRow = headerRow + 1;
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? dataStartRow;
                var apamaOpen = new Dictionary<DateTime, int>();
                var apamaClose = new Dictionary<DateTime, int>();
                var apturaOpen = new Dictionary<DateTime, int>();
                var apturaClose = new Dictionary<DateTime, int>();
                var allDates = new HashSet<DateTime>();
                for (int r = dataStartRow; r <= lastRow; r++)
                {
                    string equip = GetCellString(ws.Cell(r, 4)).Trim(); // D
                    if (string.IsNullOrWhiteSpace(equip)) continue;
                    string group = GetGroupName(equip);
                    if (string.IsNullOrWhiteSpace(group)) continue;
                    DateTime? openDate = TryReadDate(ws.Cell(r, 3));  // C
                    DateTime? closeDate = TryReadDate(ws.Cell(r, 8)); // H
                    if (openDate.HasValue)
                    {
                        allDates.Add(openDate.Value.Date);
                        if (string.Equals(group, "APAMA", StringComparison.OrdinalIgnoreCase))
                            AddCount(apamaOpen, openDate.Value.Date);
                        else if (string.Equals(group, "APTURA", StringComparison.OrdinalIgnoreCase))
                            AddCount(apturaOpen, openDate.Value.Date);
                    }
                    if (closeDate.HasValue)
                    {
                        allDates.Add(closeDate.Value.Date);
                        if (string.Equals(group, "APAMA", StringComparison.OrdinalIgnoreCase))
                            AddCount(apamaClose, closeDate.Value.Date);
                        else if (string.Equals(group, "APTURA", StringComparison.OrdinalIgnoreCase))
                            AddCount(apturaClose, closeDate.Value.Date);
                    }
                }
                var dateList = new List<DateTime>();
                if (allDates.Count > 0)
                {
                    var minDate = allDates.Min();
                    var maxDate = allDates.Max();
                    for (var d = minDate; d <= maxDate; d = d.AddDays(1))
                        dateList.Add(d);
                }
                // 출력물만 사진처럼 구성
                var dt = new DataTable();
                dt.Columns.Add("A");
                dt.Columns.Add("B");
                dt.Columns.Add("C");
                dt.Columns.Add("D");
                dt.Columns.Add("E");
                dt.Columns.Add("F");
                dt.Columns.Add("G");
                dt.Columns.Add("H");
                dt.Columns.Add("I");
                // 그룹명 행
                var row0 = dt.NewRow();
                row0["C"] = "APAMA";
                row0["H"] = "APTURA";
                dt.Rows.Add(row0);
                // 헤더 행
                var row1 = dt.NewRow();
                row1["B"] = "Date";
                row1["C"] = "Open count";
                row1["D"] = "Close count";
                row1["G"] = "Date";
                row1["H"] = "Open count";
                row1["I"] = "Close count";
                dt.Rows.Add(row1);
                // 데이터 행
                foreach (var d in dateList)
                {
                    var row = dt.NewRow();
                    row["B"] = d.ToString("yyyy-MM-dd");
                    row["C"] = GetCount(apamaOpen, d);
                    row["D"] = GetCount(apamaClose, d);
                    row["G"] = d.ToString("yyyy-MM-dd");
                    row["H"] = GetCount(apturaOpen, d);
                    row["I"] = GetCount(apturaClose, d);
                    dt.Rows.Add(row);
                }
                return dt;
            }
        }
        private int FindHeaderRow(IXLWorksheet ws)
        {
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            if (lastRow == 0) return -1;
            for (int r = 1; r <= lastRow; r++)
            {
                string c = GetCellString(ws.Cell(r, 3)).Trim();
                string d = GetCellString(ws.Cell(r, 4)).Trim();
                string h = GetCellString(ws.Cell(r, 8)).Trim();
                if (string.Equals(c, "Date", StringComparison.OrdinalIgnoreCase) &&
                    d.IndexOf("TCB", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return r;
                }
                // 샘플 구조 fallback
                if (string.Equals(c, "Date", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(h, "Implementation", StringComparison.OrdinalIgnoreCase))
                {
                    return r;
                }
            }
            // 샘플 파일 fallback
            return 16;
        }
        private string GetGroupName(string equip)
        {
            if (string.IsNullOrWhiteSpace(equip)) return string.Empty;
            if (equip.IndexOf("APAMA", StringComparison.OrdinalIgnoreCase) >= 0)
                return "APAMA";
            if (equip.IndexOf("APTURA", StringComparison.OrdinalIgnoreCase) >= 0)
                return "APTURA";
            return string.Empty;
        }
        private void AddCount(Dictionary<DateTime, int> dict, DateTime date)
        {
            int cur;
            if (!dict.TryGetValue(date, out cur)) cur = 0;
            dict[date] = cur + 1;
        }
        
        private int GetCount(Dictionary<DateTime, int> dict, DateTime d)
        {
            int v;
            if (!dict.TryGetValue(d, out v)) v = 0;
            return v;
        }

        private void AddGroupRow(DataTable dt, string groupName, Dictionary<DateTime, int> dict, List<DateTime> dateList)
        {
            var row = dt.NewRow();
            row[0] = groupName;
            for (int i = 0; i < dateList.Count; i++)
            {
                int v;
                if (!dict.TryGetValue(dateList[i], out v)) v = 0;
                row[i + 1] = v;
            }
            dt.Rows.Add(row);
        }
        private static string GetCellString(IXLCell cell)
        {
            if (cell == null) return string.Empty;
            switch (cell.DataType)
            {
                case XLDataType.Text:
                    return cell.GetString().Trim();
                case XLDataType.Number:
                    return cell.GetDouble().ToString(CultureInfo.InvariantCulture).Trim();
                case XLDataType.Boolean:
                    return cell.GetBoolean() ? "TRUE" : "FALSE";
                case XLDataType.DateTime:
                    return cell.GetDateTime().ToString("yyyy-MM-dd HH:mm:ss");
                case XLDataType.Blank:
                    return string.Empty;
                default:
                    var s = cell.GetString();
                    if (!string.IsNullOrEmpty(s)) return s.Trim();
                    return cell.Value.ToString().Trim();
            }
        }
        private static DateTime? TryReadDate(IXLCell cell)
        {
            if (cell == null) return null;
            if (cell.DataType == XLDataType.DateTime)
                return cell.GetDateTime().Date;
            if (cell.DataType == XLDataType.Number)
            {
                try { return DateTime.FromOADate(cell.GetDouble()).Date; }
                catch { }
            }
            var s = cell.GetString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            DateTime dt;
            if (DateTime.TryParse(s, out dt))
                return dt.Date;
            return null;
        }
        private void CopyGridForExcel()
        {
            if (dgv == null || dgv.Rows.Count == 0) return;
            var oldMode = dgv.ClipboardCopyMode;
            var oldMultiSelect = dgv.MultiSelect;
            var oldSelectionMode = dgv.SelectionMode;
            try
            {
                dgv.MultiSelect = true;
                dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
                dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                dgv.ClearSelection();
                dgv.SelectAll();
                var dataObj = dgv.GetClipboardContent();
                if (dataObj != null)
                {
                    Clipboard.Clear();
                    Clipboard.SetDataObject(dataObj, true);
                }
            }
            finally
            {
                dgv.ClearSelection();
                dgv.MultiSelect = oldMultiSelect;
                dgv.SelectionMode = oldSelectionMode;
                dgv.ClipboardCopyMode = oldMode;
            }
        }
    }
}