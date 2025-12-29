using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class AlarmPivotForm4 : Form
    {
        private Button btnLoad, btnCopy;
        private Label lblFile, lblHint;
        private TabControl tabSheets;
        private OpenFileDialog ofd;

        private const string ColAlarm = "Alarm Name";
        private const string ColCount = "Count";
        private const string ColMin = "Total Minutes";
        private const string ColMinPerCount = "Min/Count";

        public AlarmPivotForm4()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "Alarm 횟수/시간 합산(KTCB-장비별 MTBA)";

            Width = 1200;
            Height = 800;

            btnLoad = new Button { Text = "파일 불러오기 (.xlsx/.xls)", Left = 15, Top = 15, Width = 180, Height = 32 };
            btnLoad.Click += BtnLoad_Click;

            btnCopy = new Button { Text = "데이터 복사", Left = 205, Top = 15, Width = 150, Height = 32 };
            btnCopy.Click += BtnCopy_Click;

            lblFile = new Label { Left = 370, Top = 22, Width = 800, Text = "파일: (없음)" };
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1100,
                Text = "F=Alarm Name, G/H=Start/End DateTime, I=Duration Minutes. 1~4시트: Lists\\APAMA_ALID.txt, 5~6시트: Lists\\APTURA_ALID.txt",
                AutoSize = false
            };

            tabSheets = new TabControl
            {
                Left = 15,
                Top = 80,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 95,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            ofd = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx;*.xls",
                Title = "집계 대상 엑셀 파일 선택"
            };

            Controls.Add(btnLoad);
            Controls.Add(btnCopy);
            Controls.Add(lblFile);
            Controls.Add(lblHint);
            Controls.Add(tabSheets);
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var path = ofd.FileName;
                lblFile.Text = "파일: " + Path.GetFileName(path);

                tabSheets.TabPages.Clear();

                // ✅ 리스트 로드 (실행파일 루트\Lists\*.txt)
                var apama = LoadWhitelistFromRoot(@"Lists\APAMA_ALID.txt");
                var aptura = LoadWhitelistFromRoot(@"Lists\APTURA_ALID.txt");

                using (var wb = new XLWorkbook(path))
                {
                    int sheetCount = wb.Worksheets.Count;

                    for (int i = 1; i <= sheetCount; i++)
                    {
                        var ws = wb.Worksheet(i);

                        // 1~4 => APAMA, 5~6 => APTURA, 그 외는 필터 없음(안전)
                        HashSet<string> whitelist = null;
                        if (i >= 1 && i <= 4) whitelist = apama;
                        else if (i >= 5 && i <= 6) whitelist = aptura;

                        var dt = BuildSheetAlarmSummary(ws, whitelist);
                        
                        var grid = CreateGrid();
                        grid.DataSource = dt;

                        // 표시/정렬

                        if (grid.Columns.Contains(ColCount))
                            grid.Columns[ColCount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        if (grid.Columns.Contains(ColMin))
                            grid.Columns[ColMin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        if (grid.Columns.Contains(ColMinPerCount))
                            grid.Columns[ColMinPerCount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                        grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

                        var tab = new TabPage(i + ". " + ws.Name) { Padding = new Padding(0) };
                        tab.Controls.Add(grid);
                        tabSheets.TabPages.Add(tab);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyCurrentSheetSummary();
        }

        // =========================
        // 핵심: 시트별 Alarm 발생횟수 + Minutes 합산
        // =========================
        private sealed class AlarmAgg
        {
            public int Count;
            public double Minutes;
        }

        private DataTable BuildSheetAlarmSummary(IXLWorksheet ws, HashSet<string> whitelist)
        {
            var sums = new Dictionary<string, AlarmAgg>(StringComparer.OrdinalIgnoreCase);

            var lastRow = ws.LastRowUsed();
            var lastCell = ws.LastCellUsed();
            if (lastRow == null || lastCell == null)
                return MakeEmptyTable();

            int lastRowNo = lastRow.RowNumber();

            // 첫 행은 헤더, 데이터는 2행부터
            for (int r = 2; r <= lastRowNo; r++)
            {
                // F열(6): Alarm Name
                string alarm = GetCellString(ws.Cell(r, 6));
                if (string.IsNullOrWhiteSpace(alarm)) continue;
                alarm = alarm.Trim();

                // 화이트리스트 필터
                if (whitelist != null && whitelist.Count > 0 && !whitelist.Contains(alarm))
                    continue;

                // I열(9): Minutes
                double minutes = TryReadDouble(ws.Cell(r, 9));

                // I가 비어있거나 0이면(G/H로 계산)
                if (minutes <= 0)
                {
                    DateTime? start = TryReadDateTime(ws.Cell(r, 7)); // G
                    DateTime? end = TryReadDateTime(ws.Cell(r, 8));   // H
                    if (start.HasValue && end.HasValue)
                    {
                        var diff = end.Value - start.Value;
                        if (diff.TotalMinutes < 0) diff = diff + TimeSpan.FromDays(1);
                        minutes = diff.TotalMinutes;
                    }
                }

                // 유효한 데이터만 카운트/합산
                if (minutes <= 0) continue;

                AlarmAgg agg;
                if (!sums.TryGetValue(alarm, out agg))
                {
                    agg = new AlarmAgg();
                    sums[alarm] = agg;
                }

                agg.Count += 1;
                agg.Minutes += minutes;
            }

            var dt = new DataTable();
            dt.Columns.Add(ColAlarm);
            dt.Columns.Add(ColCount, typeof(int));
            dt.Columns.Add(ColMin, typeof(double));
            dt.Columns.Add(ColMinPerCount, typeof(double));

            // 정렬: 총 분 내림차순 → Count 내림차순 → 알람명
            foreach (var kv in sums
                .OrderByDescending(x => x.Value.Minutes)
                .ThenByDescending(x => x.Value.Count)
                .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                var alarm = kv.Key;
                var agg = kv.Value;

                if (agg.Count <= 0) continue;

                double totalMin = agg.Minutes;
                double minPer = totalMin / Math.Max(1, agg.Count);

                var row = dt.NewRow();
                row[ColAlarm] = alarm;
                row[ColCount] = agg.Count;
                row[ColMin] = Math.Round(totalMin, 0);          // 분 합계는 정수 느낌으로
                row[ColMinPerCount] = Math.Round(minPer, 1);    // 분/횟수는 소수 1자리
                dt.Rows.Add(row);
            }

            return dt;
        }

        private DataTable MakeEmptyTable()
        {
            var dt = new DataTable();
            dt.Columns.Add(ColAlarm);
            dt.Columns.Add(ColCount, typeof(int));
            dt.Columns.Add(ColMin, typeof(double));
            dt.Columns.Add(ColMinPerCount, typeof(double));
            return dt;
        }

        // =========================
        // Lists 로드
        // =========================
        private HashSet<string> LoadWhitelistFromRoot(string relativePath)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var fullPath = Path.Combine(baseDir, relativePath);

            if (!File.Exists(fullPath))
                throw new FileNotFoundException("화이트리스트 파일을 찾을 수 없습니다: " + fullPath);

            var hs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in File.ReadAllLines(fullPath))
            {
                var s = (line ?? "").Trim();
                if (s.Length == 0) continue;
                hs.Add(s);
            }

            return hs;
        }

        // =========================
        // Copy (헤더 없이, 버튼으로만)
        // =========================
        private void CopyCurrentSheetSummary()
        {
            var grid = GetCurrentGrid();
            if (grid == null) return;

            var dt = grid.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;

            // Alarm Name + Count + Total Minutes + Min/Count (헤더 제외)
            var sb = new StringBuilder();
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];

                sb.Append(Convert.ToString(row[ColAlarm]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColCount]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColMin]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColMinPerCount]));
                sb.Append('\n');
            }

            Clipboard.Clear();
            Clipboard.SetText(sb.ToString());
        }

        private DataGridView GetCurrentGrid()
        {
            if (tabSheets.TabPages.Count == 0) return null;
            var tab = tabSheets.SelectedTab;
            if (tab == null || tab.Controls.Count == 0) return null;
            return tab.Controls[0] as DataGridView;
        }

        // =========================
        // Grid 생성 (읽기전용/CTRL+C 차단)
        // =========================
        private DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                BorderStyle = BorderStyle.None,

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

            grid.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) e.SuppressKeyPress = true;
                if ((e.Control && e.KeyCode == Keys.C) || (e.Control && e.KeyCode == Keys.Insert)) e.SuppressKeyPress = true;
            };

            return grid;
        }

        // =========================
        // Cell Utils
        // =========================
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

        private static double TryReadDouble(IXLCell cell)
        {
            if (cell == null) return 0;

            try
            {
                if (cell.DataType == XLDataType.Number)
                    return cell.GetDouble();

                var s = cell.GetString().Trim();
                if (string.IsNullOrEmpty(s)) return 0;

                double d;
                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;

                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                    return d;

            }
            catch { }

            return 0;
        }

        private static DateTime? TryReadDateTime(IXLCell cell)
        {
            if (cell == null) return null;

            try
            {
                if (cell.DataType == XLDataType.DateTime)
                    return cell.GetDateTime();

                if (cell.DataType == XLDataType.Number)
                {
                    try { return DateTime.FromOADate(cell.GetDouble()); }
                    catch { }
                }

                var s = cell.GetString().Trim();
                if (string.IsNullOrEmpty(s)) return null;

                DateTime dt;
                if (DateTime.TryParse(s, out dt)) return dt;

                string[] fmts =
                {
                    "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss",
                    "yyyy-MM-dd H:mm:ss",  "yyyy/MM/dd H:mm:ss",
                    "yyyy-MM-dd", "yyyy/MM/dd",
                    "M/d/yyyy H:mm:ss", "MM/dd/yyyy HH:mm:ss",
                    "M/d/yyyy", "MM/dd/yyyy"
                };

                if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    return dt;

                if (DateTime.TryParseExact(s, fmts, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                    return dt;
            }
            catch { }

            return null;
        }
    }
}