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
    public partial class DailyTrackerPivotForm : Form
    {
        private Button btnLoad, btnCopy;
        private Label lblFile, lblHint;
        private TabControl tabModels;
        private OpenFileDialog ofd;

        private const string TabApama = "APAMA";
        private const string TabAptura = "APTURA";

        private const string ColDesc = "Description";
        private const string ColFreq = "Total Frequency";
        private const string ColTimeMin = "Total Time (min)";

        public DailyTrackerPivotForm()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "Daily Tracker Pivot (Description별 Frequency/Time 합산)";
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
                Text = "A=Description, B=Frequency, C=Time. 시트1~4=APAMA, 시트5~6=APTURA (각각 Description별 합산)",
                AutoSize = false
            };

            tabModels = new TabControl
            {
                Left = 15,
                Top = 80,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 95,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            tabModels.TabPages.Clear();
            tabModels.TabPages.Add(new TabPage(TabApama) { Padding = new Padding(0) });
            tabModels.TabPages.Add(new TabPage(TabAptura) { Padding = new Padding(0) });

            ofd = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx;*.xls",
                Title = "Daily Tracker 엑셀 파일 선택"
            };

            Controls.Add(btnLoad);
            Controls.Add(btnCopy);
            Controls.Add(lblFile);
            Controls.Add(lblHint);
            Controls.Add(tabModels);
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var path = ofd.FileName;
                lblFile.Text = "파일: " + Path.GetFileName(path);

                using (var wb = new XLWorkbook(path))
                {
                    // 시트1~4 => APAMA, 시트5~6 => APTURA
                    var dtApama = BuildModelPivot(wb, 1, 4);
                    var dtAptura = BuildModelPivot(wb, 5, 6);

                    BindToTab(TabApama, dtApama);
                    BindToTab(TabAptura, dtAptura);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyCurrentModelTable();
        }

        // =========================
        // 모델별(여러 시트) 합산
        // =========================
        private sealed class Agg
        {
            public long TotalFreq;
            public double TotalMinutes;
        }
        private DataTable BuildModelPivot(XLWorkbook wb, int fromSheetIndex, int toSheetIndex)
        {
            var sums = new Dictionary<string, Agg>(StringComparer.OrdinalIgnoreCase);

            int maxSheet = wb.Worksheets.Count;
            int start = Math.Max(1, fromSheetIndex);
            int end = Math.Min(maxSheet, toSheetIndex);

            for (int si = start; si <= end; si++)
            {
                var ws = wb.Worksheet(si);
                if (ws == null) continue;

                var used = ws.RangeUsed();
                if (used == null) continue;

                // 1행은 헤더, 데이터는 2행부터
                foreach (var row in used.Rows().Skip(1))
                {
                    string desc = GetCellString(row.Cell(1));       // A
                    if (string.IsNullOrWhiteSpace(desc)) continue;
                    desc = desc.Trim();

                    long freq = TryReadLong(row.Cell(2));           // B
                    double minutes = TryReadMinutes(row.Cell(3));   // C

                    // 둘 다 0이면 의미 없는 행이라 스킵(원하면 freq만 있어도 포함 가능)
                    if (freq == 0 && minutes <= 0) continue;

                    Agg agg;
                    if (!sums.TryGetValue(desc, out agg))
                    {
                        agg = new Agg();
                        sums[desc] = agg;
                    }

                    agg.TotalFreq += freq;
                    agg.TotalMinutes += minutes;
                }
            }

            var dt = new DataTable();
            dt.Columns.Add(ColDesc);
            dt.Columns.Add(ColFreq, typeof(long));
            dt.Columns.Add(ColTimeMin, typeof(double));

            foreach (var kv in sums
                .OrderByDescending(x => x.Value.TotalMinutes)
                .ThenByDescending(x => x.Value.TotalFreq)
                .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                var row = dt.NewRow();
                row[ColDesc] = kv.Key;
                row[ColFreq] = kv.Value.TotalFreq;
                row[ColTimeMin] = Math.Round(kv.Value.TotalMinutes, 1);
                dt.Rows.Add(row);
            }

            return dt;
        }

        // =========================
        // 탭 바인딩
        // =========================
        private void BindToTab(string tabName, DataTable dt)
        {
            TabPage tab = null;
            foreach (TabPage t in tabModels.TabPages)
            {
                if (string.Equals(t.Text, tabName, StringComparison.OrdinalIgnoreCase))
                {
                    tab = t;
                    break;
                }
            }
            if (tab == null) return;

            tab.Controls.Clear();

            var grid = CreateGrid();
            grid.DataSource = dt;

            if (grid.Columns.Contains(ColFreq))
                grid.Columns[ColFreq].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            if (grid.Columns.Contains(ColTimeMin))
                grid.Columns[ColTimeMin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            tab.Controls.Add(grid);
        }

        // =========================
        // Copy (헤더 없이, 버튼으로만)
        // =========================
        private void CopyCurrentModelTable()
        {
            var grid = GetCurrentGrid();
            if (grid == null) return;

            var dt = grid.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;

            // Description + Total Frequency + Total Time(min) (헤더 제외)
            var sb = new StringBuilder();
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];
                sb.Append(Convert.ToString(row[ColDesc]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColFreq]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColTimeMin]));
                sb.Append('\n');
            }

            Clipboard.Clear();
            Clipboard.SetText(sb.ToString());
        }

        private DataGridView GetCurrentGrid()
        {
            if (tabModels.TabPages.Count == 0) return null;
            var tab = tabModels.SelectedTab;
            if (tab == null || tab.Controls.Count == 0) return null;
            return tab.Controls[0] as DataGridView;
        }

        // =========================
        // Grid (읽기전용/CTRL+C 차단)
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

        private static long TryReadLong(IXLCell cell)
        {
            if (cell == null) return 0;

            try
            {
                if (cell.DataType == XLDataType.Number)
                    return (long)Math.Round(cell.GetDouble(), 0);

                var s = cell.GetString().Trim();
                if (string.IsNullOrEmpty(s)) return 0;

                long v;
                if (long.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return v;
                if (long.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.CurrentCulture, out v)) return v;

                // 혹시 소수 포함이면 double로 읽어서 반올림
                double d;
                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return (long)Math.Round(d, 0);
                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return (long)Math.Round(d, 0);
            }
            catch { }

            return 0;
        }

        // Time 컬럼(C)을 "분"으로 환산해서 합산
        // - 숫자면:  (1) 이미 '분'일 수도 있고 (2) 엑셀 time fraction(하루=1)일 수도 있음
        // - DateTime면: TimeOfDay를 분으로
        // - 문자열이면: TimeSpan/DateTime/숫자 순서로 시도
        private static double TryReadMinutes(IXLCell cell)
        {
            if (cell == null) return 0;

            try
            {
                if (cell.DataType == XLDataType.DateTime)
                {
                    return cell.GetDateTime().TimeOfDay.TotalMinutes;
                }

                if (cell.DataType == XLDataType.Number)
                {
                    double v = cell.GetDouble();

                    // 엑셀 time fraction 가능성(0~1 사이면 하루 비율)
                    // 0.5 = 12시간 = 720분
                    if (v > 0 && v < 1.0)
                        return v * 24.0 * 60.0;

                    // 그 외는 "분"으로 간주(일반적으로 분 단위 입력일 확률이 높음)
                    return v;
                }

                var s = cell.GetString().Trim();
                if (string.IsNullOrEmpty(s)) return 0;

                // TimeSpan 시도 (hh:mm:ss / hh:mm 등)
                TimeSpan ts;
                if (TimeSpan.TryParse(s, out ts))
                    return ts.TotalMinutes;

                // DateTime 시도
                DateTime dt;
                if (DateTime.TryParse(s, out dt))
                    return dt.TimeOfDay.TotalMinutes;

                // 숫자 시도
                double d;
                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                {
                    if (d > 0 && d < 1.0) return d * 24.0 * 60.0;
                    return d;
                }

                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                {
                    if (d > 0 && d < 1.0) return d * 24.0 * 60.0;
                    return d;
                }
            }
            catch { }

            return 0;
        }
    }
}