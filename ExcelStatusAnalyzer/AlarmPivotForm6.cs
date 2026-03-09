using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class AlarmPivotForm6 : Form
    {
        private Button btnLoad, btnCopy;
        private Label lblFile, lblHint;
        private TabControl tabSheets;
        private OpenFileDialog ofd;

        private HashSet<string> _apamaWhitelist;
        private HashSet<string> _apturaWhitelist;

        private const string ColAlarm = "Alarm Name";
        private const string ColCount = "Count";
        private const string ColAvgHour = "Avg Hours";
        private const string ColMaxHour = "Max Hours";

        public AlarmPivotForm6()
        {
            BuildUi();

            _apamaWhitelist = LoadWhitelistFromRoot(@"Lists\APAMA_ALID.txt");
            _apturaWhitelist = LoadWhitelistFromRoot(@"Lists\APTURA_ALID.txt");
        }

        private void BuildUi()
        {
            Text = "3회 이상 & 3시간 이상 Alarm(KTCB-장비별 MTBA))";
            Width = 1200;
            Height = 800;

            btnLoad = new Button
            {
                Text = "파일 불러오기 (.xlsx/.xls)",
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
                Width = 150,
                Height = 32
            };
            btnCopy.Click += BtnCopy_Click;
            
            lblFile = new Label
            {
                Left = 370,
                Top = 22,
                Width = 800,
                Text = "파일: (없음)"
            };
            
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1120,
                Text = "F=Alarm Name, G=Start, H=End. WhiteList 적용 후, 같은 Alarm Name이 3회 이상이고 각 건의 지속시간이 3시간 이상인 Alarm만 추출. 종료시간이 9999-99-99 99:99:99 같은 비정상 값은 제외.",
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
                
                using (var wb = new XLWorkbook(path))
                {
                    int sheetCount = wb.Worksheets.Count;
                    
                    for (int i = 1; i <= sheetCount; i++)
                    {
                        var ws = wb.Worksheet(i);
                        
                        HashSet<string> whitelist = null;
                        if (i >= 1 && i <= 4) whitelist = _apamaWhitelist;
                        else if (i >= 5 && i <= 6) whitelist = _apturaWhitelist;
                        
                        var dt = BuildSheetTable(ws, whitelist);
                        
                        var grid = CreateGrid();
                        grid.DataSource = dt;
                        
                        if (grid.Columns.Contains(ColCount))
                            grid.Columns[ColCount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        
                        if (grid.Columns.Contains(ColAvgHour))
                            grid.Columns[ColAvgHour].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        
                        if (grid.Columns.Contains(ColMaxHour))
                            grid.Columns[ColMaxHour].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        
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

        private sealed class AlarmAgg
        {
            public int Count;
            public List<double> Hours = new List<double>();
        }

        private DataTable BuildSheetTable(IXLWorksheet ws, HashSet<string> whitelist)
        {
            var sums = new Dictionary<string, AlarmAgg>(StringComparer.OrdinalIgnoreCase);
            
            var lastRow = ws.LastRowUsed();
            var lastCell = ws.LastCellUsed();
            if (lastRow == null || lastCell == null)
                return MakeEmptyTable();
            
            int lastRowNo = lastRow.RowNumber();
            
            // 1행은 헤더, 2행부터 데이터
            for (int r = 2; r <= lastRowNo; r++)
            {
                string alarm = GetCellString(ws.Cell(r, 6)); // F
                if (string.IsNullOrWhiteSpace(alarm)) continue;
                alarm = alarm.Trim();
                
                if (whitelist != null && whitelist.Count > 0 && !whitelist.Contains(alarm))
                    continue;
                
                // 종료시간 비정상값 제외
                if (IsInvalidEndTime(ws.Cell(r, 8))) continue;
                
                DateTime? startDt = TryReadDateTime(ws.Cell(r, 7)); // G
                DateTime? endDt = TryReadDateTime(ws.Cell(r, 8));   // H
                
                if (!startDt.HasValue || !endDt.HasValue) continue;
                
                var diff = endDt.Value - startDt.Value;
                if (diff.TotalSeconds < 0) continue;
                
                double hours = diff.TotalHours;
                
                // 각 건이 3시간 이상이어야 함
                if (hours < 3.0) continue;
                
                AlarmAgg agg;
                if (!sums.TryGetValue(alarm, out agg))
                {
                    agg = new AlarmAgg();
                    sums[alarm] = agg;
                }
                
                agg.Count += 1;
                agg.Hours.Add(hours);
            }

            var dt = MakeEmptyTable();
            
            // 3회 이상 AND 3시간 이상(위에서 이미 3시간 이상만 카운트함)
            foreach (var kv in sums
                .Where(x => x.Value.Count >= 3)
                .OrderByDescending(x => x.Value.Count)
                .ThenByDescending(x => x.Value.Hours.Count > 0 ? x.Value.Hours.Max() : 0)
                .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                var alarm = kv.Key;
                var agg = kv.Value;
                
                double avgHour = agg.Hours.Count > 0 ? agg.Hours.Average() : 0;
                double maxHour = agg.Hours.Count > 0 ? agg.Hours.Max() : 0;
                
                var row = dt.NewRow();
                row[ColAlarm] = alarm;
                row[ColCount] = agg.Count;
                row[ColAvgHour] = Math.Round(avgHour, 2);
                row[ColMaxHour] = Math.Round(maxHour, 2);
                dt.Rows.Add(row);
            }

            return dt;
        }

        private DataTable MakeEmptyTable()
        {
            var dt = new DataTable();
            dt.Columns.Add(ColAlarm);
            dt.Columns.Add(ColCount, typeof(int));
            dt.Columns.Add(ColAvgHour, typeof(double));
            dt.Columns.Add(ColMaxHour, typeof(double));
            return dt;
        }

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

        private bool IsInvalidEndTime(IXLCell cell)
        {
            if (cell == null) return true;
            
            var raw = GetCellString(cell);
            if (string.IsNullOrWhiteSpace(raw)) return true;
            
            if (raw.Contains("9999-99-99") || raw.Contains("99:99:99"))
                return true;
            
            return false;
        }

        private void CopyCurrentSheetSummary()
        {
            var grid = GetCurrentGrid();
            if (grid == null) return;
            
            var dt = grid.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;
            
            var sb = new System.Text.StringBuilder();
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];
                sb.Append(Convert.ToString(row[ColAlarm]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColCount]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColAvgHour]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row[ColMaxHour]));
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