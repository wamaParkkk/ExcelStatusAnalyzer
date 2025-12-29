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
    public partial class AlarmPivotForm3 : Form
    {
        private Button btnLoad, btnCopy;
        private Label lblFile, lblHint;
        private DataGridView dgv;
        private OpenFileDialog ofd;

        public AlarmPivotForm3()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "Alarm 통합 합계 (KTCB-모델별)";
            Width = 900;
            Height = 760;

            btnLoad = new Button { Text = "파일 불러오기 (.xlsx/.xls)", Left = 15, Top = 15, Width = 180, Height = 32 };
            btnLoad.Click += BtnLoad_Click;

            btnCopy = new Button { Text = "복사 (Alarm + 합계)", Left = 205, Top = 15, Width = 170, Height = 32 };
            btnCopy.Click += BtnCopy_Click;

            lblFile = new Label { Left = 390, Top = 22, Width = 480, Text = "파일: (없음)" };
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 850,
                Text = "모든 시트의 B열(2열) Alarm을 합쳐 Alarm별 총 발생 횟수를 집계합니다. (헤더 유무 상관없음)"
            };

            dgv = CreateGrid();
            dgv.Left = 15;
            dgv.Top = 80;
            dgv.Width = ClientSize.Width - 30;
            dgv.Height = ClientSize.Height - 95;
            dgv.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

            ofd = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx;*.xls",
                Title = "집계 대상 엑셀 파일 선택"
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
                lblFile.Text = "파일: " + Path.GetFileName(path);
                
                // 1) 집계
                var dt = BuildMergedAlarmDatePivotTable(path);
                
                // 2) 바인딩 (표시 문제 방지)
                dgv.DataSource = null;
                dgv.Columns.Clear();
                dgv.AutoGenerateColumns = true;
                dgv.DataSource = dt;
                
                // 3) 그리드 표시 옵션 (날짜 컬럼이 많으니 스크롤/폭 고정이 유리)
                dgv.ScrollBars = ScrollBars.Both;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
                
                // 4) TOTAL(총합) 열 표시/정렬/우측 고정 느낌(맨 뒤로)
                if (dgv.Columns.Contains(TotalColName))
                {
                    dgv.Columns[TotalColName].HeaderText = "총합";
                    dgv.Columns[TotalColName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dgv.Columns[TotalColName].DisplayIndex = dgv.Columns.Count - 1;
                }
                
                // 5) 숫자 컬럼 우측 정렬
                for (int c = 1; c < dgv.Columns.Count; c++)
                    dgv.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                
                // 6) 컬럼 폭 자동 조정(너무 많으면 느릴 수 있어 DisplayedCells 추천)
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                
                // 결과가 없을 때만 안내
                if (dt == null || dt.Rows.Count == 0)
                {
                    MessageBox.Show("집계된 데이터가 없습니다.\n(A1=No, B1=Alarm Name, C1~날짜 / 2행부터 Count 구조인지 확인)",
                        "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyCurrentAlarmDateCounts();
        }

        private DataGridView CreateGrid()
        {
            var grid = new DataGridView();

            grid.ReadOnly = true;
            grid.EditMode = DataGridViewEditMode.EditProgrammatically;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;

            grid.ScrollBars = ScrollBars.Both;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            // Ctrl+C 금지 + 개별 복사 금지
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = false;

            grid.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) e.SuppressKeyPress = true;
                if ((e.Control && e.KeyCode == Keys.C) || (e.Control && e.KeyCode == Keys.Insert)) e.SuppressKeyPress = true;
            };

            return grid;
        }

        private const string TotalColName = "TOTAL";
        private DataTable BuildMergedAlarmDatePivotTable(string excelPath)
        {
            // alarm -> (date -> count)
            var merged = new Dictionary<string, Dictionary<DateTime, long>>(StringComparer.OrdinalIgnoreCase);
            var allDates = new HashSet<DateTime>();
            
            using (var wb = new XLWorkbook(excelPath))
            {
                foreach (var ws in wb.Worksheets)
                {
                    var lastRow = ws.LastRowUsed();
                    var lastCell = ws.LastCellUsed();
                    if (lastRow == null || lastCell == null) continue;
                    
                    int lastRowNo = lastRow.RowNumber();
                    int lastColNo = lastCell.Address.ColumnNumber;
                    
                    // C1부터 날짜 헤더 읽기
                    var dateCols = new List<Tuple<int, DateTime>>(); // (col, date)
                    bool started = false;
                    
                    for (int col = 3; col <= lastColNo; col++)
                    {
                        var d = TryReadDate(ws.Cell(1, col));
                        if (!d.HasValue)
                        {
                            if (started) break; // 날짜 시작 이후 빈칸이면 종료로 간주
                            continue;
                        }
                        
                        started = true;
                        var day = d.Value.Date;
                        dateCols.Add(Tuple.Create(col, day));
                        allDates.Add(day);
                    }

                    if (dateCols.Count == 0) continue;
                    
                    // 2행부터 Alarm별 날짜 카운트 누적
                    for (int r = 2; r <= lastRowNo; r++)
                    {
                        string alarm = ws.Cell(r, 2).GetString().Trim(); // B열
                        if (string.IsNullOrWhiteSpace(alarm)) continue;
                        
                        // 혹시 합계행이 있으면 제외(안전)
                        if (string.Equals(alarm, "TOTAL", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(alarm, "합계", StringComparison.OrdinalIgnoreCase))
                            continue;
                        
                        Dictionary<DateTime, long> inner;
                        if (!merged.TryGetValue(alarm, out inner))
                        {
                            inner = new Dictionary<DateTime, long>();
                            merged[alarm] = inner;
                        }
                        
                        for (int i = 0; i < dateCols.Count; i++)
                        {
                            int col = dateCols[i].Item1;
                            DateTime day = dateCols[i].Item2;
                            
                            long v = TryReadLong(ws.Cell(r, col));
                            if (v == 0) continue;
                            
                            long cur;
                            if (!inner.TryGetValue(day, out cur)) cur = 0;
                            inner[day] = cur + v;
                        }
                    }
                }
            }

            // 날짜 연속 채우기(min~max) + 컬럼 순서 확정
            var dateList = new List<DateTime>();
            if (allDates.Count > 0)
            {
                var min = allDates.Min();
                var max = allDates.Max();
                for (var d = min; d <= max; d = d.AddDays(1))
                    dateList.Add(d);
            }

            // 결과 테이블: Alarm Name + 날짜 컬럼들 + TOTAL
            var dt = new DataTable();
            dt.Columns.Add("Alarm Name");
            for (int i = 0; i < dateList.Count; i++)
                dt.Columns.Add(dateList[i].ToString("yyyy-MM-dd"), typeof(long));
            dt.Columns.Add(TotalColName, typeof(long));
            
            foreach (var alarm in merged.Keys)
            {
                var row = dt.NewRow();
                row[0] = alarm;
                
                long total = 0;
                var inner = merged[alarm];
                
                for (int i = 0; i < dateList.Count; i++)
                {
                    long v;
                    if (!inner.TryGetValue(dateList[i], out v)) v = 0;
                    row[i + 1] = v;
                    total += v;
                }
                
                row[TotalColName] = total;
                
                // 총합 0이면 제외
                if (total == 0) continue;
                
                dt.Rows.Add(row);
            }

            // TOTAL 내림차순 정렬 후 DataTable로 확정
            if (dt.Rows.Count > 0)
            {
                dt.DefaultView.Sort = "[" + TotalColName + "] DESC, [Alarm Name] ASC";
                dt = dt.DefaultView.ToTable();
            }

            return dt;
        }

        private static DateTime? TryReadDate(IXLCell cell)
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
                    "M/d", "MM/dd", "M/d/yyyy", "MM/dd/yyyy",
                    "yyyy-MM-dd", "yyyy/MM/dd",
                    "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss"
                };

                if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    return dt;
                
                if (DateTime.TryParseExact(s, fmts, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                    return dt;
            }
            catch { }

            return null;
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
                
                long n;
                if (long.TryParse(s.Replace(",", ""), out n)) return n;
                
                double d;
                if (double.TryParse(s.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return (long)Math.Round(d, 0);
            }
            catch { }

            return 0;
        }

        private void CopyCurrentAlarmDateCounts()
        {
            var dt = dgv.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;
            
            const string alarmCol = "Alarm Name";
            
            if (!dt.Columns.Contains(alarmCol))
            {
                MessageBox.Show("복사 실패: 'Alarm Name' 컬럼을 찾을 수 없습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // 복사 대상: Alarm Name + 날짜 컬럼들 (TOTAL 제외)
            var valueColIndexes = new List<int>();
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var name = dt.Columns[c].ColumnName;
                if (string.Equals(name, TotalColName, StringComparison.OrdinalIgnoreCase))
                    continue; // TOTAL 제외
                valueColIndexes.Add(c);
            }
            
            var sb = new System.Text.StringBuilder();
            
            // 헤더 복사 X (요구사항)
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];
                
                // Alarm Name
                sb.Append(Convert.ToString(row[alarmCol]));
                
                // 날짜 값들(숫자)
                for (int i = 1; i < valueColIndexes.Count; i++)
                {
                    int idx = valueColIndexes[i];
                    long num;
                    if (!long.TryParse(Convert.ToString(row[idx]), out num)) num = 0;
                    sb.Append('\t').Append(num.ToString());
                }
                
                sb.Append('\n');
            }
            
            Clipboard.Clear();
            Clipboard.SetText(sb.ToString());
        }
    }
}