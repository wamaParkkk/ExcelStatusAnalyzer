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
    public partial class AlarmPivotForm : Form
    {
        private Button btnLoad, btnCopy;
        private Label lblFile, lblHint;
        private TabControl tabSheets;
        private OpenFileDialog ofd;

        private const string TotalColName = "TOTAL";

        public AlarmPivotForm()
        {
            BuildUi();
        }
        
        private void BuildUi()
        {
            Text = "Alarm Name × 날짜별 발생횟수";
            Width = 1200;
            Height = 800;
            
            btnLoad = new Button { Text = "파일 불러오기 (.xlsx/.xls)", Left = 15, Top = 15, Width = 180, Height = 32 };
            btnLoad.Click += BtnLoad_Click;

            btnCopy = new Button { Text = "데이터 복사", Left = 205, Top = 15, Width = 150, Height = 32 };
            btnCopy.Click += BtnCopy_Click;

            lblFile = new Label { Left = 370, Top = 22, Width = 780, Text = "파일: (없음)" };
            
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1130,
                Text = "A열(A1=Alarm Name)과 C열(C1=Date) 기준. 시트 1~6 각각 피벗.",
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
                    int sheetCount = Math.Min(6, wb.Worksheets.Count);
                    if (sheetCount == 0)
                    {
                        MessageBox.Show("시트를 찾을 수 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    for (int i = 1; i <= sheetCount; i++)
                    {
                        var ws = wb.Worksheet(i);
                        var dt = BuildPivotForWorksheet(ws);
                        
                        var grid = CreateGrid();
                        grid.DataSource = dt;                        

                        // 총합 열 "보이기" + 헤더/정렬 지정
                        if (grid.Columns.Contains(TotalColName))
                        {
                            grid.Columns[TotalColName].HeaderText = "총합";
                            grid.Columns[TotalColName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            // 혹시 순서가 꼬이면 맨 뒤로
                            grid.Columns[TotalColName].DisplayIndex = grid.Columns.Count - 1;
                        }

                        // 숫자열 오른쪽 정렬 (0=Alarm Name, 나머지는 숫자)
                        for (int c = 1; c < grid.Columns.Count; c++)
                        {
                            grid.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        }
                        
                        // 화면에 보이는 만큼 컬럼 폭 조정
                        grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

                        var tab = new TabPage(i + ". " + ws.Name);
                        grid.Parent = tab;
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
            CopyCurrentAlarmDateCounts();
        }

        // === 시트 하나를 피벗 테이블로 변환 ===
        //  - A열: Alarm Name, C열: Date
        //  - 첫 행(A1/C1)은 헤더로 간주, 데이터는 2행부터
        //  - 날짜는 yyyy-MM-dd 로 열을 만들고, 알람명 × 날짜별 발생횟수 카운팅
        private DataTable BuildPivotForWorksheet(IXLWorksheet ws)
        {
            var used = ws.RangeUsed();
            
            var dtEmpty = new DataTable();
            dtEmpty.Columns.Add("Alarm Name");
            if (used == null) return dtEmpty;
            
            var counts = new Dictionary<string, Dictionary<DateTime, int>>(StringComparer.OrdinalIgnoreCase);
            var allDates = new HashSet<DateTime>();
            
            // 데이터는 2행부터 (A1=Alarm Name, C1=Date)
            foreach (var row in used.Rows().Skip(1))
            {
                string alarm = GetCellString(row.Cell(1));
                if (string.IsNullOrWhiteSpace(alarm)) continue;
                
                DateTime? d = TryReadDate(row.Cell(3));
                if (!d.HasValue) continue;
                
                var day = d.Value.Date;
                allDates.Add(day);
                
                Dictionary<DateTime, int> inner;
                if (!counts.TryGetValue(alarm, out inner))
                {
                    inner = new Dictionary<DateTime, int>();
                    counts[alarm] = inner;
                }

                int c;
                if (!inner.TryGetValue(day, out c)) c = 0;
                inner[day] = c + 1;
            }

            if (allDates.Count == 0) return dtEmpty;
            
            // 연속 날짜(최소~최대) 생성
            var minDate = allDates.Min();
            var maxDate = allDates.Max();
            var dateColumns = new List<DateTime>();
            for (var d = minDate; d <= maxDate; d = d.AddDays(1))
                dateColumns.Add(d);
            
            var dt = new DataTable();
            dt.Columns.Add("Alarm Name");
            for (int i = 0; i < dateColumns.Count; i++)
                dt.Columns.Add(dateColumns[i].ToString("yyyy-MM-dd"), typeof(int));
            
            // 합계 컬럼 추가 (정렬/복사용)
            dt.Columns.Add(TotalColName, typeof(int));
            
            foreach (var alarm in counts.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                var rowOut = dt.NewRow();
                rowOut[0] = alarm;
                
                var inner = counts[alarm];
                int total = 0;
                
                for (int i = 0; i < dateColumns.Count; i++)
                {
                    int v;
                    if (!inner.TryGetValue(dateColumns[i], out v)) v = 0;
                    rowOut[i + 1] = v;
                    total += v;
                }
                rowOut[TotalColName] = total;
                
                dt.Rows.Add(rowOut);
            }

            // 총합 내림차순 정렬 (같으면 알람명 오름차순)
            dt.DefaultView.Sort = "[" + TotalColName + "] DESC, [Alarm Name] ASC";
            dt = dt.DefaultView.ToTable();
            
            return dt;
        }

        private static string GetCellString(IXLCell cell)
        {
            if (cell == null) return string.Empty;
            
            switch (cell.DataType)
            {
                case XLDataType.Text:
                    return cell.GetString().Trim();
                
                case XLDataType.Number:
                    return cell.GetDouble().ToString(System.Globalization.CultureInfo.InvariantCulture).Trim();
                
                case XLDataType.Boolean:
                    return cell.GetBoolean() ? "TRUE" : "FALSE";
                
                case XLDataType.DateTime:
                    return cell.GetDateTime().ToString("yyyy-MM-dd HH:mm:ss");

                case XLDataType.Blank:
                    return string.Empty;

                default:                    
                    var s = cell.GetString();
                    if (!string.IsNullOrEmpty(s)) return s.Trim();
                    return cell.Value.ToString().Trim(); // XLCellValue -> string
            }
        }

        private static DateTime? TryReadDate(IXLCell cell)
        {
            if (cell == null) return null;

            if (cell.DataType == XLDataType.DateTime)
                return cell.GetDateTime();

            if (cell.DataType == XLDataType.Number)
            {
                // OA Date 가능성
                try { return DateTime.FromOADate(cell.GetDouble()); }
                catch { }
            }

            var s = cell.GetString().Trim();
            if (string.IsNullOrEmpty(s)) return null;

            DateTime dt;
            if (DateTime.TryParse(s, out dt))
                return dt;

            // yyyy-MM-dd 전용 재시도
            if (DateTime.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                                       DateTimeStyles.None, out dt))
                return dt;

            return null;
        }

        // Grid 생성
        private DataGridView CreateGrid()
        {
            var grid = new DataGridView();

            // 탭 페이지 전체를 채우기
            grid.Dock = DockStyle.Fill;
            grid.Margin = new Padding(0);
            grid.BorderStyle = BorderStyle.None;

            // 보기 전용 + 스크롤 + 기본 복사 금지
            grid.ReadOnly = true;
            grid.EditMode = DataGridViewEditMode.EditProgrammatically;
            grid.ScrollBars = ScrollBars.Both;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = false;
            grid.AllowUserToAddRows = false;
            grid.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) e.SuppressKeyPress = true;
                if ((e.Control && e.KeyCode == Keys.C) || (e.Control && e.KeyCode == Keys.Insert)) e.SuppressKeyPress = true;
            };

            return grid;
        }

        private void CopyCurrentAlarmDateCounts()
        {
            var grid = GetCurrentGrid();
            if (grid == null) return;
            
            var dt = grid.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;
            
            // 열 인덱스: 0=Alarm Name, 1~(N-1)=날짜 열..., 마지막에 숨김 합계열(TOTAL)이 있을 수 있음
            // 복사: Alarm Name + 날짜열 값들만 (헤더 X, 합계열 X)
            var sb = new System.Text.StringBuilder();
            
            // 날짜 열 인덱스 수집 (총합 열 제외)
            var dateColIndexes = new List<int>();
            for (int c = 1; c < dt.Columns.Count; c++)
            {
                var colName = dt.Columns[c].ColumnName;
                if (string.Equals(colName, TotalColName, StringComparison.Ordinal))
                    continue; // 합계 열 제외
                dateColIndexes.Add(c);
            }
            
            // 각 행: [Alarm Name] \t [날짜1값] \t [날짜2값] ...
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];
                
                // 첫 칸: Alarm Name
                sb.Append(Convert.ToString(row["Alarm Name"]));
                
                // 날짜 값들: 숫자만
                for (int i = 0; i < dateColIndexes.Count; i++)
                {
                    var idx = dateColIndexes[i];
                    var cell = row[idx];
                    int num;
                    if (!int.TryParse(Convert.ToString(cell), out num))
                        num = 0;
                    sb.Append('\t').Append(num.ToString()); // 헤더 미포함, 숫자만
                }

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
    }
}