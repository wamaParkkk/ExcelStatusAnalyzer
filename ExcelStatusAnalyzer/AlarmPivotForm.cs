﻿using ClosedXML.Excel;
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
        private CheckBox chkDay, chkSwing, chkNight;
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

            chkDay = new CheckBox { Left = 15, Top = 52, Width = 60, Text = "Day", Checked = true };
            chkSwing = new CheckBox { Left = 85, Top = 52, Width = 70, Text = "Swing", Checked = true };
            chkNight = new CheckBox { Left = 165, Top = 52, Width = 70, Text = "Night", Checked = true };

            lblFile = new Label { Left = 370, Top = 22, Width = 780, Text = "파일: (없음)" };
            
            lblHint = new Label
            {
                Left = 245,
                Top = 57,
                Width = 900,
                Text = "Day 06:00–13:59, Swing 14:00–21:59, Night 22:00–05:59",
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
            Controls.Add(chkDay);
            Controls.Add(chkSwing);
            Controls.Add(chkNight);
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

                var incDay = chkDay.Checked;
                var incSwing = chkSwing.Checked;
                var incNight = chkNight.Checked;

                // 세 개 모두 미체크면 → 전체 포함(안내 없이 조용히 전체 선택과 동일 처리)
                if (!incDay && !incSwing && !incNight)
                {
                    incDay = incSwing = incNight = true;
                }

                using (var wb = new XLWorkbook(path))
                {
                    int sheetCount = Math.Min(6, wb.Worksheets.Count);
                    for (int i = 1; i <= sheetCount; i++)
                    {
                        var ws = wb.Worksheet(i);
                        
                        // 교대 선택을 전달
                        var dt = BuildPivotForWorksheet(ws, incDay, incSwing, incNight);
                        
                        var grid = CreateGrid();
                        var tab = new TabPage(i + ". " + ws.Name) { Padding = new Padding(0) };
                        tab.Controls.Add(grid);
                        grid.DataSource = dt;
                        
                        // 총합 열 보이기 + 정렬/정렬 방향 유지
                        if (grid.Columns.Contains(TotalColName))
                        {
                            grid.Columns[TotalColName].HeaderText = "총합";
                            grid.Columns[TotalColName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            grid.Columns[TotalColName].DisplayIndex = grid.Columns.Count - 1;
                        }

                        for (int c = 1; c < grid.Columns.Count; c++)
                            grid.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        
                        grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
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
        private DataTable BuildPivotForWorksheet(IXLWorksheet ws, bool incDay, bool incSwing, bool incNight)
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
                
                // 전체 DateTime 필요 (시간대 필터 때문에)
                DateTime? stamp = TryReadDate(row.Cell(3));
                if (!stamp.HasValue) continue;
                
                // 교대(시간대) 필터 적용
                if (!ShouldIncludeByShift(stamp.Value.TimeOfDay, incDay, incSwing, incNight))
                    continue;
                
                var day = stamp.Value.Date; // 집계는 '날짜' 단위 (열은 yyyy-MM-dd)
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
            
            // 날짜를 연속 범위(min~max)로 확장
            var minDate = allDates.Min();
            var maxDate = allDates.Max();
            var dateColumns = new List<DateTime>();
            for (var d = minDate; d <= maxDate; d = d.AddDays(1))
                dateColumns.Add(d);
            
            // 결과 테이블
            var dt = new DataTable();
            dt.Columns.Add("Alarm Name");
            for (int i = 0; i < dateColumns.Count; i++)
                dt.Columns.Add(dateColumns[i].ToString("yyyy-MM-dd"), typeof(int));
            dt.Columns.Add(TotalColName, typeof(int)); // 총합(표시/정렬용)
            
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
            
            // 총합 내림차순, 동률이면 알람명 오름차순
            dt.DefaultView.Sort = "[" + TotalColName + "] DESC, [Alarm Name] ASC";
            dt = dt.DefaultView.ToTable();
            
            return dt;
        }

        // 교대 포함 여부 판단 (Day: 06:00–13:59:59, Swing: 14:00–21:59:59, Night: 22:00–05:59:59)
        private bool ShouldIncludeByShift(TimeSpan t, bool incDay, bool incSwing, bool incNight)
        {
            // 세 개 모두 체크면 → 전체 포함(기존과 동일)
            if (incDay && incSwing && incNight) return true;
            
            bool inDay = (t >= new TimeSpan(6, 0, 0) && t <= new TimeSpan(13, 59, 59));
            bool inSwing = (t >= new TimeSpan(14, 0, 0) && t <= new TimeSpan(21, 59, 59));
            bool inNight = (t >= new TimeSpan(22, 0, 0) || t <= new TimeSpan(5, 59, 59)); // 자정 넘어감
            
            bool ok = (incDay && inDay) || (incSwing && inSwing) || (incNight && inNight);
            
            // 세 개 모두 미체크면 → 전체 포함(안전장치)
            if (!incDay && !incSwing && !incNight) return true;
            
            return ok;
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