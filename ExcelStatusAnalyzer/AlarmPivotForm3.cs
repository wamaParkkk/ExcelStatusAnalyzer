using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
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
            Text = "Alarm Pivot Form 3 (모든 시트 B열 Alarm 통합 합계) (KTCB-모델별)";
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

                var dt = BuildMergedAlarmCountTable(path);

                dgv.DataSource = null;
                dgv.Columns.Clear();
                dgv.AutoGenerateColumns = true;
                dgv.DataSource = dt;
                
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

                MessageBox.Show("Rows = " + ((dt == null) ? 0 : dt.Rows.Count));
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyAlarmAndTotalOnly();
        }
        private DataTable BuildMergedAlarmCountTable(string excelPath)
        {
            var dict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            
            using (var wb = new XLWorkbook(excelPath))
            {
                foreach (var ws in wb.Worksheets)
                {
                    // RangeUsed 말고, LastRowUsed 기준으로 B열 스캔
                    var lastRow = ws.LastRowUsed();
                    if (lastRow == null) continue;
                    
                    int last = lastRow.RowNumber();
                    if (last <= 0) continue;
                    
                    var alarmsInSheet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    
                    for (int r = 1; r <= last; r++)
                    {
                        // B열(2열)
                        var cell = ws.Cell(r, 2);
                        var alarm = cell.GetString();   // ✅ 텍스트 셀엔 이게 가장 안전
                        if (string.IsNullOrWhiteSpace(alarm)) continue;
                        
                        alarmsInSheet.Add(alarm.Trim());
                    }

                    foreach (var a in alarmsInSheet)
                    {
                        int cur;
                        if (!dict.TryGetValue(a, out cur)) cur = 0;
                        dict[a] = cur + 1;
                    }
                }
            }

            var dt = new DataTable();
            dt.Columns.Add("Alarm");
            dt.Columns.Add("합계", typeof(int));
            
            foreach (var kv in dict.OrderByDescending(x => x.Value)
                                   .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                dt.Rows.Add(kv.Key, kv.Value);
            }

            return dt;
        }

        private static string GetCellString(IXLCell cell)
        {
            if (cell == null) return string.Empty;

            // ClosedXML 셀 타입별 안전 처리 (C#7.3)
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
                    return cell.Value.ToString().Trim();
            }
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

        private void CopyAlarmAndTotalOnly()
        {
            var dt = dgv.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;

            // 헤더 없이: Alarm \t 합계
            var sb = new System.Text.StringBuilder();
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                var row = dt.Rows[r];
                sb.Append(Convert.ToString(row["Alarm"]));
                sb.Append('\t');
                sb.Append(Convert.ToString(row["합계"]));
                sb.Append('\n');
            }

            Clipboard.Clear();
            Clipboard.SetText(sb.ToString());
        }
    }
}