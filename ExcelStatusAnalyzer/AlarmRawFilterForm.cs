using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class AlarmRawFilterForm : Form
    {
        private Button btnLoad, btnCopy;
        private Label lblFile, lblHint;
        private TabControl tabSheets;
        private OpenFileDialog ofd;

        private HashSet<string> _apamaWhitelist;
        private HashSet<string> _apturaWhitelist;
        
        public AlarmRawFilterForm()
        {
            BuildUi();
            
            _apamaWhitelist = LoadWhitelistFromRoot(@"Lists\APAMA_ALID.txt");
            _apturaWhitelist = LoadWhitelistFromRoot(@"Lists\APTURA_ALID.txt");
        }

        private void BuildUi()
        {
            Text = "Alarm Raw Filter Form (화이트리스트 필터)";
            Width = 1300;
            Height = 820;
            
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
                Text = "현재 탭 복사",
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
                Width = 850,
                Text = "파일: (없음)"
            };
            
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1200,
                AutoSize = false,
                Text = "AlarmPivotForm4와 같은 양식의 엑셀을 불러와, F열 Alarm Name 기준으로 화이트리스트에 포함된 행만 시트별로 원본 그대로 추출합니다."
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
                Title = "대상 엑셀 파일 선택"
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
                        
                        var dt = BuildFilteredRawTable(ws, whitelist);
                        
                        var grid = CreateGrid();
                        grid.DataSource = dt;
                        grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                        
                        var tab = new TabPage(i + ". " + ws.Name) { Padding = new Padding(0) };
                        tab.Controls.Add(grid);
                        tabSheets.TabPages.Add(tab);
                    }
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
            CopyCurrentTab();
        }

        private DataTable BuildFilteredRawTable(IXLWorksheet ws, HashSet<string> whitelist)
        {
            var used = ws.RangeUsed();
            var dt = new DataTable();
            
            if (used == null)
                return dt;
            
            int firstRow = used.FirstRow().RowNumber();
            int lastRow = used.LastRow().RowNumber();
            int lastCol = used.LastColumn().ColumnNumber();
            
            // 1행 헤더 생성
            for (int c = 1; c <= lastCol; c++)
            {
                string header = ws.Cell(firstRow, c).GetString().Trim();
                if (string.IsNullOrWhiteSpace(header))
                    header = "Column" + c;
                
                // 중복 헤더 방지
                string finalHeader = header;
                int dup = 1;
                while (dt.Columns.Contains(finalHeader))
                {
                    finalHeader = header + "_" + dup;
                    dup++;
                }
                
                dt.Columns.Add(finalHeader);
            }

            // 2행부터 데이터
            for (int r = firstRow + 1; r <= lastRow; r++)
            {
                string alarm = GetCellString(ws.Cell(r, 6)); // F열 Alarm Name
                if (string.IsNullOrWhiteSpace(alarm)) continue;
                alarm = alarm.Trim();
                
                if (whitelist != null && whitelist.Count > 0 && !whitelist.Contains(alarm))
                    continue;
                
                var row = dt.NewRow();
                for (int c = 1; c <= lastCol; c++)
                {
                    row[c - 1] = GetCellString(ws.Cell(r, c));
                }
                dt.Rows.Add(row);
            }
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

        private void CopyCurrentTab()
        {
            var grid = GetCurrentGrid();
            if (grid == null) return;
            
            var dt = grid.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;
            
            var sb = new System.Text.StringBuilder();
            
            // 헤더 포함 복사
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                if (c > 0) sb.Append('\t');
                sb.Append(dt.Columns[c].ColumnName);
            }
            sb.Append('\n');
            
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    if (c > 0) sb.Append('\t');
                    sb.Append(Convert.ToString(dt.Rows[r][c]));
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
    }
}