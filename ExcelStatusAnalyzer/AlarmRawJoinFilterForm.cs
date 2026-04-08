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
    public partial class AlarmRawJoinFilterForm : Form
    {
        private Button btnLoadRaw;
        private Button btnLoadMap;
        private Button btnCopy;
        private Button btnC1;
        private Button btnC2;
        private Button btnLeft;
        private Button btnRight;
        
        private Label lblRawFile;
        private Label lblMapFile;
        private Label lblHint;
        
        private DataGridView dgv;
        
        private OpenFileDialog ofdRaw;
        private OpenFileDialog ofdMap;
        
        private string _rawPath;
        private string _mapPath;
        
        private DataTable _allData;
        
        // true면 "포함된 항목 제외"
        private bool _filterC1;
        private bool _filterC2;
        private bool _filterLeft;
        private bool _filterRight;

        private const int DuplicateWithinSeconds = 1; // 필요시 1, 2, 5 등으로 변경

        public AlarmRawJoinFilterForm()
        {
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "Alarm Raw Join Filter Form";
            Width = 1450;
            Height = 850;
            
            btnLoadRaw = new Button
            {
                Text = "1번 파일 불러오기",
                Left = 15,
                Top = 15,
                Width = 150,
                Height = 32
            };
            btnLoadRaw.Click += BtnLoadRaw_Click;
            
            btnLoadMap = new Button
            {
                Text = "2번 파일 불러오기",
                Left = 175,
                Top = 15,
                Width = 150,
                Height = 32
            };
            btnLoadMap.Click += BtnLoadMap_Click;
            
            btnCopy = new Button
            {
                Text = "데이터 복사",
                Left = 335,
                Top = 15,
                Width = 120,
                Height = 32
            };
            btnCopy.Click += BtnCopy_Click;
            
            btnC1 = new Button
            {
                Text = "C1 제외",
                Left = 475,
                Top = 15,
                Width = 80,
                Height = 32
            };
            btnC1.Click += (s, e) =>
            {
                _filterC1 = !_filterC1;
                UpdateFilterButtonStyle();
                ApplyFilters();
            };

            btnC2 = new Button
            {
                Text = "C2 제외",
                Left = 565,
                Top = 15,
                Width = 80,
                Height = 32
            };
            btnC2.Click += (s, e) =>
            {
                _filterC2 = !_filterC2;
                UpdateFilterButtonStyle();
                ApplyFilters();
            };

            btnLeft = new Button
            {
                Text = "Left 제외",
                Left = 655,
                Top = 15,
                Width = 90,
                Height = 32
            };
            btnLeft.Click += (s, e) =>
            {
                _filterLeft = !_filterLeft;
                UpdateFilterButtonStyle();
                ApplyFilters();
            };

            btnRight = new Button
            {
                Text = "Right 제외",
                Left = 755,
                Top = 15,
                Width = 95,
                Height = 32
            };
            btnRight.Click += (s, e) =>
            {
                _filterRight = !_filterRight;
                UpdateFilterButtonStyle();
                ApplyFilters();
            };

            lblRawFile = new Label
            {
                Left = 15,
                Top = 55,
                Width = 1200,
                Text = "1번 파일: (없음)"
            };

            lblMapFile = new Label
            {
                Left = 15,
                Top = 78,
                Width = 1200,
                Text = "2번 파일: (없음)"
            };

            lblHint = new Label
            {
                Left = 15,
                Top = 102,
                Width = 1350,
                Height = 22,
                Text = "1번 파일 F열 알람 내용 = 2번 파일 B열 ID 를 매칭하여, 2번 파일의 J열 Text / K열 Root Cause 를 붙입니다. 버튼은 중복 선택 가능하며, 선택된 키워드가 포함된 알람은 제외됩니다."
            };

            dgv = new DataGridView
            {
                Left = 15,
                Top = 130,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 145,
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

            ofdRaw = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx;*.xls",
                Title = "1번 파일 선택"
            };

            ofdMap = new OpenFileDialog
            {
                Filter = "CSV|*.csv;*.CSV|Excel|*.xlsx;*.xls|All Files|*.*",
                Title = "2번 파일 선택"
            };

            Controls.Add(btnLoadRaw);
            Controls.Add(btnLoadMap);
            Controls.Add(btnCopy);
            Controls.Add(btnC1);
            Controls.Add(btnC2);
            Controls.Add(btnLeft);
            Controls.Add(btnRight);
            Controls.Add(lblRawFile);
            Controls.Add(lblMapFile);
            Controls.Add(lblHint);
            Controls.Add(dgv);
            
            UpdateFilterButtonStyle();
        }

        private void BtnLoadRaw_Click(object sender, EventArgs e)
        {
            if (ofdRaw.ShowDialog() != DialogResult.OK) return;
            
            _rawPath = ofdRaw.FileName;
            lblRawFile.Text = "1번 파일: " + Path.GetFileName(_rawPath);
            
            TryBuildAndBind();
        }

        private void BtnLoadMap_Click(object sender, EventArgs e)
        {
            if (ofdMap.ShowDialog() != DialogResult.OK) return;
            
            _mapPath = ofdMap.FileName;
            lblMapFile.Text = "2번 파일: " + Path.GetFileName(_mapPath);
            
            TryBuildAndBind();
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyGrid();
        }

        private void TryBuildAndBind()
        {
            if (string.IsNullOrWhiteSpace(_rawPath) || string.IsNullOrWhiteSpace(_mapPath))
                return;
            
            try
            {
                _allData = BuildJoinedTable(_rawPath, _mapPath);
                ApplyFilters();
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message,
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable BuildJoinedTable(string rawPath, string mapPath)
        {
            var rawDt = ReadRawExcel(rawPath);
            var mapDict = ReadMapFile(mapPath);
            
            var dt = new DataTable();
            
            int alarmColIndex = rawDt.Columns.Contains("Alarm Name")
                ? rawDt.Columns["Alarm Name"].Ordinal
                : 5;
            
            // Alarm Name까지 원본 컬럼 추가
            for (int i = 0; i <= alarmColIndex; i++)
                dt.Columns.Add(rawDt.Columns[i].ColumnName);
            
            // Alarm Name 바로 오른쪽에 배치
            dt.Columns.Add("Map Text");
            dt.Columns.Add("Root Cause");
            
            // 나머지 원본 컬럼 추가
            for (int i = alarmColIndex + 1; i < rawDt.Columns.Count; i++)
                dt.Columns.Add(rawDt.Columns[i].ColumnName);
            
            foreach (DataRow srcRow in rawDt.Rows)
            {
                var alarmIdRaw = Convert.ToString(srcRow["Alarm Name"]);
                var alarmId = NormalizeMapKey(alarmIdRaw);
                
                string mapText = string.Empty;
                string rootCause = string.Empty;
                
                MapItem item;
                if (!string.IsNullOrWhiteSpace(alarmId) && mapDict.TryGetValue(alarmId, out item))
                {
                    mapText = item.Text;
                    rootCause = item.RootCause;
                }

                var newRow = dt.NewRow();
                
                // Alarm Name까지 복사
                for (int i = 0; i <= alarmColIndex; i++)
                    newRow[i] = srcRow[i];
                
                // 추가 컬럼 삽입
                newRow[alarmColIndex + 1] = mapText;
                newRow[alarmColIndex + 2] = rootCause;
                
                // 나머지 컬럼 복사
                for (int i = alarmColIndex + 1; i < rawDt.Columns.Count; i++)
                    newRow[i + 2] = srcRow[i];
                
                dt.Rows.Add(newRow);
            }

            return dt;
        }

        // 1번 파일: AlarmRawFilterForm과 동일 포맷
        private DataTable ReadRawExcel(string path)
        {
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheets.First();
                var used = ws.RangeUsed();
                var dt = new DataTable();
                
                if (used == null) return dt;
                
                int firstRow = used.FirstRow().RowNumber();
                int lastRow = used.LastRow().RowNumber();
                int lastCol = used.LastColumn().ColumnNumber();
                
                // 헤더
                for (int c = 1; c <= lastCol; c++)
                {
                    string header = ws.Cell(firstRow, c).GetString().Trim();
                    if (string.IsNullOrWhiteSpace(header))
                        header = "Column" + c;
                    
                    string finalHeader = header;
                    int dup = 1;
                    while (dt.Columns.Contains(finalHeader))
                    {
                        finalHeader = header + "_" + dup;
                        dup++;
                    }

                    dt.Columns.Add(finalHeader);
                }

                // F열/G열 헤더 강제 보정
                if (dt.Columns.Count >= 6)
                    dt.Columns[5].ColumnName = "Alarm Name";
                if (dt.Columns.Count >= 7)
                    dt.Columns[6].ColumnName = "Start Time";
                
                // 먼저 원본 행을 전부 메모리에 적재
                var rawList = new List<RawAlarmRow>();
                
                for (int r = firstRow + 1; r <= lastRow; r++)
                {
                    string alarmName = GetCellString(ws.Cell(r, 6)).Trim(); // F
                    if (string.IsNullOrWhiteSpace(alarmName))
                        continue;
                    
                    DateTime? startTime = TryReadDateTime(ws.Cell(r, 7)); // G
                    if (!startTime.HasValue)
                        continue;
                    
                    var values = new string[lastCol];
                    for (int c = 1; c <= lastCol; c++)
                        values[c - 1] = GetCellString(ws.Cell(r, c));
                    
                    rawList.Add(new RawAlarmRow
                    {
                        AlarmName = NormalizeAlarmKey(alarmName),
                        StartTime = startTime.Value,
                        Values = values
                    });
                }

                // Alarm별 + 시간순 정렬 후 중복 제거
                var deduped = new List<RawAlarmRow>();
                
                foreach (var grp in rawList
                    .GroupBy(x => x.AlarmName, StringComparer.OrdinalIgnoreCase))
                {
                    DateTime? lastKept = null;
                    
                    foreach (var item in grp.OrderBy(x => x.StartTime))
                    {
                        if (lastKept.HasValue)
                        {
                            var diffSec = Math.Abs((item.StartTime - lastKept.Value).TotalSeconds);
                            if (diffSec <= DuplicateWithinSeconds)
                                continue; // 중복 제거
                        }

                        deduped.Add(item);
                        lastKept = item.StartTime;
                    }
                }

                // 원래 시간순으로 다시 보고 싶으면 정렬
                foreach (var item in deduped.OrderBy(x => x.StartTime))
                {
                    var row = dt.NewRow();
                    for (int i = 0; i < item.Values.Length; i++)
                        row[i] = item.Values[i];
                    dt.Rows.Add(row);
                }

                return dt;
            }
        }

        private sealed class RawAlarmRow
        {
            public string AlarmName;
            public DateTime StartTime;
            public string[] Values;
        }

        private string NormalizeAlarmKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            
            s = s.Trim();
            
            while (s.Contains("  "))
                s = s.Replace("  ", " ");
            
            return s.ToUpperInvariant();
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
                    "M/d/yyyy H:mm:ss",    "MM/dd/yyyy HH:mm:ss",
                    "yyyy-MM-dd",          "yyyy/MM/dd"
                };

                if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    return dt;
                
                if (DateTime.TryParseExact(s, fmts, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                    return dt;
            }
            catch { }
            
            return null;
        }

        private sealed class MapItem
        {
            public string Text;
            public string RootCause;
        }

        // 2번 파일: B열 ID, J열 Text, K열 Root Cause
        private Dictionary<string, MapItem> ReadMapFile(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            
            if (ext == ".csv")
                return ReadMapCsv(path);
            
            return ReadMapExcel(path);
        }

        private Dictionary<string, MapItem> ReadMapCsv(string path)
        {
            var dict = new Dictionary<string, MapItem>(StringComparer.OrdinalIgnoreCase);
            
            var rows = ReadCsvRows(path);
            if (rows.Count == 0) return dict;
            
            for (int r = 1; r < rows.Count; r++)
            {
                var row = rows[r];
                if (row.Count < 11) continue;
                
                var idRaw = SafeGet(row, 1).Trim();         // B
                var text = SafeGet(row, 9).Trim();          // J
                var rootCause = SafeGet(row, 10).Trim();    // K
                
                var id = NormalizeMapKey(idRaw);
                if (string.IsNullOrWhiteSpace(id)) continue;
                
                dict[id] = new MapItem
                {
                    Text = text,
                    RootCause = rootCause
                };
            }

            return dict;
        }

        private Dictionary<string, MapItem> ReadMapExcel(string path)
        {
            var dict = new Dictionary<string, MapItem>(StringComparer.OrdinalIgnoreCase);
            
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheets.First();
                var used = ws.RangeUsed();
                if (used == null) return dict;
                
                int firstRow = used.FirstRow().RowNumber();
                int lastRow = used.LastRow().RowNumber();
                
                for (int r = firstRow + 1; r <= lastRow; r++)
                {
                    var idRaw = GetCellString(ws.Cell(r, 2)).Trim();        // B
                    var text = GetCellString(ws.Cell(r, 10)).Trim();        // J
                    var rootCause = GetCellString(ws.Cell(r, 11)).Trim();   // K
                    
                    var id = NormalizeMapKey(idRaw);
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    
                    dict[id] = new MapItem
                    {
                        Text = text,
                        RootCause = rootCause
                    };
                }
            }

            return dict;
        }

        private void ApplyFilters()
        {
            if (_allData == null)
            {
                dgv.DataSource = null;
                return;
            }

            var filtered = _allData.Clone();
            
            foreach (DataRow row in _allData.Rows)
            {
                var alarm = Convert.ToString(row["Alarm Name"]);
                if (!PassFilter(alarm)) continue;
                filtered.ImportRow(row);
            }

            dgv.DataSource = filtered;
        }

        private bool PassFilter(string alarm)
        {
            if (string.IsNullOrWhiteSpace(alarm)) return false;
            
            // 선택된 문자열이 포함되면 제외
            if (_filterC1 && alarm.IndexOf("C1", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;
            
            if (_filterC2 && alarm.IndexOf("C2", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;
            
            if (_filterLeft && alarm.IndexOf("Left", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;
            
            if (_filterRight && alarm.IndexOf("Right", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;
            
            return true;
        }

        private void UpdateFilterButtonStyle()
        {
            SetFilterButtonStyle(btnC1, _filterC1);
            SetFilterButtonStyle(btnC2, _filterC2);
            SetFilterButtonStyle(btnLeft, _filterLeft);
            SetFilterButtonStyle(btnRight, _filterRight);
        }

        private void SetFilterButtonStyle(Button btn, bool isOn)
        {
            if (isOn)
            {
                btn.BackColor = System.Drawing.Color.LightCoral;
                btn.FlatStyle = FlatStyle.Popup;
            }
            else
            {
                btn.BackColor = System.Drawing.SystemColors.Control;
                btn.FlatStyle = FlatStyle.Standard;
            }
        }

        private List<List<string>> ReadCsvRows(string path)
        {
            string text;
            
            var bytes = File.ReadAllBytes(path);
            
            bool hasUtf8Bom = bytes.Length >= 3 &&
                              bytes[0] == 0xEF &&
                              bytes[1] == 0xBB &&
                              bytes[2] == 0xBF;
            
            if (hasUtf8Bom)
                text = System.Text.Encoding.UTF8.GetString(bytes);
            else
                text = System.Text.Encoding.GetEncoding(949).GetString(bytes);
            
            var result = new List<List<string>>();
            using (var sr = new StringReader(text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    result.Add(ParseCsvLine(line));
                }
            }

            return result;
        }

        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            
            if (line == null)
            {
                result.Add(string.Empty);
                return result;
            }
            
            var sb = new System.Text.StringBuilder();
            bool inQuotes = false;
            
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                
                if (ch == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (ch == ',' && !inQuotes)
                {
                    result.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    sb.Append(ch);
                }
            }

            result.Add(sb.ToString().Trim());
            return result;
        }

        private string SafeGet(List<string> row, int index)
        {
            if (row == null) return string.Empty;
            if (index < 0 || index >= row.Count) return string.Empty;
            return row[index] ?? string.Empty;
        }

        private string NormalizeMapKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            
            s = s.Trim();
            
            // 앞의 IDE_ 제거
            if (s.StartsWith("IDE_", StringComparison.OrdinalIgnoreCase))
                s = s.Substring(4);
            
            while (s.Contains("  "))
                s = s.Replace("  ", " ");
            
            return s.ToUpperInvariant();
        }

        private void CopyGrid()
        {
            var dt = dgv.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;
            
            var sb = new System.Text.StringBuilder();
            
            // 헤더 포함
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