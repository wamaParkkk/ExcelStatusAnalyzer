using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class ErrorMessagePivotForm : Form
    {
        private Button btnLoad;
        private Button btnLoadTxt;
        private Button btnCopy;
        private Label lblFile;
        private Label lblHint;
        private DataGridView dgv;
        private OpenFileDialog ofd;
        
        private const string ColMessage = "Message";
        private const string TotalColName = "TOTAL";
        
        public ErrorMessagePivotForm()
        {
            BuildUi();
        }
        
        private void BuildUi()
        {
            Text = "Error Message Pivot Form";
            Width = 1200;
            Height = 780;
            
            btnLoad = new Button
            {
                Text = "CSV 파일 불러오기",
                Left = 15,
                Top = 15,
                Width = 170,
                Height = 32
            };
            btnLoad.Click += BtnLoad_Click;
            
            btnLoadTxt = new Button
            {
                Text = "TXT 로그 불러오기",
                Left = 195,
                Top = 15,
                Width = 170,
                Height = 32
            };
            btnLoadTxt.Click += BtnLoadTxt_Click;
            
            btnCopy = new Button
            {
                Text = "데이터 복사",
                Left = 375,
                Top = 15,
                Width = 140,
                Height = 32
            };
            btnCopy.Click += BtnCopy_Click;
            
            lblFile = new Label
            {
                Left = 530,
                Top = 22,
                Width = 620,
                Text = "파일: (없음)"
            };
            
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1120,
                AutoSize = false,
                Text = "CSV 집계 또는 TXT 로그 집계를 선택해서 날짜별 에러 횟수를 표시합니다."
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
                Filter = "CSV/TXT|*.csv;*.CSV;*.txt;*.TXT|All Files|*.*",
                Title = "파일 선택",
                Multiselect = true
            };
            
            Controls.Add(btnLoad);
            Controls.Add(btnLoadTxt);
            Controls.Add(btnCopy);
            Controls.Add(lblFile);
            Controls.Add(lblHint);
            Controls.Add(dgv);
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            ofd.Filter = "CSV|*.csv;*.CSV|All Files|*.*";
            ofd.Title = "Error Raw CSV 파일 선택";
            
            if (ofd.ShowDialog() != DialogResult.OK) return;
            
            try
            {
                var files = ofd.FileNames;
                if (files == null || files.Length == 0) return;
                
                lblFile.Text = "CSV 파일 수: " + files.Length;
                
                var dt = BuildMessageDatePivotTable(files);
                BindResult(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLoadTxt_Click(object sender, EventArgs e)
        {
            ofd.Filter = "TXT|*.txt;*.TXT|All Files|*.*";
            ofd.Title = "TXT 로그 파일 선택";
            
            if (ofd.ShowDialog() != DialogResult.OK) return;
            
            try
            {
                var files = ofd.FileNames;
                if (files == null || files.Length == 0) return;
                
                lblFile.Text = "TXT 파일 수: " + files.Length;
                
                var dt = BuildTxtLogDatePivotTable(files);
                BindResult(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindResult(DataTable dt)
        {
            dgv.DataSource = null;
            dgv.Columns.Clear();
            dgv.AutoGenerateColumns = true;
            dgv.DataSource = dt;
            
            for (int c = 1; c < dgv.Columns.Count; c++)
                dgv.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            
            dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("집계할 데이터가 없습니다.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            CopyGridForExcel();
        }

        private DataTable BuildMessageDatePivotTable(string[] filePaths)
        {
            var merged = new Dictionary<string, Dictionary<DateTime, int>>(StringComparer.OrdinalIgnoreCase);
            var allDates = new HashSet<DateTime>();
            var rx = new Regex(@"\((e_code_[0-9A-Za-z]+|X[0-9A-Za-z]+)\)", RegexOptions.IgnoreCase);
            
            foreach (var path in filePaths.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var rows = ReadCsvRows(path);
                if (rows.Count == 0) continue;
                
                var header = rows[0];
                
                int messageCol = -1;
                int triggerDateCol = -1;
                
                for (int i = 0; i < header.Count; i++)
                {
                    var h = (header[i] ?? string.Empty).Trim().Replace("\uFEFF", "");
                    
                    if (messageCol < 0 &&
                        (string.Equals(h, "Message", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(h, "Message(s)", StringComparison.OrdinalIgnoreCase)))
                    {
                        messageCol = i;
                    }
                    
                    if (triggerDateCol < 0 &&
                        (string.Equals(h, "TriggerDate", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(h, "Trigger Date", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(h, "DATE", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(h, "Date", StringComparison.OrdinalIgnoreCase)))
                    {
                        triggerDateCol = i;
                    }
                }

                if (messageCol < 0) continue;
                
                var fileDateFromName = ExtractDateFromFileName(path);
                
                for (int r = 1; r < rows.Count; r++)
                {
                    var row = rows[r];
                    if (row == null || row.Count <= messageCol) continue;
                    
                    string message = ExtractMessageOnly(row, messageCol, rx);
                    if (string.IsNullOrWhiteSpace(message)) continue;
                    
                    DateTime? rowDate = null;
                    
                    if (triggerDateCol >= 0 && row.Count > triggerDateCol)
                        rowDate = ParseRowDate(row[triggerDateCol]);
                    
                    if (!rowDate.HasValue)
                        rowDate = fileDateFromName;
                    
                    if (!rowDate.HasValue)
                        continue;
                    
                    allDates.Add(rowDate.Value.Date);
                    
                    Dictionary<DateTime, int> inner;
                    if (!merged.TryGetValue(message, out inner))
                    {
                        inner = new Dictionary<DateTime, int>();
                        merged[message] = inner;
                    }
                    
                    int cur;
                    if (!inner.TryGetValue(rowDate.Value.Date, out cur)) cur = 0;
                    inner[rowDate.Value.Date] = cur + 1;
                }
            }

            return BuildPivotTable(merged, allDates);
        }

        private DataTable BuildTxtLogDatePivotTable(string[] filePaths)
        {
            var merged = new Dictionary<string, Dictionary<DateTime, int>>(StringComparer.OrdinalIgnoreCase);
            var allDates = new HashSet<DateTime>();
            
            foreach (var path in filePaths.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var rows = ReadDelimitedRows(path);
                if (rows.Count < 3) continue;
                
                // 1행: "*** Error Log File Header (CSV File)"
                // 2행: "*** Parsing List : OCCUR_TIME,ERROR_NO,EXPLANATION,..."
                var headerRow = ParseTxtHeaderRow(rows);
                if (headerRow.Count == 0) continue;
                
                int occurTimeCol = FindExactColumnIndex(headerRow, "OCCUR_TIME");
                int explanationCol = FindExactColumnIndex(headerRow, "EXPLANATION");
                
                if (occurTimeCol < 0 || explanationCol < 0) continue;
                
                var fileDateFromName = ExtractDateFromFileName(path);
                
                // 데이터는 3행부터
                for (int r = 2; r < rows.Count; r++)
                {
                    var row = rows[r];
                    if (row == null || row.Count == 0) continue;
                    if (row.Count <= Math.Max(occurTimeCol, explanationCol)) continue;
                    
                    var explanation = (row[explanationCol] ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(explanation)) continue;
                    
                    DateTime? rowDate = ParseTxtOccurDate(row[occurTimeCol]);
                    if (!rowDate.HasValue)
                        rowDate = fileDateFromName;
                    
                    if (!rowDate.HasValue)
                        continue;
                    
                    allDates.Add(rowDate.Value.Date);
                    
                    Dictionary<DateTime, int> inner;
                    if (!merged.TryGetValue(explanation, out inner))
                    {
                        inner = new Dictionary<DateTime, int>();
                        merged[explanation] = inner;
                    }
                    
                    int cur;
                    if (!inner.TryGetValue(rowDate.Value.Date, out cur)) cur = 0;
                    inner[rowDate.Value.Date] = cur + 1;
                }
            }

            return BuildPivotTable(merged, allDates);
        }

        private DataTable BuildPivotTable(Dictionary<string, Dictionary<DateTime, int>> merged, HashSet<DateTime> allDates)
        {
            var dateList = new List<DateTime>();
            
            if (allDates.Count > 0)
            {
                var minDate = allDates.Min();
                var maxDate = allDates.Max();
                
                var today = DateTime.Today;
                if (today > maxDate)
                    maxDate = today;
                
                for (var d = minDate; d <= maxDate; d = d.AddDays(1))
                    dateList.Add(d);
            }

            var dtResult = new DataTable();
            dtResult.Columns.Add(ColMessage);
            for (int i = 0; i < dateList.Count; i++)
                dtResult.Columns.Add(dateList[i].ToString("yyyy-MM-dd"), typeof(int));
            dtResult.Columns.Add(TotalColName, typeof(int));
            
            foreach (var kv in merged)
            {
                var row = dtResult.NewRow();
                row[0] = kv.Key;
                
                int total = 0;
                for (int i = 0; i < dateList.Count; i++)
                {
                    int v;
                    if (!kv.Value.TryGetValue(dateList[i], out v)) v = 0;
                    row[i + 1] = v;
                    total += v;
                }

                row[TotalColName] = total;
                dtResult.Rows.Add(row);
            }

            dtResult.DefaultView.Sort = "[" + TotalColName + "] DESC, [" + ColMessage + "] ASC";
            return dtResult.DefaultView.ToTable();
        }

        private List<string> ParseTxtHeaderRow(List<List<string>> rows)
        {
            if (rows == null || rows.Count < 2) return new List<string>();
            
            // 두 번째 줄:
            // *** Parsing List : OCCUR_TIME,ERROR_NO,EXPLANATION,...
            var lineParts = rows[1];
            if (lineParts == null || lineParts.Count == 0) return new List<string>();
            
            // ReadDelimitedRows에서 이미 ParseCsvLine을 한번 거쳤기 때문에
            // 다시 ","로 합쳐서 원문처럼 복원
            var line = string.Join(",", lineParts);
            
            var idx = line.IndexOf(':');
            if (idx >= 0)
                line = line.Substring(idx + 1);
            
            return ParseCsvLine(line);
        }

        private int FindExactColumnIndex(List<string> header, string colName)
        {
            for (int i = 0; i < header.Count; i++)
            {
                if (string.Equals((header[i] ?? string.Empty).Trim(), colName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
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
                text = Encoding.UTF8.GetString(bytes);
            else
                text = Encoding.GetEncoding(949).GetString(bytes);
            
            var result = new List<List<string>>();
            using (var sr = new StringReader(text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                    result.Add(ParseCsvLine(line));
            }

            return result;
        }

        private List<List<string>> ReadDelimitedRows(string path)
        {
            string text;
            var bytes = File.ReadAllBytes(path);
            
            bool hasUtf8Bom = bytes.Length >= 3 &&
                              bytes[0] == 0xEF &&
                              bytes[1] == 0xBB &&
                              bytes[2] == 0xBF;
            
            if (hasUtf8Bom)
                text = Encoding.UTF8.GetString(bytes);
            else
                text = Encoding.GetEncoding(949).GetString(bytes);
            
            var result = new List<List<string>>();
            using (var sr = new StringReader(text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                    result.Add(ParseCsvLine(line));
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

            var sb = new StringBuilder();
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

        private DateTime? ExtractDateFromFileName(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            
            DateTime dt;
            
            var m1 = Regex.Match(name, @"(20\d{2}-\d{2}-\d{2})");
            if (m1.Success && DateTime.TryParseExact(m1.Groups[1].Value, "yyyy-MM-dd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            var m2 = Regex.Match(name, @"(20\d{6})");
            if (m2.Success && DateTime.TryParseExact(m2.Groups[1].Value, "yyyyMMdd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            return null;
        }

        private DateTime? ParseRowDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            
            s = s.Trim();
            
            DateTime dt;
            
            if (DateTime.TryParseExact(s,
                new[] { "yyyy-MM-dd", "yyyy/M/d", "MM/dd/yyyy", "M/d/yyyy", "yyyyMMdd" },
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out dt))
                return dt.Date;
            
            if (DateTime.TryParse(s, out dt))
                return dt.Date;
            
            return null;
        }

        private DateTime? ParseTxtOccurDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            
            s = s.Trim();
            
            DateTime dt;
            
            if (DateTime.TryParseExact(s,
                new[] { "yyyy-MM-dd HH:mm:ss.fff", "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss.fff", "yyyy/MM/dd HH:mm:ss" },
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out dt))
                return dt.Date;
            
            if (DateTime.TryParse(s, out dt))
                return dt.Date;
            
            return null;
        }

        private string ExtractMessageOnly(List<string> row, int messageCol, Regex rx)
        {
            if (row == null || row.Count <= messageCol) return string.Empty;
            
            string message = (row[messageCol] ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(message)) return string.Empty;
            
            if (rx.IsMatch(message))
                return message;
            
            var sb = new StringBuilder(message);
            
            for (int i = messageCol + 1; i < row.Count; i++)
            {
                var part = (row[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(part)) continue;
                
                sb.Append(", ").Append(part);
                
                if (rx.IsMatch(sb.ToString()))
                    return sb.ToString();
            }

            return string.Empty;
        }

        private void CopyGridForExcel()
        {
            if (dgv == null || dgv.Rows.Count == 0) return;
            
            var oldMode = dgv.ClipboardCopyMode;
            var oldMultiSelect = dgv.MultiSelect;
            var oldSelectionMode = dgv.SelectionMode;
            var hiddenCols = new List<DataGridViewColumn>();
            
            try
            {
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (string.Equals(col.Name, TotalColName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(col.HeaderText, TotalColName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (col.Visible)
                        {
                            col.Visible = false;
                            hiddenCols.Add(col);
                        }
                    }
                }

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
                foreach (var col in hiddenCols)
                    col.Visible = true;
                
                dgv.ClearSelection();
                dgv.MultiSelect = oldMultiSelect;
                dgv.SelectionMode = oldSelectionMode;
                dgv.ClipboardCopyMode = oldMode;
            }
        }
    }
}