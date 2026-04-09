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
                Text = "파일 불러오기 (.csv 여러개 선택)",
                Left = 15,
                Top = 15,
                Width = 220,
                Height = 32
            };
            btnLoad.Click += BtnLoad_Click;
            
            btnCopy = new Button
            {
                Text = "데이터 복사",
                Left = 245,
                Top = 15,
                Width = 140,
                Height = 32
            };
            btnCopy.Click += BtnCopy_Click;
            
            lblFile = new Label
            {
                Left = 400,
                Top = 22,
                Width = 760,
                Text = "파일: (없음)"
            };
            
            lblHint = new Label
            {
                Left = 15,
                Top = 52,
                Width = 1120,
                AutoSize = false,
                Text = "여러 개의 일자별 CSV 파일을 선택하여 Message에 e_code_XX 패턴이 포함된 Error만 날짜별로 집계합니다."
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
                Filter = "CSV|*.csv;*.CSV|All Files|*.*",
                Title = "Error Raw CSV 파일 선택",
                Multiselect = true
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
                var files = ofd.FileNames;
                if (files == null || files.Length == 0) return;
                
                lblFile.Text = "파일 수: " + files.Length;
                
                var dt = BuildMessageDatePivotTable(files);
                
                dgv.DataSource = null;
                dgv.Columns.Clear();
                dgv.AutoGenerateColumns = true;
                dgv.DataSource = dt;
                
                for (int c = 1; c < dgv.Columns.Count; c++)
                    dgv.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("e_code_XX 패턴이 포함된 Message가 없습니다.",
                        "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            CopyGrid();
        }

        private DataTable BuildMessageDatePivotTable(string[] filePaths)
        {
            // message -> (date -> count)
            var merged = new Dictionary<string, Dictionary<DateTime, int>>(StringComparer.OrdinalIgnoreCase);
            var allDates = new HashSet<DateTime>();
            
            foreach (var path in filePaths)
            {
                var fileDate = ExtractDateFromFileName(path);
                if (!fileDate.HasValue)
                    continue;
                
                var rows = ReadCsvRows(path);
                if (rows.Count == 0) continue;
                
                var header = rows[0];
                int messageCol = FindColumnIndex(header, ColMessage);
                if (messageCol < 0) continue;
                
                var rx = new Regex(@"e_code_[0-9A-Za-z]+", RegexOptions.IgnoreCase);
                
                // 2행부터 데이터
                for (int r = 1; r < rows.Count; r++)
                {
                    var row = rows[r];
                    if (row.Count <= messageCol) continue;
                    
                    var message = (row[messageCol] ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(message)) continue;
                    if (!rx.IsMatch(message)) continue;
                    
                    allDates.Add(fileDate.Value.Date);
                    
                    Dictionary<DateTime, int> inner;
                    if (!merged.TryGetValue(message, out inner))
                    {
                        inner = new Dictionary<DateTime, int>();
                        merged[message] = inner;
                    }
                    
                    int cur;
                    if (!inner.TryGetValue(fileDate.Value.Date, out cur)) cur = 0;
                    inner[fileDate.Value.Date] = cur + 1;
                }
            }

            List<DateTime> dateList = new List<DateTime>();
            
            if (allDates.Count > 0)
            {
                var minDate = allDates.Min();
                var maxDate = allDates.Max();
                
                for (var d = minDate; d <= maxDate; d = d.AddDays(1))
                    dateList.Add(d);
            }

            var dt = new DataTable();
            dt.Columns.Add(ColMessage);
            for (int i = 0; i < dateList.Count; i++)
                dt.Columns.Add(dateList[i].ToString("yyyy-MM-dd"), typeof(int));
            dt.Columns.Add(TotalColName, typeof(int));
            
            foreach (var kv in merged)
            {
                var row = dt.NewRow();
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
                dt.Rows.Add(row);
            }

            dt.DefaultView.Sort = "[" + TotalColName + "] DESC, [Message] ASC";
            return dt.DefaultView.ToTable();
        }

        private DateTime? ExtractDateFromFileName(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            
            DateTime dt;
            
            // yyyy-MM-dd (뒤에 _, -, 공백 등이 와도 잡히게 \b 제거)
            var m1 = Regex.Match(name, @"(20\d{2}-\d{2}-\d{2})");
            if (m1.Success && DateTime.TryParseExact(m1.Groups[1].Value, "yyyy-MM-dd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            // yyyyMMdd
            var m2 = Regex.Match(name, @"(20\d{6})");
            if (m2.Success && DateTime.TryParseExact(m2.Groups[1].Value, "yyyyMMdd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            return null;
        }

        private List<List<string>> ReadCsvRows(string path)
        {
            string text;
            
            var bytes = File.ReadAllBytes(path);
            
            // UTF-8 BOM 여부 확인
            bool hasUtf8Bom = bytes.Length >= 3 &&
                              bytes[0] == 0xEF &&
                              bytes[1] == 0xBB &&
                              bytes[2] == 0xBF;
            
            if (hasUtf8Bom)
                text = Encoding.UTF8.GetString(bytes);
            else
                text = Encoding.GetEncoding(949).GetString(bytes); // cp949 우선
            
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

        private int FindColumnIndex(List<string> header, string colName)
        {
            for (int i = 0; i < header.Count; i++)
            {
                if (string.Equals((header[i] ?? string.Empty).Trim(), colName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private void CopyGrid()
        {
            var dt = dgv.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0) return;
            
            var sb = new StringBuilder();
            
            // TOTAL 제외한 컬럼 인덱스 수집
            var colIndexes = new List<int>();
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var colName = dt.Columns[c].ColumnName;
                if (string.Equals(colName, TotalColName, StringComparison.OrdinalIgnoreCase))
                    continue;
                
                colIndexes.Add(c);
            }

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int i = 0; i < colIndexes.Count; i++)
                {
                    if (i > 0) sb.Append('\t');
                    sb.Append(Convert.ToString(dt.Rows[r][colIndexes[i]]));
                }
                sb.Append('\n');
            }

            Clipboard.Clear();
            Clipboard.SetText(sb.ToString());
        }
    }
}