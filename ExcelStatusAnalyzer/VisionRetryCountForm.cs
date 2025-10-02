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
    public partial class VisionRetryCountForm : Form
    {
        private Button btnLoad;
        private DataGridView dgvResult;
        private Label lblFile, lblSummary;
        private OpenFileDialog ofd;

        public VisionRetryCountForm()
        {
            BuildUi();
        }
        
        private void BuildUi()
        {
            Text = "Vision Retry Count (LEFT/RIGHT 횟수 분포 집계)";
            Width = 800;
            Height = 600;

            btnLoad = new Button { Text = "파일 불러오기 (.csv/.xlsx/xls)", Left = 15, Top = 15, Width = 220, Height = 32 };
            btnLoad.Click += BtnLoad_Click;

            lblFile = new Label { Left = 250, Top = 22, Width = 500, Text = "파일: (없음)" };

            dgvResult = new DataGridView
            {
                Left = 15,
                Top = 60,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 120,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = false,
                EditMode = DataGridViewEditMode.EditProgrammatically,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = true,                
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToResizeColumns = true
            };

            dgvResult.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) e.SuppressKeyPress = true;
            };

            lblSummary = new Label
            {
                Left = 15,
                Top = ClientSize.Height - 50,
                Width = ClientSize.Width - 30,
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right,
                Text = ""
            };

            ofd = new OpenFileDialog
            {
                Filter = "All Supported|*.csv;*.xlsx;*.xls|CSV|*.csv|Excel|*.xlsx;*.xls",
                Title = "집계 대상 파일 선택"
            };

            Controls.Add(btnLoad);
            Controls.Add(lblFile);
            Controls.Add(dgvResult);
            Controls.Add(lblSummary);
        }
        
        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var path = ofd.FileName;
                lblFile.Text = "파일: " + Path.GetFileName(path);

                // D열=LEFT/RIGHT, F열=숫자. 헤더 없음.
                List<DirCount> rows;
                var ext = Path.GetExtension(path).ToLowerInvariant();
                switch (ext)
                {
                    case ".csv":
                        rows = LoadCsvDF(path);
                        break;
                    case ".xlsx":
                    case ".xls":
                        rows = LoadExcelDF(path);
                        break;
                    default:
                        throw new Exception("지원하지 않는 확장자입니다.");
                }

                // --- 집계: k = max(F-1, 0) 전부 카운트 (제한 없음) ---
                var leftCounts = new Dictionary<int, int>();
                var rightCounts = new Dictionary<int, int>();
                
                for (int i = 0; i < rows.Count; i++)
                {
                    string dir = (rows[i].Dir ?? string.Empty).Trim();
                    int k = rows[i].Count - 1;
                    if (k < 0) k = 0;
                    
                    if (string.Equals(dir, "LEFT", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!leftCounts.ContainsKey(k)) leftCounts[k] = 0;
                        leftCounts[k]++;
                    }
                    else if (string.Equals(dir, "RIGHT", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!rightCounts.ContainsKey(k)) rightCounts[k] = 0;
                        rightCounts[k]++;
                    }
                    // 그 외 값은 무시
                }

                // --- 결과 테이블 빌드(회수 전부) ---
                var dt = BuildDynamicTable(leftCounts, rightCounts);
                dgvResult.DataSource = dt;
                
                // 숫자 가독성
                if (dgvResult.Columns.Count >= 4)
                {
                    dgvResult.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dgvResult.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dgvResult.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                
                // 요약 라벨
                int leftTotal = 0, rightTotal = 0, maxK = 0;
                foreach (var kv in leftCounts) { leftTotal += kv.Value; if (kv.Key > maxK) maxK = kv.Key; }
                foreach (var kv in rightCounts) { rightTotal += kv.Value; if (kv.Key > maxK) maxK = kv.Key; }
                
                lblSummary.Text = "Left 합계: " + leftTotal.ToString("#,0") + 
                    "  |  Right 합계: " + rightTotal.ToString("#,0") + 
                    "  |  회수 범위: 0~" + maxK;
            }
            catch (Exception ex)
            {
                MessageBox.Show("집계 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<DirCount> LoadCsvDF(string path)
        {
            var list = new List<DirCount>();
            foreach (var fields in ReadCsv(path, null))
            {
                if (fields.Length < 6) continue;
                string dVal = (fields[3] ?? "").Trim();
                if (string.IsNullOrWhiteSpace(dVal)) continue;

                int f = ParseIntString(fields[5]);
                list.Add(new DirCount { Dir = dVal, Count = f });
            }
            return list;
        }

        // === 로더들: D열(4), F열(6)만 읽는다 ===
        private List<DirCount> LoadExcelDF(string path)
        {
            var list = new List<DirCount>();
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheets.First(); // 첫 번째 시트
                var used = ws.RangeUsed();
                if (used == null) return list;
                
                foreach (var row in used.Rows())
                {
                    string dVal = CellToString(row.Cell(4)); // D열
                    if (string.IsNullOrWhiteSpace(dVal)) continue;
                    
                    var c6 = row.Cell(6); // F열
                    int f = ParseIntCell(c6);
                    list.Add(new DirCount { Dir = dVal, Count = f });
                }
            }
            return list;
        }        

        // === CSV 파서(따옴표, 콤마 처리) ===
        private IEnumerable<string[]> ReadCsv(string path, Encoding enc)
        {
            if (enc == null) enc = Encoding.UTF8;
            using (var sr = new StreamReader(path, enc))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    yield return SplitCsvLine(line).ToArray();
                }
            }
        }

        private List<string> SplitCsvLine(string line)
        {
            var list = new List<string>();
            if (line == null) return list;
            
            var sb = new StringBuilder();
            bool inQuotes = false;
            
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                
                if (inQuotes)
                {
                    if (ch == '\"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '\"')
                        {
                            sb.Append('\"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }
                else
                {
                    if (ch == ',')
                    {
                        list.Add(sb.ToString());
                        sb.Clear();
                    }
                    else if (ch == '\"')
                    {
                        inQuotes = true;
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }
            }
            list.Add(sb.ToString());
            return list;
        }

        private int ParseIntString(string s)
        {
            s = (s ?? "").Trim();

            int iv;
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out iv))
                return iv;

            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return (int)Math.Round(d);

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                return (int)Math.Round(d);

            s = s.Replace(",", "").Replace(" ", "");
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return (int)Math.Round(d);

            return 0;
        }

        private string CellToString(IXLCell cell)
        {
            if (cell == null) return string.Empty;
            switch (cell.DataType)
            {
                case XLDataType.DateTime:
                    return cell.GetDateTime().ToString("yyyy-MM-dd HH:mm:ss");
                case XLDataType.Boolean:
                    return cell.GetBoolean() ? "TRUE" : "FALSE";
                case XLDataType.Number:
                    return cell.GetDouble().ToString(CultureInfo.InvariantCulture);
                default:
                    return cell.GetString().Trim();
            }
        }

        private int ParseIntCell(IXLCell cell)
        {
            if (cell.DataType == XLDataType.Number)
            {
                return (int)Math.Round(cell.GetDouble());
            }
            var s = CellToString(cell);
            return ParseIntString(s);
        }

        private DataTable BuildDynamicTable(Dictionary<int, int> left, Dictionary<int, int> right)
        {
            var dt = new DataTable();
            dt.Columns.Add("회수");
            dt.Columns.Add("Left", typeof(int));
            dt.Columns.Add("Right", typeof(int));
            dt.Columns.Add("합계", typeof(int));
            
            // 최소는 0회, 최대는 등장한 회수의 최댓값
            int minK = 0;
            int maxK = 0;
            
            foreach (var kv in left)
                if (kv.Key > maxK) maxK = kv.Key;
            
            foreach (var kv in right)
                if (kv.Key > maxK) maxK = kv.Key;
            
            // 데이터가 하나도 없으면 합계만 0으로 표시하고 반환
            if (left.Count == 0 && right.Count == 0)
            {
                dt.Rows.Add("합계", 0, 0, 0);
                return dt;
            }
            
            // 0 ~ maxK까지 빠짐없이 채워 넣기 (없는 회수는 0)
            for (int k = minK; k <= maxK; k++)
            {
                int l = left.ContainsKey(k) ? left[k] : 0;
                int r = right.ContainsKey(k) ? right[k] : 0;
                dt.Rows.Add(k + "회", l, r, l + r);
            }
            
            // 맨 아래 합계 행
            int lsum = 0; foreach (var kv in left) lsum += kv.Value;
            int rsum = 0; foreach (var kv in right) rsum += kv.Value;
            dt.Rows.Add("합계", lsum, rsum, lsum + rsum);
            
            return dt;
        }
    }
}