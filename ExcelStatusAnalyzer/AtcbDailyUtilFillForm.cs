using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class AtcbDailyUtilFillForm : Form
    {
        private Button btnLoadAndApply;
        private DataGridView dgv;
        private Label lblFile;
        private TextBox txtLog;
        private OpenFileDialog ofd;
        
        private const string TargetPath = @"C:\Users\156607\Amkor_Project\Document\장비 가동률 데이터\ATCB\ATCB_Daily 가동현황.xlsx";
        
        // Source CSV 컬럼 (1-based)
        private const int SrcColEquip = 4;     // D열: 장비번호
        private const int SrcColRunTime = 14;  // N열: 가동 시간_TTL
                                               
        // Target Excel (1-based)
        private const int TargetSheetIndex = 1; // 첫 번째 시트
        private const int TargetColDate = 1;    // A열
        private const int TargetColATCB01 = 5;  // E열
        private const int TargetColATCB02 = 12; // L열
        private const int TargetRowStart = 2;   // 데이터 시작 행 (필요시 조정)
        
        public AtcbDailyUtilFillForm()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            BuildUi();
        }
        
        private void BuildUi()
        {
            Text = "ATCB Daily 가동률 자동 입력";
            Width = 1200;
            Height = 780;
            
            btnLoadAndApply = new Button
            {
                Left = 15,
                Top = 15,
                Width = 280,
                Height = 34,
                Text = "CSV 불러오기 + Target 자동 입력"
            };
            btnLoadAndApply.Click += BtnLoadAndApply_Click;
            
            lblFile = new Label
            {
                Left = 310,
                Top = 23,
                Width = 850,
                Text = "파일: (없음)"
            };

            dgv = new DataGridView
            {
                Left = 15,
                Top = 60,
                Width = ClientSize.Width - 30,
                Height = 420,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                ScrollBars = ScrollBars.Both
            };

            txtLog = new TextBox
            {
                Left = 15,
                Top = 495,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 510,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };

            ofd = new OpenFileDialog
            {
                Filter = "CSV|*.csv;*.CSV",
                Title = "ATCB Source CSV 선택"
            };

            Controls.Add(btnLoadAndApply);
            Controls.Add(lblFile);
            Controls.Add(dgv);
            Controls.Add(txtLog);
        }

        private void BtnLoadAndApply_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;
            
            txtLog.Clear();
            
            try
            {
                var srcPath = ofd.FileName;
                lblFile.Text = "파일: " + Path.GetFileName(srcPath);
                
                var reportDate = ExtractDateFromFileName(srcPath);
                if (!reportDate.HasValue)
                    throw new Exception("CSV 파일명에서 날짜(yyyyMMdd 또는 yyyy-MM-dd)를 찾지 못했습니다.");
                
                Log("[날짜] " + reportDate.Value.ToString("yyyy-MM-dd"));
                
                var utilMap = ReadSourceCsv(srcPath);
                
                if (!utilMap.ContainsKey("ATCB-01"))
                    utilMap["ATCB-01"] = 0.00;
                
                if (!utilMap.ContainsKey("ATCB-02"))
                    utilMap["ATCB-02"] = 0.00;
                
                Log("[계산] ATCB-01 = " + utilMap["ATCB-01"].ToString("0.00"));
                Log("[계산] ATCB-02 = " + utilMap["ATCB-02"].ToString("0.00"));
                
                if (!File.Exists(TargetPath))
                    throw new Exception("Target 파일이 존재하지 않습니다: " + TargetPath);
                
                ApplyToTarget(TargetPath, reportDate.Value.Date, utilMap);
                
                dgv.DataSource = BuildPreviewTable(reportDate.Value.Date, utilMap);
                
                MessageBox.Show("완료! Target 엑셀에 자동 입력했습니다.",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message,
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("[ERROR] " + ex);
            }
        }

        private DateTime? ExtractDateFromFileName(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            
            DateTime dt;
            
            var m1 = Regex.Match(name, @"(20\d{6})");
            if (m1.Success && DateTime.TryParseExact(m1.Groups[1].Value, "yyyyMMdd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            var m2 = Regex.Match(name, @"(20\d{2}-\d{2}-\d{2})");
            if (m2.Success && DateTime.TryParseExact(m2.Groups[1].Value, "yyyy-MM-dd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            return null;
        }

        private Dictionary<string, double> ReadSourceCsv(string path)
        {
            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            
            var lines = File.ReadAllLines(path, Encoding.GetEncoding(949));
            if (lines.Length <= 1) return map;
            
            // 1행 헤더, 2행부터 데이터
            for (int i = 1; i < lines.Length; i++)
            {
                var cols = ParseCsvLine(lines[i]);
                if (cols.Count < SrcColRunTime) continue;
                
                string equip = SafeGet(cols, SrcColEquip - 1).Trim();
                if (!string.Equals(equip, "ATCB-01", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(equip, "ATCB-02", StringComparison.OrdinalIgnoreCase))
                    continue;
                
                double runMinutes = ParseDouble(SafeGet(cols, SrcColRunTime - 1));
                double utilPercent = 0.00;
                
                if (runMinutes > 0)
                    utilPercent = Math.Round((runMinutes / 1440.0) * 100.0, 2, MidpointRounding.AwayFromZero);
                
                map[equip] = utilPercent;
            }

            return map;
        }

        private void ApplyToTarget(string targetPath, DateTime reportDate, Dictionary<string, double> utilMap)
        {
            using (var wb = new XLWorkbook(targetPath))
            {
                var ws = wb.Worksheet(TargetSheetIndex);
                
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? TargetRowStart;
                bool foundRow = false;
                
                for (int r = TargetRowStart; r <= lastRow; r++)
                {
                    var dt = TryReadExcelDate(ws.Cell(r, TargetColDate));
                    if (!dt.HasValue) continue;
                    
                    if (dt.Value.Date != reportDate.Date) continue;
                    
                    ws.Cell(r, TargetColATCB01).Value = utilMap.ContainsKey("ATCB-01") ? utilMap["ATCB-01"] : 0.00;
                    ws.Cell(r, TargetColATCB02).Value = utilMap.ContainsKey("ATCB-02") ? utilMap["ATCB-02"] : 0.00;
                    
                    ws.Cell(r, TargetColATCB01).Style.NumberFormat.Format = "0.00";
                    ws.Cell(r, TargetColATCB02).Style.NumberFormat.Format = "0.00";
                    
                    foundRow = true;
                    Log("[입력 완료] Row " + r + " / 날짜 " + reportDate.ToString("yyyy-MM-dd"));
                    break;
                }

                if (!foundRow)
                    throw new Exception("Target 엑셀 A열에서 날짜 " + reportDate.ToString("yyyy-MM-dd") + " 를 찾지 못했습니다.");
                
                wb.Save();
            }
        }

        private DataTable BuildPreviewTable(DateTime reportDate, Dictionary<string, double> utilMap)
        {
            var dt = new DataTable();
            dt.Columns.Add("Date");
            dt.Columns.Add("Equip");
            dt.Columns.Add("Util(%)", typeof(double));
            
            var row1 = dt.NewRow();
            row1["Date"] = reportDate.ToString("yyyy-MM-dd");
            row1["Equip"] = "ATCB-01";
            row1["Util(%)"] = utilMap.ContainsKey("ATCB-01") ? utilMap["ATCB-01"] : 0.00;
            dt.Rows.Add(row1);
            
            var row2 = dt.NewRow();
            row2["Date"] = reportDate.ToString("yyyy-MM-dd");
            row2["Equip"] = "ATCB-02";
            row2["Util(%)"] = utilMap.ContainsKey("ATCB-02") ? utilMap["ATCB-02"] : 0.00;
            dt.Rows.Add(row2);
            
            return dt;
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

        private string SafeGet(List<string> cols, int index)
        {
            if (cols == null) return string.Empty;
            if (index < 0 || index >= cols.Count) return string.Empty;
            return cols[index] ?? string.Empty;
        }

        private double ParseDouble(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0.0;
            
            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return d;
            
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                return d;
            
            return 0.0;
        }

        private DateTime? TryReadExcelDate(IXLCell cell)
        {
            if (cell == null) return null;
            
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
            if (DateTime.TryParse(s, out dt))
                return dt;

            return null;
        }

        private void Log(string msg)
        {
            txtLog.AppendText(msg + Environment.NewLine);
        }
    }
}