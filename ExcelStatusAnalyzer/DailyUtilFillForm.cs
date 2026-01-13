using ClosedXML.Excel;
using ExcelDataReader;
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
    public partial class DailyUtilFillForm : Form
    {
        private Button btnLoadAndApply;
        private DataGridView dgv;
        private Label lblFile;
        private TextBox txtLog;
        private OpenFileDialog ofd;
        
        // 고정 타겟 경로
        private const string TargetPath = @"C:\Users\156607\Amkor_Project\Document\장비 가동률 데이터\FL_CL\FL_CL Daily 가동현황.xlsx";
        
        // 타겟 시트/컬럼
        private const int TargetSheetIndex = 1; // 첫번째 시트
        private const int TargetRowStart = 4;   // 4행부터
        private const int ColDate = 2;          // B
        private const int ColEquip = 3;         // C
        private const int ColUtil = 6;          // F
                                                
        // 소스 컬럼 (요구사항)
        private const int SrcColEquip = 1;      // A
        private const int SrcColUtil = 4;       // D
        
        private readonly HashSet<string> _allowedEquip = new HashSet<string>(
            Enumerable.Range(1, 16).Select(i => $"BA-FC-{i:00}"),
            StringComparer.OrdinalIgnoreCase);
        
        public DailyUtilFillForm()
        {
            // (일부 환경에서 필요) ExcelDataReader 코드페이지 등록
            try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); } catch { }
            BuildUi();
        }
        
        private void BuildUi()
        {
            Text = "가동률 자동 채움 (Source -> FL_CL Daily 가동현황.xlsx)";
            Width = 1200;
            Height = 780;
            
            btnLoadAndApply = new Button
            {
                Left = 15,
                Top = 15,
                Width = 260,
                Height = 34,
                Text = "소스 엑셀 불러오기 + 타겟에 적용"
            };
            btnLoadAndApply.Click += BtnLoadAndApply_Click;
            
            lblFile = new Label
            {
                Left = 290,
                Top = 23,
                Width = 880,
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
                Filter = "Excel|*.xlsx;*.xls",
                Title = "가동률 소스 파일 선택"
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
                
                // 1) 소스에서 날짜(yyyymmdd) 추출
                var reportDate = ExtractDateFromFileName(srcPath);
                if (!reportDate.HasValue)
                    throw new Exception("소스 파일명에서 yyyymmdd 날짜를 찾지 못했습니다. 예: ..._20251201.xls");
                
                Log($"[소스 날짜] {reportDate.Value:yyyy-MM-dd}");
                
                // 2) 소스 엑셀 읽어서 (장비명 -> Util) 맵 생성
                var utilMap = ReadSourceUtilMap(srcPath);
                if (utilMap.Count == 0)
                    throw new Exception("소스 파일에서 (A=장비명, D=가동률) 데이터를 찾지 못했습니다.");
                
                Log($"[소스] 추출 장비 수: {utilMap.Count}");
                
                // 3) 타겟 엑셀에 채우기
                if (!File.Exists(TargetPath))
                    throw new Exception("타겟 파일이 존재하지 않습니다: " + TargetPath);
                
                var result = ApplyToTarget(TargetPath, reportDate.Value.Date, utilMap);
                
                // 4) 미리보기 테이블
                dgv.DataSource = BuildPreviewTable(utilMap, result);
                
                Log($"[적용 완료] 업데이트 행 수: {result.UpdatedRows}");
                
                if (result.NotFoundInTarget.Count > 0)
                {
                    Log("[타겟에 없는 장비/날짜 매칭]");
                    foreach (var s in result.NotFoundInTarget)
                        Log("  - " + s);
                }

                if (result.NotFoundInSource.Count > 0)
                {
                    Log("[소스에 없는 장비(BA-FC-01~16 중)]");
                    foreach (var s in result.NotFoundInSource)
                        Log("  - " + s);
                }

                MessageBox.Show("완료! 타겟 엑셀에 가동률을 채웠습니다.\n(엑셀이 열려있으면 저장이 실패할 수 있어요.)",
                    "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("[ERROR] " + ex);
            }
        }

        // ---------------------------
        // 소스 파일명에서 yyyymmdd 추출
        // ---------------------------
        private static DateTime? ExtractDateFromFileName(string path)
        {
            var name = Path.GetFileNameWithoutExtension(path);
            var m = Regex.Match(name, @"(20\d{6})"); // 20xxxxxx
            if (!m.Success) return null;
            
            DateTime dt;
            if (DateTime.TryParseExact(m.Groups[1].Value, "yyyyMMdd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            return null;
        }

        // ---------------------------
        // 소스에서 A=장비명, D=가동률 읽기
        // ---------------------------
        private Dictionary<string, double> ReadSourceUtilMap(string srcPath)
        {
            // 허용 장비만
            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            
            // 1) ExcelDataReader로 시도 (xls/xlsx)
            try
            {
                using (var fs = File.Open(srcPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(fs))
                {
                    var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false
                        }
                    });

                    if (ds.Tables.Count == 0) return map;
                    var t = ds.Tables[0]; // 첫 시트
                    
                    for (int r = 0; r < t.Rows.Count; r++)
                    {
                        var equip = SafeToString(t.Rows[r][SrcColEquip - 1]);
                        if (string.IsNullOrWhiteSpace(equip)) continue;
                        
                        equip = equip.Trim();
                        if (!_allowedEquip.Contains(equip)) continue;
                        
                        var utilObj = t.Rows[r][SrcColUtil - 1];
                        var util = ParseUtilPercent(utilObj);
                        if (!util.HasValue) continue;
                        
                        map[equip] = util.Value; // 0~100
                    }
                }
                return map;
            }
            catch
            {
                // 2) 여기서 실패하면 (HTML형 xls 등) -> fallback 필요
                //    네 환경에서 실제 파일이 HTML형이면 아래 fallback을 살려야 함.
                return ReadSourceUtilMap_HtmlFallback(srcPath);
            }
        }

        // HTML형 xls fallback (폴더가 같이 있을 때)
        private Dictionary<string, double> ReadSourceUtilMap_HtmlFallback(string srcPath)
        {
            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            
            // 예: 표준_20251201.files/sheet001.htm
            var baseName = Path.GetFileNameWithoutExtension(srcPath);
            var dir = Path.Combine(Path.GetDirectoryName(srcPath), baseName + ".files");
            var sheet1 = Path.Combine(dir, "sheet001.htm");
            
            if (!File.Exists(sheet1))
                throw new Exception("소스가 HTML형 xls로 보입니다. 그러나 sheet001.htm를 찾지 못했습니다:\n" + sheet1);
            
            var html = File.ReadAllText(sheet1, Encoding.UTF8);
            
            // 아주 단순 파서(필요 최소): <tr><td>...</td>... 추출
            // A열/ D열만 필요하므로 row별 td를 뽑아 A(0),D(3) 사용
            foreach (var row in ExtractHtmlTableRows(html))
            {
                if (row.Count < 4) continue;
                
                var equip = (row[0] ?? "").Trim();
                if (string.IsNullOrWhiteSpace(equip)) continue;
                if (!_allowedEquip.Contains(equip)) continue;
                
                var util = ParseUtilPercent(row[3]);
                if (!util.HasValue) continue;
                
                map[equip] = util.Value;
            }

            return map;
        }

        private static List<List<string>> ExtractHtmlTableRows(string html)
        {
            var rows = new List<List<string>>();
            // 매우 단순 정규식 기반 (엑셀 HTML 저장형에 잘 맞는 편)
            var trMatches = Regex.Matches(html, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match tr in trMatches)
            {
                var tds = new List<string>();
                var tdMatches = Regex.Matches(tr.Groups[1].Value, @"<t[dh][^>]*>(.*?)</t[dh]>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                foreach (Match td in tdMatches)
                {
                    var text = Regex.Replace(td.Groups[1].Value, "<.*?>", string.Empty);
                    text = System.Net.WebUtility.HtmlDecode(text);
                    tds.Add(text.Trim());
                }
                if (tds.Count > 0) rows.Add(tds);
            }
            return rows;
        }

        // ---------------------------
        // 타겟에 적용
        // ---------------------------
        private sealed class ApplyResult
        {
            public int UpdatedRows;
            public List<string> NotFoundInTarget = new List<string>();
            public List<string> NotFoundInSource = new List<string>();
        }

        private ApplyResult ApplyToTarget(string targetPath, DateTime reportDate, Dictionary<string, double> utilMap)
        {
            var result = new ApplyResult();
            
            using (var wb = new XLWorkbook(targetPath))
            {
                var ws = wb.Worksheet(TargetSheetIndex);
                
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? TargetRowStart - 1;
                if (lastRow < TargetRowStart) lastRow = TargetRowStart;
                
                // 타겟에 존재하는 BA-FC-01~16을 수집해서, 소스에 없는 것도 표시 가능
                var targetEquipSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                
                for (int r = TargetRowStart; r <= lastRow; r++)
                {
                    var dateCell = ws.Cell(r, ColDate);
                    var equipCell = ws.Cell(r, ColEquip);
                    
                    var equip = equipCell.GetString().Trim();
                    if (string.IsNullOrWhiteSpace(equip)) continue;
                    
                    if (!_allowedEquip.Contains(equip)) continue;
                    targetEquipSet.Add(equip);
                    
                    var rowDate = TryReadExcelDate(dateCell);
                    if (!rowDate.HasValue) continue;
                    
                    if (rowDate.Value.Date != reportDate.Date) continue;
                    
                    double util;
                    if (!utilMap.TryGetValue(equip, out util))
                    {
                        // 소스에 없으면 그냥 넘어감
                        continue;
                    }

                    // F열(Util%)에 숫자로 입력
                    ws.Cell(r, ColUtil).Value = util;
                    result.UpdatedRows++;
                }

                // 타겟에서 찾지 못한 장비(날짜 포함) 표시
                foreach (var kv in utilMap)
                {
                    if (!targetEquipSet.Contains(kv.Key))
                        result.NotFoundInTarget.Add($"{kv.Key} (타겟에 장비명 없음)");
                }

                // BA-FC-01~16 중 소스에 없는 것 표시
                foreach (var eq in _allowedEquip.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    if (!utilMap.ContainsKey(eq))
                        result.NotFoundInSource.Add(eq);
                }

                // 저장 (엑셀이 열려있으면 여기서 예외 발생)
                wb.Save();
            }

            return result;
        }

        private static DateTime? TryReadExcelDate(IXLCell cell)
        {
            if (cell == null) return null;
            
            if (cell.DataType == XLDataType.DateTime)
                return cell.GetDateTime();
            
            if (cell.DataType == XLDataType.Number)
            {
                try { return DateTime.FromOADate(cell.GetDouble()); } catch { }
            }
            
            var s = cell.GetString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            
            DateTime dt;
            if (DateTime.TryParse(s, out dt)) return dt;
            
            return null;
        }

        // ---------------------------
        // Preview Table
        // ---------------------------
        private DataTable BuildPreviewTable(Dictionary<string, double> utilMap, ApplyResult result)
        {
            var dt = new DataTable();
            dt.Columns.Add("Equip");
            dt.Columns.Add("Util(%)", typeof(double));
            
            foreach (var kv in utilMap.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
            {
                var row = dt.NewRow();
                row[0] = kv.Key;
                row[1] = kv.Value;
                dt.Rows.Add(row);
            }

            return dt;
        }

        // ---------------------------
        // Parse Utils
        // ---------------------------
        private static string SafeToString(object o)
        {
            if (o == null || o == DBNull.Value) return string.Empty;
            return Convert.ToString(o).Trim();
        }

        // 가동률 파싱: "85.3%", "85.3", 0.853 등 모두 대응
        private static double? ParseUtilPercent(object v)
        {
            if (v == null || v == DBNull.Value) return null;
            
            if (v is double)
            {
                var d = (double)v;
                if (d > 0 && d <= 1.0) return Math.Round(d * 100.0, 1); // 0~1이면 퍼센트로 변환
                return Math.Round(d, 1);
            }

            if (v is float)
            {
                var d = Convert.ToDouble(v);
                if (d > 0 && d <= 1.0) return Math.Round(d * 100.0, 1);
                return Math.Round(d, 1);
            }

            var s = Convert.ToString(v).Trim();
            if (string.IsNullOrEmpty(s)) return null;
            
            s = s.Replace("%", "").Trim();
            double d2;
            
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d2) ||
                double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d2))
            {
                if (d2 > 0 && d2 <= 1.0) return Math.Round(d2 * 100.0, 1);
                return Math.Round(d2, 1);
            }

            return null;
        }

        private void Log(string msg)
        {
            txtLog.AppendText(msg + Environment.NewLine);
        }
    }
}