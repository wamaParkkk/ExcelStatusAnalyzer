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
    public partial class UeMonthlyAutoFillForm : Form
    {
        private Button btnLoadAndApply;
        private DataGridView dgv;
        private Label lblFile;
        private TextBox txtLog;
        private OpenFileDialog ofd;
        
        private Dictionary<string, TargetBlockInfo> _targetMap;
        
        // 고정 타겟 경로
        private const string TargetPath = @"C:\Users\156607\Amkor_Project\Document\장비 가동률 데이터\UE 월별 실적.xlsx";
        
        // 소스 컬럼
        private const int SrcColEquip = 1; // A
        private const int SrcColUtil = 4;  // D
        private const int SrcColRun = 5;   // E
        
        private readonly HashSet<string> _allowedEquip = new HashSet<string>(new[]
        {
           "RLTC-01","RLTC-02","RLTC-03","RLTC-04","RLTC-05","RLTC-06",
           "TC02","TC03","TC04","TC05","TC06","TC07",
           "KTCB-01","KTCB-02","KTCB-03","KTCB-04","KTCB-05","KTCB-06"
        }, StringComparer.OrdinalIgnoreCase);
        
        public UeMonthlyAutoFillForm()
        {
            try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); } catch { }
            
            BuildUi();
            InitTargetMap();
        }

        private void BuildUi()
        {
            Text = "UE 월별 실적 자동 입력";
            Width = 1250;
            Height = 800;
            
            btnLoadAndApply = new Button
            {
                Left = 15,
                Top = 15,
                Width = 280,
                Height = 34,
                Text = "Source 엑셀 불러오기 + Target 자동 입력"
            };
            btnLoadAndApply.Click += BtnLoadAndApply_Click;
            
            lblFile = new Label
            {
                Left = 310,
                Top = 23,
                Width = 900,
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
                Title = "UE 실적 Source 파일 선택"
            };
            
            Controls.Add(btnLoadAndApply);
            Controls.Add(lblFile);
            Controls.Add(dgv);
            Controls.Add(txtLog);
        }

        private sealed class TargetBlockInfo
        {
            public int BaseRow;
            public int BlockIndex; // 1-based
        }

        private void InitTargetMap()
        {
            _targetMap = new Dictionary<string, TargetBlockInfo>(StringComparer.OrdinalIgnoreCase);
            
            // RLTC-01 ~ RLTC-08 : row 2 시작, 5열 간격
            var rltc = new[] { "RLTC-01", "RLTC-02", "RLTC-03", "RLTC-04", "RLTC-05", "RLTC-06", "RLTC-07", "RLTC-08" };
            for (int i = 0; i < rltc.Length; i++)
            {
                _targetMap[rltc[i]] = new TargetBlockInfo
                {
                    BaseRow = 2,
                    BlockIndex = i + 1
                };
            }
            
            // TC02 ~ TC07 : row 36 시작
            var tc = new[] { "TC02", "TC03", "TC04", "TC05", "TC06", "TC07" };
            for (int i = 0; i < tc.Length; i++)
            {
                _targetMap[tc[i]] = new TargetBlockInfo
                {
                    BaseRow = 36,
                    BlockIndex = i + 1
                };
            }
            
            // KTCB-01 ~ KTCB-06 : row 70 시작
            var ktcb = new[] { "KTCB-01", "KTCB-02", "KTCB-03", "KTCB-04", "KTCB-05", "KTCB-06" };
            for (int i = 0; i < ktcb.Length; i++)
            {
                _targetMap[ktcb[i]] = new TargetBlockInfo
                {
                    BaseRow = 70,
                    BlockIndex = i + 1
                };
            }
        }

        private void BtnLoadAndApply_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;
            
            txtLog.Clear();
            
            try
            {
                var srcPath = ofd.FileName;
                lblFile.Text = "파일: " + Path.GetFileName(srcPath);
                
                // 파일명 날짜 추출
                var reportDate = ExtractDateFromFileName(srcPath);
                if (!reportDate.HasValue)
                    throw new Exception("소스 파일명에서 yyyymmdd 또는 yyyy-MM-dd 날짜를 찾지 못했습니다.");
                
                Log($"[소스 날짜] {reportDate.Value:yyyy-MM-dd}");
                
                // 소스 읽기
                var srcMap = ReadSourceDataMap(srcPath);
                if (srcMap.Count == 0)
                    throw new Exception("소스 파일에서 장비 / Util / 가동시간 데이터를 찾지 못했습니다.");
                
                Log($"[소스] 추출 장비 수: {srcMap.Count}");
                
                if (!File.Exists(TargetPath))
                    throw new Exception("타겟 파일이 존재하지 않습니다: " + TargetPath);
                
                var result = ApplyToTarget(TargetPath, reportDate.Value.Date, srcMap);
                
                dgv.DataSource = BuildPreviewTable(srcMap);
                
                Log($"[적용 완료] 업데이트 행 수: {result.UpdatedRows}");
                
                if (result.NotFoundInTarget.Count > 0)
                {
                    Log("[타겟에 없는 장비/날짜]");
                    foreach (var s in result.NotFoundInTarget)
                        Log(" - " + s);
                }
                
                if (result.NotFoundInSource.Count > 0)
                {
                    Log("[소스에 없는 장비]");
                    foreach (var s in result.NotFoundInSource)
                        Log(" - " + s);
                }
                
                MessageBox.Show("완료! 타겟 엑셀에 Util / 가동시간을 입력했습니다.",
                    "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처리 실패: " + ex.Message,
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("[ERROR] " + ex);
            }
        }

        // ---------------------------
        // 날짜 추출
        // ---------------------------
        private static DateTime? ExtractDateFromFileName(string path)
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

        // ---------------------------
        // Source 읽기
        // ---------------------------
        private sealed class SourceItem
        {
            public double? Util;
            public double? RunTime;
        }

        private Dictionary<string, SourceItem> ReadSourceDataMap(string srcPath)
        {
            var map = new Dictionary<string, SourceItem>(StringComparer.OrdinalIgnoreCase);
            
            // 1) ExcelDataReader로 시도
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
                    var t = ds.Tables[0];
                    
                    for (int r = 0; r < t.Rows.Count; r++)
                    {
                        var equipRaw = SafeToString(t.Rows[r][SrcColEquip - 1]);
                        if (string.IsNullOrWhiteSpace(equipRaw)) continue;
                        
                        var equip = NormalizeEquipName(equipRaw);
                        if (string.IsNullOrWhiteSpace(equip)) continue;
                        if (!_allowedEquip.Contains(equip)) continue;
                        
                        var util = ParseNumber(t.Rows[r][SrcColUtil - 1]);
                        var run = ParseNumber(t.Rows[r][SrcColRun - 1]);
                        
                        if (!util.HasValue && !run.HasValue) continue;
                        
                        map[equip] = new SourceItem
                        {
                            Util = util,
                            RunTime = run
                        };
                    }
                }

                return map;
            }
            catch
            {
                // 2) HTML형 xls fallback
                return ReadSourceDataMap_HtmlFallback(srcPath);
            }
        }

        private Dictionary<string, SourceItem> ReadSourceDataMap_HtmlFallback(string srcPath)
        {
            var map = new Dictionary<string, SourceItem>(StringComparer.OrdinalIgnoreCase);
            
            var baseName = Path.GetFileNameWithoutExtension(srcPath);
            var dir = Path.Combine(Path.GetDirectoryName(srcPath), baseName + ".files");
            var sheet1 = Path.Combine(dir, "sheet001.htm");
            
            if (!File.Exists(sheet1))
                throw new Exception("소스가 HTML형 xls로 보입니다. 그러나 sheet001.htm를 찾지 못했습니다:\n" + sheet1);
            
            string html;
            try { html = File.ReadAllText(sheet1, Encoding.UTF8); }
            catch { html = File.ReadAllText(sheet1, Encoding.Default); }
            
            foreach (var row in ExtractHtmlTableRows(html))
            {
                if (row.Count < 5) continue;
                
                var equipRaw = (row[0] ?? "").Trim(); // A
                if (string.IsNullOrWhiteSpace(equipRaw)) continue;
                
                var equip = NormalizeEquipName(equipRaw);
                if (string.IsNullOrWhiteSpace(equip)) continue;
                if (!_allowedEquip.Contains(equip)) continue;
                
                var util = ParseNumber(row[3]); // D
                var run = ParseNumber(row[4]);  // E
                
                if (!util.HasValue && !run.HasValue) continue;
                
                map[equip] = new SourceItem
                {
                    Util = util,
                    RunTime = run
                };
            }

            return map;
        }

        private static string NormalizeEquipName(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            
            var s = raw.Trim();
            
            var parts = s.Split(new[] { ',', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            s = (parts.Length > 0) ? parts[0].Trim() : s;
            
            var tokens = s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length > 0) s = tokens[0].Trim();
            
            return s;
        }

        // ---------------------------
        // Target 적용
        // ---------------------------
        private sealed class ApplyResult
        {
            public int UpdatedRows;
            public List<string> NotFoundInTarget = new List<string>();
            public List<string> NotFoundInSource = new List<string>();
        }

        private ApplyResult ApplyToTarget(string targetPath, DateTime reportDate, Dictionary<string, SourceItem> srcMap)
        {
            var result = new ApplyResult();
            
            using (var wb = new XLWorkbook(targetPath))
            {
                string targetSheetName = reportDate.Month.ToString("00") + "월";
                if (!wb.Worksheets.Any(x => x.Name == targetSheetName))
                    throw new Exception("타겟 파일에서 시트를 찾지 못했습니다: " + targetSheetName);
                
                var ws = wb.Worksheet(targetSheetName);
                
                foreach (var kv in srcMap)
                {
                    var equip = kv.Key;
                    var item = kv.Value;
                    
                    TargetBlockInfo block;
                    if (!_targetMap.TryGetValue(equip, out block))
                    {
                        result.NotFoundInTarget.Add($"{equip} (타겟 블록 매핑 없음)");
                        continue;
                    }
                    
                    int dateCol = GetDateCol(block.BlockIndex);
                    int equipCol = GetEquipCol(block.BlockIndex);
                    int utilCol = GetUtilCol(block.BlockIndex);
                    int runCol = GetRunCol(block.BlockIndex);
                    
                    int lastRow = ws.LastRowUsed()?.RowNumber() ?? block.BaseRow;
                    
                    bool wrote = false;
                    
                    for (int r = block.BaseRow; r <= lastRow; r++)
                    {
                        var dt = TryReadExcelDate(ws.Cell(r, dateCol));
                        if (!dt.HasValue) continue;
                        if (dt.Value.Date != reportDate.Date) continue;
                        
                        var curEquip = ws.Cell(r, equipCol).GetString().Trim();
                        if (string.IsNullOrWhiteSpace(curEquip))
                            ws.Cell(r, equipCol).Value = equip;
                        
                        if (item.Util.HasValue)
                        {
                            ws.Cell(r, utilCol).Value = item.Util.Value;
                            ws.Cell(r, utilCol).Style.NumberFormat.Format = "0.00";
                        }
                        
                        if (item.RunTime.HasValue)
                        {
                            ws.Cell(r, runCol).Value = item.RunTime.Value;
                            ws.Cell(r, runCol).Style.NumberFormat.Format = "0.00";
                        }
                        
                        result.UpdatedRows++;
                        wrote = true;
                        break;
                    }

                    if (!wrote)
                        result.NotFoundInTarget.Add($"{equip} ({targetSheetName} 시트에서 날짜 {reportDate:yyyy-MM-dd} 행을 못 찾음)");
                }
                foreach (var eq in _allowedEquip.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                {
                    if (!srcMap.ContainsKey(eq))
                        result.NotFoundInSource.Add(eq);
                }

                wb.Save();
            }

            return result;
        }

        // 블록 계산: 5열 간격
        private static int GetDateCol(int blockIndex) => 1 + (blockIndex - 1) * 5; // A, F, K ...
        private static int GetEquipCol(int blockIndex) => 2 + (blockIndex - 1) * 5; // B, G, L ...
        private static int GetUtilCol(int blockIndex) => 3 + (blockIndex - 1) * 5; // C, H, M ...
        private static int GetRunCol(int blockIndex) => 4 + (blockIndex - 1) * 5;  // D, I, N ...
        
        private static DateTime? TryReadExcelDate(IXLCell cell)
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
            if (DateTime.TryParse(s, out dt)) return dt;
            
            return null;
        }

        // ---------------------------
        // Preview
        // ---------------------------
        private DataTable BuildPreviewTable(Dictionary<string, SourceItem> srcMap)
        {
            var dt = new DataTable();
            dt.Columns.Add("Equip");
            dt.Columns.Add("Util", typeof(double));
            dt.Columns.Add("RunTime", typeof(double));
            
            foreach (var kv in srcMap.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                var row = dt.NewRow();
                row["Equip"] = kv.Key;
                row["Util"] = kv.Value.Util.HasValue ? (object)kv.Value.Util.Value : DBNull.Value;
                row["RunTime"] = kv.Value.RunTime.HasValue ? (object)kv.Value.RunTime.Value : DBNull.Value;
                dt.Rows.Add(row);
            }

            return dt;
        }

        // ---------------------------
        // Utils
        // ---------------------------
        private static string SafeToString(object o)
        {
            if (o == null || o == DBNull.Value) return string.Empty;
            return Convert.ToString(o).Trim();
        }

        private static double? ParseNumber(object v)
        {
            if (v == null || v == DBNull.Value) return null;
            
            if (v is double) return (double)v;
            if (v is float) return Convert.ToDouble(v);
            if (v is int) return Convert.ToDouble(v);
            if (v is long) return Convert.ToDouble(v);
            
            var s = Convert.ToString(v);
            if (string.IsNullOrWhiteSpace(s)) return null;
            
            s = s.Trim().Replace("%", "");
            
            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d) ||
                double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                return d;
            
            return null;
        }

        private static List<List<string>> ExtractHtmlTableRows(string html)
        {
            var rows = new List<List<string>>();
            
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

        private void Log(string msg)
        {
            txtLog.AppendText(msg + Environment.NewLine);
        }
    }
}