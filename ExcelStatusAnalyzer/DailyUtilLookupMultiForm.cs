// Source Excel format: A=EquipName, D=Util(%)
// Supports .xlsx/.xls via ExcelDataReader, and HTML-style .xls fallback (.files\sheet001.htm)

using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class DailyUtilLookupMultiForm : Form
    {
        private Button btnLoad;
        private Button btnAddRow;
        private Label lblFile;
        private TextBox txtLog;
        private OpenFileDialog ofd;
        
        private Panel pnlRows;
        private int _rowSeq = 0;
        
        // 엑셀에서 읽은 (장비명 -> Util)
        private Dictionary<string, double> _utilByEquip = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        
        // 소스 컬럼 (요구사항)
        private const int SrcColEquip = 1; // A
        private const int SrcColUtil = 4;  // D
        
        public DailyUtilLookupMultiForm()
        {
            // (일부 환경에서 필요) ExcelDataReader 코드페이지 등록
            try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); } catch { }
            
            BuildUi();
            AddEquipRow(); // 기본 1줄
        }

        private void BuildUi()
        {
            Text = "가동률 조회 (멀티 장비) - Source Excel A=장비명, D=가동률";
            Width = 1100;
            Height = 1000;
            
            btnLoad = new Button
            {
                Left = 15,
                Top = 15,
                Width = 200,
                Height = 34,
                Text = "엑셀 불러오기"
            };
            btnLoad.Click += BtnLoad_Click;
            
            btnAddRow = new Button
            {
                Left = 225,
                Top = 15,
                Width = 180,
                Height = 34,
                Text = "장비 입력칸 추가"
            };
            btnAddRow.Click += (s, e) => AddEquipRow();
            
            lblFile = new Label
            {
                Left = 420,
                Top = 23,
                Width = 650,
                Text = "파일: (없음)"
            };
            
            pnlRows = new Panel
            {
                Left = 15,
                Top = 60,
                Width = ClientSize.Width - 30,
                Height = 830,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            
            txtLog = new TextBox
            {
                Left = 15,
                Top = 895,
                Width = ClientSize.Width - 30,
                Height = ClientSize.Height - 900,
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

            Controls.Add(btnLoad);
            Controls.Add(btnAddRow);
            Controls.Add(lblFile);
            Controls.Add(pnlRows);
            Controls.Add(txtLog);
        }

        // ---------------------------
        // UI Row 생성/조회/복사
        // ---------------------------
        private void AddEquipRow()
        {
            _rowSeq++;
            
            int top = (_rowSeq - 1) * 38 + 10;
            
            var txtEquip = new TextBox
            {
                Left = 10,
                Top = top,
                Width = 240,
                Tag = _rowSeq
            };
            
            var btnQuery = new Button
            {
                Left = 260,
                Top = top - 1,
                Width = 70,
                Height = 26,
                Text = "조회",
                Tag = _rowSeq
            };
            
            var lbl = new Label
            {
                Left = 340,
                Top = top + 4,
                Width = 100,
                Text = "Util: -",
                Tag = null
            };
            
            var btnCopy = new Button
            {
                Left = 450,
                Top = top - 1,
                Width = 80,
                Height = 26,
                Text = "복사",
                Tag = lbl
            };
            
            var btnRemove = new Button
            {
                Left = 540,
                Top = top - 1,
                Width = 70,
                Height = 26,
                Text = "삭제"
            };
            
            // 조회 버튼
            btnQuery.Click += (s, e) =>
            {
                QueryOneRow(txtEquip.Text, lbl);
            };
            
            // Enter로 조회
            txtEquip.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    QueryOneRow(txtEquip.Text, lbl);
                }
            };
            
            // 복사 버튼 (개별)
            btnCopy.Click += (s, e) =>
            {
                var targetLbl = (Label)((Button)s).Tag;
                if (targetLbl.Tag == null)
                {
                    MessageBox.Show("복사할 Util 값이 없습니다. 먼저 조회하세요.", "안내",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var util = (double)targetLbl.Tag;
                Clipboard.Clear();
                Clipboard.SetText(util.ToString("0.00", CultureInfo.InvariantCulture));
            };

            // 삭제 버튼
            btnRemove.Click += (s, e) =>
            {
                pnlRows.Controls.Remove(txtEquip);
                pnlRows.Controls.Remove(btnQuery);
                pnlRows.Controls.Remove(lbl);
                pnlRows.Controls.Remove(btnCopy);
                pnlRows.Controls.Remove(btnRemove);
                pnlRows.Refresh();
            };

            pnlRows.Controls.Add(txtEquip);
            pnlRows.Controls.Add(btnQuery);
            pnlRows.Controls.Add(lbl);
            pnlRows.Controls.Add(btnCopy);
            pnlRows.Controls.Add(btnRemove);
        }

        private void QueryOneRow(string equipInput, Label lbl)
        {
            if (_utilByEquip == null || _utilByEquip.Count == 0)
            {
                lbl.Text = "Util: - (엑셀을 먼저 불러오세요)";
                lbl.Tag = null;
                return;
            }

            var equipKey = NormalizeFirstEquipName(equipInput);
            
            if (string.IsNullOrWhiteSpace(equipKey))
            {
                lbl.Text = "Util: - (장비명 입력 필요)";
                lbl.Tag = null;
                return;
            }
            
            double util;
            if (_utilByEquip.TryGetValue(equipKey, out util))
            {
                // 표시만 2자리 (값은 원본 유지)
                lbl.Text = "Util: " + util.ToString("0.00", CultureInfo.InvariantCulture) + " %";
                lbl.Tag = util; // 복사용 저장
            }
            else
            {
                lbl.Text = "Util: - (엑셀에서 미발견)";
                lbl.Tag = null;
            }
        }

        // ---------------------------
        // Load Excel
        // ---------------------------
        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;
            
            txtLog.Clear();
            
            try
            {
                var srcPath = ofd.FileName;
                lblFile.Text = "파일: " + Path.GetFileName(srcPath);
                
                var map = ReadSourceUtilMap(srcPath);
                _utilByEquip = map;
                
                Log("[로드 성공] 장비 수: " + _utilByEquip.Count);
                
                // 디버그: 실제로 어떤 키로 들어갔는지 샘플 출력(최대 20개)
                var sample = _utilByEquip.Keys.Take(20).ToArray();
                Log("샘플키(20): " + (sample.Length == 0 ? "(없음)" : string.Join(", ", sample)));
                
                MessageBox.Show("엑셀 로드 완료!\n장비명을 입력하고 [조회]를 누르세요.",
                    "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("로딩 실패: " + ex.Message, "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("[ERROR] " + ex);
            }
        }

        // ---------------------------
        // Source Read (ExcelDataReader + HTML fallback)
        // ---------------------------
        private Dictionary<string, double> ReadSourceUtilMap(string srcPath)
        {
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
                        var equipRaw = SafeToString(t.Rows[r][SrcColEquip - 1]);
                        if (string.IsNullOrWhiteSpace(equipRaw)) continue;
                        
                        // 장비명이 "FC04, FC04" 처럼 2개면 첫번째만 사용
                        var equip = NormalizeFirstEquipName(equipRaw);
                        if (string.IsNullOrWhiteSpace(equip)) continue;
                        
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
                // 2) 실패하면 (HTML형 xls 등) -> fallback
                return ReadSourceUtilMap_HtmlFallback(srcPath);
            }
        }

        // HTML형 xls fallback: 예) 표준_20251201.xls -> 표준_20251201.files\sheet001.htm
        private Dictionary<string, double> ReadSourceUtilMap_HtmlFallback(string srcPath)
        {
            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            
            var baseName = Path.GetFileNameWithoutExtension(srcPath);
            var dir = Path.Combine(Path.GetDirectoryName(srcPath), baseName + ".files");
            var sheet1 = Path.Combine(dir, "sheet001.htm");
            
            if (!File.Exists(sheet1))
                throw new Exception("소스가 HTML형 xls로 보입니다. 그러나 sheet001.htm를 찾지 못했습니다:\n" + sheet1);
            
            // 인코딩은 파일에 따라 다를 수 있어 UTF-8 우선, 실패 시 Default
            string html;
            try { html = File.ReadAllText(sheet1, Encoding.UTF8); }
            catch { html = File.ReadAllText(sheet1, Encoding.Default); }
            
            foreach (var row in ExtractHtmlTableRows(html))
            {
                if (row.Count < 4) continue;
                
                var equipRaw = (row[0] ?? "").Trim();
                if (string.IsNullOrWhiteSpace(equipRaw)) continue;
                
                var equip = NormalizeFirstEquipName(equipRaw);
                if (string.IsNullOrWhiteSpace(equip)) continue;
                
                var util = ParseUtilPercent(row[3]);
                if (!util.HasValue) continue;
                
                // 헤더 라인 스킵(장비/equip/name)
                var low = equip.ToLowerInvariant();
                if (low.Contains("장비") || low.Contains("equip") || low.Contains("name"))
                    continue;
                
                map[equip] = util.Value;
            }

            return map;
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
                    tds.Add((text ?? "").Trim());
                }

                if (tds.Count > 0) rows.Add(tds);
            }

            return rows;
        }

        // ---------------------------
        // Helpers
        // ---------------------------
        private static string SafeToString(object o)
        {
            if (o == null || o == DBNull.Value) return string.Empty;
            return Convert.ToString(o).Trim();
        }

        // 장비명이 "FC04, FC04" / "FC04 / FC04" / "FC04 FC04" 형태면 첫번째만 사용
        private static string NormalizeFirstEquipName(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            
            var s = raw.Trim();
            
            // 콤마/슬래시/개행/탭 등 구분자 기준 첫 덩어리
            var parts = s.Split(new[] { ',', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            s = (parts.Length > 0) ? parts[0].Trim() : s;
            
            // 공백으로 두 개가 붙는 경우 첫 토큰
            var tokens = s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length > 0) s = tokens[0].Trim();
            
            return s;
        }

        // 가동률 파싱: "85.3%", "85.3", 0.853 등 모두 대응
        // 반올림/정수화 없이 "값 그대로" 반환 (0~1이면 *100 변환만)
        private static double? ParseUtilPercent(object v)
        {
            if (v == null || v == DBNull.Value) return null;
            
            if (v is double)
            {
                var d = (double)v;
                if (d > 0 && d <= 1.0) return d * 100.0;
                return d;
            }
            
            if (v is float)
            {
                var d = Convert.ToDouble(v);
                if (d > 0 && d <= 1.0) return d * 100.0;
                return d;
            }
            
            var s = Convert.ToString(v);
            if (s == null) return null;
            
            s = s.Trim();
            if (string.IsNullOrEmpty(s)) return null;
            
            s = s.Replace("%", "").Trim();
            
            double d2;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d2) ||
                double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d2))
            {
                if (d2 > 0 && d2 <= 1.0) return d2 * 100.0;
                return d2;
            }

            return null;
        }

        private void Log(string msg)
        {
            txtLog.AppendText(msg + Environment.NewLine);
        }
    }
}