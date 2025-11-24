using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace ExcelStatusAnalyzer
{
    public partial class ShiftWebSummaryForm : Form
    {
        // UI
        private MonthCalendar cal;
        private CheckBox chkDay, chkSwing, chkNight;
        private ComboBox cboEqp;
        private TextBox txtBaseUrl;
        private Button btnFetch, btnCopy;
        private Label lblInfo;
        private DataGridView dgv;
        
        // 합계 열 내부 이름
        private const string TotalColName = "__TOTAL__";

        public ShiftWebSummaryForm()
        {
            BuildUi();
        }
        
        private void BuildUi()
        {
            Text = "교대별 장비 상태 집계 (CDT Detail)";
            Width = 1200;
            Height = 800;
            
            cal = new MonthCalendar
            {
                Left = 15,
                Top = 15,
                MaxSelectionCount = 1
            };

            chkDay = new CheckBox { Left = 260, Top = 15, Width = 140, Text = "Day (06:00~13:59)", Checked = true };
            chkSwing = new CheckBox { Left = 260, Top = 40, Width = 140, Text = "Swing (14:00~21:59)", Checked = true };
            chkNight = new CheckBox { Left = 260, Top = 65, Width = 170, Text = "Night (22:00~05:59)", Checked = true };
            
            cboEqp = new ComboBox
            {
                Left = 440,
                Top = 15,
                Width = 250,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // 표시: "RLTC-01 | 2025-1100962"  값: "RLTC-01"
            // 표시: "RLTC-01 | 2025-1100962", 값: Code/LineNo 모두 저장
            cboEqp.Items.Add(new ComboItem("RLTC-01 | 2025-1100962", "RLTC-01", "2025-1100962"));
            cboEqp.Items.Add(new ComboItem("RLTC-02 | 2025-1100166", "RLTC-02", "2025-1100166"));
            cboEqp.Items.Add(new ComboItem("RLTC-03 | 2025-1100360", "RLTC-03", "2025-1100360"));
            cboEqp.Items.Add(new ComboItem("RLTC-04 | 2025-1100603", "RLTC-04", "2025-1100603"));
            cboEqp.Items.Add(new ComboItem("RLTC-05 | 2025-1100653", "RLTC-05", "2025-1100653"));

            cboEqp.Items.Add(new ComboItem("TC01 | 2018-1100041", "TC01", "2018-1100041"));
            cboEqp.Items.Add(new ComboItem("TC02 | 2021-1101010", "TC02", "2018-1101010"));
            cboEqp.Items.Add(new ComboItem("TC03 | 2021-1101565", "TC03", "2018-1101565"));
            cboEqp.Items.Add(new ComboItem("TC04 | 2023-1100765", "TC04", "2018-1100765"));
            cboEqp.Items.Add(new ComboItem("TC05 | 2024-1100236", "TC05", "2018-1100236"));
            cboEqp.Items.Add(new ComboItem("TC06 | 2024-1100237", "TC06", "2018-1100237"));
            cboEqp.Items.Add(new ComboItem("TC07 | 2024-1100528", "TC07", "2018-1100528"));

            if (cboEqp.Items.Count > 0) cboEqp.SelectedIndex = 0;
            
            txtBaseUrl = new TextBox
            {
                Left = 440,
                Top = 48,
                Width = 560,
                Text = "http://cim_service.amkor.co.kr:8080/ysj/get_cdt_report/get_cdt_detail_report?connector_name_array=K5FCBGATCBCIM01&", // ← date_from/date_to는 아래에서 자동으로 붙입니다.
            };
            
            btnFetch = new Button { Left = 1010, Top = 15, Width = 160, Height = 32, Text = "조회" };
            btnFetch.Click += async (s, e) => await FetchAndShowAsync();
            btnCopy = new Button { Left = 1010, Top = 52, Width = 160, Height = 32, Text = "결과 복사(헤더 포함)" };
            btnCopy.Click += (s, e) => CopyGridAll();
            
            lblInfo = new Label
            {
                Left = 15,
                Top = 190,
                Width = ClientSize.Width - 30,
                Text = "장비/날짜/교대 선택 후 [조회]. Night는 22:00~다음날 05:59 범위까지 포함.",
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            
            dgv = new DataGridView
            {
                Dock = DockStyle.Bottom,
                Top = 220,
                Height = ClientSize.Height - 230,
                ReadOnly = true,
                EditMode = DataGridViewEditMode.EditProgrammatically,
                ScrollBars = ScrollBars.Both,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                AllowUserToAddRows = false
            };
            
            dgv.BringToFront();
            Controls.Add(cal);
            Controls.Add(chkDay);
            Controls.Add(chkSwing);
            Controls.Add(chkNight);
            Controls.Add(cboEqp);
            Controls.Add(txtBaseUrl);
            Controls.Add(btnFetch);
            Controls.Add(btnCopy);
            Controls.Add(lblInfo);
            Controls.Add(dgv);
        }
        // 조회 실행
        
        private async Task FetchAndShowAsync()
        {
            try
            {
                // 날짜 & 교대
                DateTime baseDate = cal.SelectionStart.Date;
                bool incDay = chkDay.Checked, incSwing = chkSwing.Checked, incNight = chkNight.Checked;
                if (!incDay && !incSwing && !incNight) { incDay = incSwing = incNight = true; }
                                                                                                
                // 장비 코드 (RLTC-0x)
                var sel = GetSelectedEqp();
                if (sel == null)
                {
                    MessageBox.Show("장비를 선택하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                string eqpCode = sel.Code;     // RLTC-01
                string eqpLine = sel.LineNo;   // 2025-...
                
                // 시간대 범위 만들기 (여러 개 선택 시 다중 호출)
                var ranges = BuildShiftRanges(baseDate, incDay, incSwing, incNight);
                
                // 누적 합계
                var sumsTime = InitTimeSums();
                var sumsCount = InitCountSums();
                string baseUrl = (txtBaseUrl.Text ?? "").Trim();
                if (string.IsNullOrEmpty(baseUrl))
                {
                    MessageBox.Show("웹 서비스 URL을 입력하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                using (var http = new HttpClient())
                {
                    http.Timeout = TimeSpan.FromSeconds(30);
                    foreach (var rg in ranges)
                    {
                        string url = BuildUrl(baseUrl, rg.Item1, rg.Item2);
                        string json = await http.GetStringAsync(url);
                        IEnumerable<JToken> records = ExtractRecords(json);
                        foreach (var rec in records)
                        {
                            if (!IsMatchingEquipment(rec, eqpCode, eqpLine)) continue;
                            // 시간(초) 합계
                            AccumulateTime(rec, sumsTime);
                            // 카운트 합계
                            AccumulateCount(rec, sumsCount);
                        }
                    }
                }

                // 결과 테이블 구성
                var dt = BuildResultTable(sumsTime, sumsCount);
                dgv.DataSource = dt;
                
                // 숫자 정렬/폭
                for (int c = 1; c < dgv.Columns.Count; c++)
                    dgv.Columns[c].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                
                // 정보 라벨
                var shiftNames = new List<string>();
                if (incDay) shiftNames.Add("Day");
                if (incSwing) shiftNames.Add("Swing");
                if (incNight) shiftNames.Add("Night");
                lblInfo.Text = $"장비={sel.Code}, 기준일={baseDate:yyyy-MM-dd}, 교대=[{string.Join(",", shiftNames)}], 호출 {ranges.Count}회 완료.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("조회 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // URL 만들기 (date_from/date_to 추가)
        private string BuildUrl(string baseUrl, DateTime from, DateTime to)
        {
            string df = from.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            string dt = to.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            // baseUrl에 이미 ?가 있으면 & 이어붙임, 없으면 ?부터
            var sb = new StringBuilder(baseUrl);
            if (baseUrl.IndexOf('?') >= 0) sb.Append("&"); else sb.Append("?");
            sb.Append("date_from=").Append(df).Append("&date_to=").Append(dt);
            return sb.ToString();
        }

        // 선택한 날짜 day 기준으로 교대별 범위 생성 + '붙는/겹치는' 범위는 병합
        private List<Tuple<DateTime, DateTime>> BuildShiftRanges(DateTime day, bool incDay, bool incSwing, bool incNight)
        {
            var raw = new List<Tuple<DateTime, DateTime>>();
            
            // Day:   day 06:00:00 ~ day 13:59:59
            if (incDay)
            {
                var from = day.AddHours(6);
                var to = day.AddHours(13).AddMinutes(59).AddSeconds(59);
                raw.Add(Tuple.Create(from, to));
            }
            
            // Swing: day 14:00:00 ~ day 21:59:59
            if (incSwing)
            {
                var from = day.AddHours(14);
                var to = day.AddHours(21).AddMinutes(59).AddSeconds(59);
                raw.Add(Tuple.Create(from, to));
            }
            
            // Night: day 22:00:00 ~ day 23:59:59  +  (day+1) 00:00:00 ~ (day+1) 05:59:59
            //  - Night 단독이어도 '전날 귀속'이 아니라 두 구간 모두 포함 (예: 8/20 선택 ⇒ 8/20 22~23:59 + 8/21 00~05:59)
            if (incNight)
            {
                var n1_from = day.AddHours(22);
                var n1_to = day.AddDays(0).AddHours(23).AddMinutes(59).AddSeconds(59);
                raw.Add(Tuple.Create(n1_from, n1_to));
                var n2_from = day.AddDays(1).Date;
                var n2_to = day.AddDays(1).AddHours(5).AddMinutes(59).AddSeconds(59);
                raw.Add(Tuple.Create(n2_from, n2_to));
            }
            
            // 붙는/겹치는 구간 병합(예: Day+Swing, Day+Swing+Night → 06:00 ~ 익일 05:59)
            return CoalesceRanges(raw);
        }

        // JSON 파싱: 배열이거나 { data:[...] } 등 유연 처리
        private IEnumerable<JToken> ExtractRecords(string json)
        {
            var root = JToken.Parse(json);
            if (root is JArray) return (JArray)root;
            
            var data = root["data"];
            if (data is JArray) return (JArray)data;
            
            var rows = root["rows"];
            if (rows is JArray) return (JArray)rows;
            
            return root.Children();
        }

        // 장비 매칭: 여러 필드 후보에서 RLTC-xx 포함 여부 검사
        private bool IsMatchingEquipment(JToken rec, string eqpCode, string eqpLine)
        {
            // 우선순위: 정확히 일치, 다음으로 포함(부분일치)
            var idToken = rec["equipId"];
            var lnToken = rec["equipLineNo"];
            string id = idToken != null ? idToken.ToString().Trim() : "";
            string ln = lnToken != null ? lnToken.ToString().Trim() : "";
            
            if (!string.IsNullOrEmpty(eqpCode))
            {
                if (id.Equals(eqpCode, StringComparison.OrdinalIgnoreCase)) return true;
                if (ln.Equals(eqpCode, StringComparison.OrdinalIgnoreCase)) return true;
                if (id.IndexOf(eqpCode, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (ln.IndexOf(eqpCode, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            
            if (!string.IsNullOrEmpty(eqpLine))
            {
                if (id.Equals(eqpLine, StringComparison.OrdinalIgnoreCase)) return true;
                if (ln.Equals(eqpLine, StringComparison.OrdinalIgnoreCase)) return true;
                if (id.IndexOf(eqpLine, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (ln.IndexOf(eqpLine, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }

            return false;
        }

        // 시간 구간 병합: 시작시간으로 정렬 후, 인접(이전 끝+1초 이상 겹치거나 닿는)하면 합침
        private List<Tuple<DateTime, DateTime>> CoalesceRanges(List<Tuple<DateTime, DateTime>> input)
        {
            var result = new List<Tuple<DateTime, DateTime>>();
            if (input == null || input.Count == 0) return result;
            input.Sort((a, b) => a.Item1.CompareTo(b.Item1));
            DateTime curStart = input[0].Item1;
            DateTime curEnd = input[0].Item2;
            for (int i = 1; i < input.Count; i++)
            {
                var s = input[i].Item1;
                var e = input[i].Item2;
                // 겹치거나 '연속'(curEnd 바로 다음 시각과 닿음)이면 병합
                if (s <= curEnd.AddSeconds(1))
                {
                    if (e > curEnd) curEnd = e;
                }
                else
                {
                    result.Add(Tuple.Create(curStart, curEnd));
                    curStart = s; curEnd = e;
                }
            }
            result.Add(Tuple.Create(curStart, curEnd));
            return result;
        }
        
        private string[] TimeKeys()
        {
            return new[]
            {
               "runTime","activeRunTime","dummyRunTime","activeDummyRunTime",
               "waitingTime","activeWaitingTime","idleTime","troubleTime",
               "dummyTroubleTime","setupTime","commDownTime","lotDownTime"
           };
        }

        private string[] CountKeys()
        {
            return new[]
            {
               "runCount","dummyRunCount","waitingCount","idleCount",
               "troubleCount","dummyTroubleCount","setupCount","commDownCount","lotDownCount"
           };
        }

        private Dictionary<string, long> InitTimeSums()
        {
            var d = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            var arr = TimeKeys();
            for (int i = 0; i < arr.Length; i++) d[arr[i]] = 0L;
            return d;
        }

        private Dictionary<string, long> InitCountSums()
        {
            var d = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            var arr = CountKeys();
            for (int i = 0; i < arr.Length; i++) d[arr[i]] = 0L;
            return d;
        }

        private void AccumulateTime(JToken rec, Dictionary<string, long> sums)
        {
            var keys = TimeKeys();
            for (int i = 0; i < keys.Length; i++)
            {
                var t = rec[keys[i]];
                if (t != null)
                {
                    long v;
                    if (long.TryParse(t.ToString(), out v))
                    {
                        long cur;
                        if (!sums.TryGetValue(keys[i], out cur)) cur = 0;
                        sums[keys[i]] = cur + v;
                    }
                }
            }
        }

        private void AccumulateCount(JToken rec, Dictionary<string, long> sums)
        {
            var keys = CountKeys();
            for (int i = 0; i < keys.Length; i++)
            {
                var t = rec[keys[i]];
                if (t != null)
                {
                    long v;
                    if (long.TryParse(t.ToString(), out v))
                    {
                        long cur;
                        if (!sums.TryGetValue(keys[i], out cur)) cur = 0;
                        sums[keys[i]] = cur + v;
                    }
                }
            }
        }

        // 결과 테이블 (표시 순서 고정, active*Time은 Count 비움)
        private DataTable BuildResultTable(Dictionary<string, long> timeSum, Dictionary<string, long> countSum)
        {
            var order = new[] {
                "run","activeRun","dummyRun","activeDummyRun",
                "waiting","activeWaiting","idle","trouble","dummyTrouble",
                "setup","commDown","lotDown"
            };

            var dt = new DataTable();
            dt.Columns.Add("항목");
            dt.Columns.Add("시간(초)", typeof(long));
            dt.Columns.Add("분");        // 표시용(문자열) → 0.0 형식
            dt.Columns.Add("시간(h)");   // 표시용(문자열) → 0.0 형식
            dt.Columns.Add("Count");     // active* 계열은 빈 칸
            
            for (int i = 0; i < order.Length; i++)
            {
                string keyBase = order[i];
                string tKey = keyBase + "Time";    // e.g., runTime
                string cKey = keyBase + "Count";   // e.g., runCount
                long sec = 0;
                long v;
                if (timeSum.TryGetValue(tKey, out v)) sec = v;
                
                // 분/시간 계산 (소수 1자리, . 사용)
                double minutes = sec / 60.0;
                double hours = sec / 3600.0;
                string minutesText = minutes.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
                string hoursText = hours.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
                
                // Count: active 계열은 표시 비움
                string cntText = "";
                if (countSum.TryGetValue(cKey, out v))
                {
                    if (!keyBase.StartsWith("active", StringComparison.OrdinalIgnoreCase))
                        cntText = v.ToString();
                }
                dt.Rows.Add(DisplayName(keyBase), sec, minutesText, hoursText, cntText);
            }

            return dt;
        }

        private string DisplayName(string keyBase)
        {
            switch (keyBase)
            {
                case "run": return "Run";
                case "activeRun": return "Active Run";
                case "dummyRun": return "Dummy Run";
                case "activeDummyRun": return "Active Dummy Run";
                case "waiting": return "Waiting";
                case "activeWaiting": return "Active Waiting";
                case "idle": return "Idle";
                case "trouble": return "Trouble";
                case "dummyTrouble": return "Dummy Trouble";
                case "setup": return "Setup";
                case "commDown": return "Comm Down";
                case "lotDown": return "Lot Down";
            }

            return keyBase;
        }

        private ComboItem GetSelectedEqp() 
        { 
            return cboEqp.SelectedItem as ComboItem; 
        }

        // 결과 전체 복사(헤더 포함)
        private void CopyGridAll()
        {
            if (dgv.DataSource == null) return;
            dgv.SelectAll();
            var obj = dgv.GetClipboardContent();
            if (obj != null) Clipboard.SetDataObject(obj);
            dgv.ClearSelection();
        }

        // 단순 키/값 보관용
        private sealed class ComboItem
        {
            public string Text;
            public string Code;   // RLTC-01
            public string LineNo; // 2025-....
            public ComboItem(string text, string code, string lineNo) { Text = text; Code = code; LineNo = lineNo; }
            public override string ToString() { return Text; }
        }
    }
}