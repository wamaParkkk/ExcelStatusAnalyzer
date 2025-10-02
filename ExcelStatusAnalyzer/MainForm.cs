using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.MetaAttributes;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ExcelStatusAnalyzer
{
    public partial class MainForm : Form
    {
        private Button btnLoad;
        private OpenFileDialog ofd;
        private DataGridView dgvSource;
        private DataGridView dgvSummary;
        private Label lblTotal;
        private Label lblCheck;
        private Label lblDownTrouble;
        private TextBox txtDownTroubleHours;
                
        // 표준 상태명(표시용) → 매칭용 키워드 목록
        private readonly Dictionary<string, string[]> _statusMap = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["Comm Down"] = new[] { "Comm Down", "Commdown", "CommDown" },
            ["Run"] = new[] { "Run" },
            ["Trouble"] = new[] { "Trouble" },
            ["Waiting"] = new[] { "Waiting" },
            ["Dummy Run"] = new[] { "Dummy Run", "DummyRun" },
            ["Dummy Trouble"] = new[] { "Dummy Trouble", "DummyTrouble" },
        };
        
        public MainForm()
        {
            InitializeComponent();
            BuildUi();
        }
        
        private void BuildUi()
        {
            this.Text = "Excel Status Analyzer";
            this.Width = 1100;
            this.Height = 800;

            btnLoad = new Button
            {
                Text = "파일 불러오기(장비 상태 레포트) (.csv/.xlsx/xls)",
                Left = 15,
                Top = 15,
                Width = 300,
                Height = 32
            };
            btnLoad.Click += BtnLoad_Click;

            ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "엑셀 파일 선택"
            };

            dgvSource = new DataGridView
            {
                Left = 15,
                Top = 60,
                Width = this.ClientSize.Width - 30,
                Height = 320,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            dgvSummary = new DataGridView
            {
                Left = 15,
                Top = 390,
                Width = this.ClientSize.Width - 30,
                Height = 240,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = false,
                EditMode = DataGridViewEditMode.EditProgrammatically,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill                
            };

            lblTotal = new Label
            {
                Left = 330,
                Top = 22,
                Width = 220,
                Text = "총합: -"
            };

            lblCheck = new Label
            {
                Left = 560,
                Top = 22,
                Width = 250,
                Text = "검증: -"
            };

            lblDownTrouble = new Label
            {
                Left = 315,
                Top = 645,
                Width = 180,
                Text = "Down+Dummy+Trouble (h):"
            };

            txtDownTroubleHours = new TextBox
            {
                Left = 500,
                Top = 640,
                Width = 120,
                ReadOnly = true,
                TextAlign = HorizontalAlignment.Right
            };                    

            this.Controls.Add(btnLoad);
            this.Controls.Add(dgvSource);
            this.Controls.Add(dgvSummary);
            this.Controls.Add(lblTotal);
            this.Controls.Add(lblCheck);
            this.Controls.Add(lblDownTrouble);
            this.Controls.Add(txtDownTroubleHours);
            
            // Vision Retry Count
            var btnVisionRetryCount = new Button { Text = "Vision retry count", Left = 15, Top = 687, Width = 200, Height = 32 };
            btnVisionRetryCount.Click += (s, e) => new VisionRetryCountForm().Show();
            this.Controls.Add(btnVisionRetryCount);

            // Alarm Pivot
            var btnAlarmPivot = new Button { Text = "Alarm 일자별 집계", Left = 225, Top = 687, Width = 200, Height = 32 };
            btnAlarmPivot.Click += (s, e) => new AlarmPivotForm().Show();
            this.Controls.Add(btnAlarmPivot);
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var allRows = LoadAndParse(ofd.FileName);
                BindSourceGrid(allRows);

                var sums = SumByStatus(allRows);
                BindSummaryGrid(sums);

                // Comm Down / Dummy / Trouble 합산 후 시간(소수) 표시
                TimeSpan sDown = TryGetSum(sums, "Comm Down");
                TimeSpan sDummyRun = TryGetSum(sums, "Dummy Run");
                TimeSpan sDummyTr = TryGetSum(sums, "Dummy Trouble");
                TimeSpan sTrouble = TryGetSum(sums, "Trouble");
                // 분 단위 합산 후 60으로 나누기
                double totalMinutes = sDown.TotalMinutes + sDummyRun.TotalMinutes + sDummyTr.TotalMinutes + sTrouble.TotalMinutes;
                double hoursDecimal = totalMinutes / 60.0;                
                txtDownTroubleHours.Text = hoursDecimal.ToString("0.00", CultureInfo.InvariantCulture);

                var grand = new TimeSpan(sums.Values.Sum(ts => ts.Ticks));
                lblTotal.Text = $"총합: {FormatTS(grand)}";
                
                // 24시간 검증(±1초 허용)
                var diff = (grand - TimeSpan.FromHours(24));
                if (diff.Duration() <= TimeSpan.FromSeconds(1))
                    lblCheck.Text = "검증: OK (≈ 24:00:00)";
                else
                    lblCheck.Text = $"검증: 24시간과 {FormatTS(diff)} 차이";
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류: " + ex.Message, "읽기 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private sealed class RowItem
        {
            public string AssetNo { get; set; }
            public string NpiActive { get; set; }
            public string Model { get; set; }
            public string EquipmentNo { get; set; }
            public TimeSpan Start { get; set; }
            public TimeSpan End { get; set; }
            public string Status { get; set; }
            public TimeSpan Duration { get; set; }
        }

        private List<RowItem> LoadAndParse(string path)
        {
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheets.First();

                // 헤더 탐색 (첫 행 가정 또는 이름 일치 확인)
                var headerRow = ws.FirstRowUsed();
                var headerCells = headerRow.Cells().Select(c => c.GetString().Trim()).ToList();

                // 한국어 헤더명 매핑
                int idxAsset = FindIndex(headerCells, "자산 번호");
                int idxNpi = FindIndex(headerCells, "NPI 활성화");
                int idxModel = FindIndex(headerCells, "모델");
                int idxEq = FindIndex(headerCells, "장비 번호");
                int idxStart = FindIndex(headerCells, "시작 시간");
                int idxEnd = FindIndex(headerCells, "종료 시간");
                int idxStatus = FindIndex(headerCells, "상태");

                if (new[] { idxAsset, idxNpi, idxModel, idxEq, idxStart, idxEnd, idxStatus }.Any(i => i < 0))
                    throw new Exception("필요한 헤더(자산 번호, NPI 활성화, 모델, 장비 번호, 시작 시간, 종료 시간, 상태)를 찾을 수 없습니다.");

                var rows = new List<RowItem>();
                foreach (var row in ws.RowsUsed().Skip(1)) // 헤더 다음부터
                {
                    var cells = row.Cells().ToList();

                    var startTS = TryReadTime(cells, idxStart);
                    var endTS = TryReadTime(cells, idxEnd);

                    if (!startTS.HasValue || !endTS.HasValue)
                        continue; // 시간 불명확한 행은 스킵

                    var statusRaw = GetCellString(cells, idxStatus);
                    var statusNorm = NormalizeStatus(statusRaw);

                    var dur = endTS.Value - startTS.Value;
                    if (dur < TimeSpan.Zero) dur = dur + TimeSpan.FromDays(1); // 자정 넘김 보정

                    rows.Add(new RowItem
                    {
                        AssetNo = GetCellString(cells, idxAsset),
                        NpiActive = GetCellString(cells, idxNpi),
                        Model = GetCellString(cells, idxModel),
                        EquipmentNo = GetCellString(cells, idxEq),
                        Start = startTS.Value,
                        End = endTS.Value,
                        Status = statusNorm,
                        Duration = dur
                    });
                }

                return rows;
            }
        }

        private static int FindIndex(List<string> headers, string name)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (string.Equals(headers[i], name, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private TimeSpan? TryReadTime(List<IXLCell> cells, int idx)
        {
            if (idx < 0 || idx >= cells.Count) return null;
            var cell = cells[idx];

            // 1) 셀 자체가 시간/날짜면 ClosedXML이 파싱
            if (cell.DataType == XLDataType.DateTime)
                return cell.GetDateTime().TimeOfDay;

            // 2) 숫자(OA Date)일 가능성
            if (cell.DataType == XLDataType.Number)
            {
                var dt = DateTime.FromOADate(cell.GetDouble());
                return dt.TimeOfDay;
            }

            // 3) 문자열 파싱 (HH:mm:ss 또는 Excel이 문자열로 들고있는 날짜/시간)
            var s = cell.GetString().Trim();

            // 시:분:초
            if (TimeSpan.TryParseExact(s, new[] { @"h\:m\:s", @"hh\:mm\:ss", @"h\:mm\:ss", @"m\:s", @"mm\:ss" }, CultureInfo.InvariantCulture, out var ts))
                return ts;

            // 날짜+시간
            if (DateTime.TryParse(s, out var dtAny))
                return dtAny.TimeOfDay;

            return null;
        }

        private static string GetCellString(List<IXLCell> cells, int idx)
        {
            if (idx < 0 || idx >= cells.Count) return string.Empty;
            return cells[idx].GetString().Trim();
        }

        private string NormalizeStatus(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "기타";
            foreach (var kv in _statusMap)
            {
                if (kv.Value.Any(alias => raw.Trim().Equals(alias, StringComparison.OrdinalIgnoreCase)))
                    return kv.Key; // 표준 상태명으로 통일
            }
            return "기타";
        }

        private void BindSourceGrid(List<RowItem> rows)
        {
            var tbl = new DataTable();
            tbl.Columns.Add("자산번호");
            tbl.Columns.Add("NPI 활성화");
            tbl.Columns.Add("모델");
            tbl.Columns.Add("장비 번호");
            tbl.Columns.Add("시작 시간");
            tbl.Columns.Add("종료 시간");
            tbl.Columns.Add("상태");
            tbl.Columns.Add("구간 합계");

            foreach (var r in rows)
            {
                tbl.Rows.Add(
                    r.AssetNo,
                    r.NpiActive,
                    r.Model,
                    r.EquipmentNo,
                    FormatTS(r.Start),
                    FormatTS(r.End),
                    r.Status,
                    FormatTS(r.Duration)
                );
            }
            dgvSource.DataSource = tbl;
        }

        private Dictionary<string, TimeSpan> SumByStatus(List<RowItem> rows)
        {
            var sums = new Dictionary<string, TimeSpan>(StringComparer.OrdinalIgnoreCase);

            foreach (var r in rows)
            {
                if (!sums.ContainsKey(r.Status))
                    sums[r.Status] = TimeSpan.Zero;

                sums[r.Status] = sums[r.Status] + r.Duration;
            }

            // 결과에 없는 상태도 0으로 보여주고 싶다면 아래 주석 해제
            foreach (var std in _statusMap.Keys)
            {
                if (!sums.ContainsKey(std))
                    sums[std] = TimeSpan.Zero;
            }

            return sums
                .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(k => k.Key, v => v.Value, StringComparer.OrdinalIgnoreCase);
        }

        private void BindSummaryGrid(Dictionary<string, TimeSpan> sums)
        {
            var tbl = new DataTable();
            tbl.Columns.Add("상태");
            tbl.Columns.Add("합계");
            tbl.Columns.Add("비율");

            var grand = new TimeSpan(sums.Values.Sum(ts => ts.Ticks));
            double grandSec = Math.Max(1, grand.TotalSeconds); // 0 분모 방지

            foreach (var kv in sums)
            {
                double pct = kv.Value.TotalSeconds / grandSec * 100.0;
                tbl.Rows.Add(kv.Key, FormatTS(kv.Value), pct.ToString("0.0") + "%");
            }

            dgvSummary.DataSource = tbl;
        }

        private static TimeSpan TryGetSum(Dictionary<string, TimeSpan> sums, string key)
        {
            TimeSpan v;
            return sums != null && sums.TryGetValue(key, out v) ? v : TimeSpan.Zero;
        }

        private static string FormatTS(TimeSpan ts)
        {
            // 시 한 자리면 0을 생략: 예) 0:04:15
            // 100시간 이상도 안전하게 표시
            long totalHours = (long)ts.TotalHours;
            var remain = ts - TimeSpan.FromHours(totalHours);
            return string.Format("{0}:{1:00}:{2:00}", totalHours, remain.Minutes, remain.Seconds);
        }
    }
}