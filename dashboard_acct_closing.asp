<!-- DATATABLES -->
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

<!-- CHART JS -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<style>
    /* Semua aturan di-scope ke .acd agar tidak bocor ke elemen lain di halaman */
    .acd {
        --brand: #1B7863;
        --brand-dark: #12584a;
        --brand-tint: #e8f3f0;
        --ink: #17262a;
        --muted: #5d6d69;
        --line: #e1e7e5;
        --good: #146c2e;
        --good-bg: #e6f4ea;
        --info: #1d4f91;
        --info-bg: #e8f0fb;
        --warn: #8a5a00;
        --warn-bg: #fff4d6;
        --bad: #b42318;
        --bad-bg: #fdecec;
        color: var(--ink);
        font-family: 'Plus Jakarta Sans', 'Segoe UI', Roboto, Arial, sans-serif;
        font-size: 14px;
    }
    .acd h2, .acd h3, .acd h4 { font-family: inherit; color: var(--ink); margin: 0; }

    /* Judul + filter periode */
    .acd-toolbar {
        display: flex; flex-wrap: wrap; align-items: flex-end; justify-content: space-between; gap: 16px;
        margin-bottom: 16px;
    }
    .acd-toolbar h2 { font-size: 20px; font-weight: 700; }
    .acd-toolbar p { margin: 2px 0 0; color: var(--muted); font-size: 13px; }
    .acd-filters { display: flex; gap: 10px; }
    .acd-filters label { display: block; font-size: 12px; font-weight: 600; color: var(--muted); margin-bottom: 4px; }
    .acd-filters select {
        height: 38px; min-width: 140px;
        padding: 0 36px 0 12px;
        font: inherit; font-size: 14px; font-weight: 600; color: var(--ink);
        background: #fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1.5l5 5 5-5' fill='none' stroke='%235d6d69' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E") no-repeat right 12px center;
        border: 1px solid var(--line); border-radius: 8px;
        -webkit-appearance: none; appearance: none;
        cursor: pointer;
    }
    .acd-filters select:focus { outline: none; border-color: var(--brand); box-shadow: 0 0 0 3px rgba(27, 120, 99, .18); }

    /* Kartu ringkasan */
    .acd-kpis {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
        gap: 14px;
        margin-bottom: 24px;
    }
    .acd-kpi {
        background: #fff;
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 16px 18px;
        display: flex; flex-direction: column; gap: 10px;
    }
    .acd-kpi-head { display: flex; align-items: center; gap: 10px; font-size: 13px; font-weight: 600; color: var(--muted); }
    .acd-kpi-icon {
        width: 32px; height: 32px; border-radius: 9px;
        display: inline-flex; align-items: center; justify-content: center;
        font-size: 15px;
    }
    .acd-kpi-value { font-size: 30px; font-weight: 700; line-height: 1.1; color: var(--ink); }
    .acd-kpi-value.small { font-size: 24px; }
    .acd-kpi-note { font-size: 12px; color: var(--muted); }
    .tone-brand { background: var(--brand-tint); color: var(--brand); }
    .tone-good  { background: var(--good-bg); color: var(--good); }
    .tone-bad   { background: var(--bad-bg); color: var(--bad); }

    /* Judul bagian */
    .acd-section-title {
        display: flex; align-items: center; gap: 10px;
        margin: 8px 0 12px;
        font-size: 12px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;
        color: var(--muted);
    }
    .acd-section-title::after { content: ""; flex: 1; height: 1px; background: var(--line); }

    .acd-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 16px;
        margin-bottom: 16px;
    }
    .acd-grid.wide-left { grid-template-columns: minmax(0, 3fr) minmax(0, 2fr); }
    @media (max-width: 1200px) { .acd-grid, .acd-grid.wide-left { grid-template-columns: minmax(0, 1fr); } }

    .acd-card {
        background: #fff;
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 18px 20px 20px;
        min-width: 0;
    }
    .acd-card-head {
        display: flex; align-items: center; justify-content: space-between; gap: 12px;
        margin-bottom: 14px;
    }
    .acd-card-head h4 { font-size: 15px; font-weight: 700; }
    .acd-card-head p { margin: 2px 0 0; font-size: 12px; color: var(--muted); }
    .acd-btn {
        display: inline-flex; align-items: center; gap: 6px;
        padding: 7px 12px;
        font: inherit; font-size: 13px; font-weight: 600;
        color: var(--brand);
        background: #fff;
        border: 1px solid #b9d9d0; border-radius: 8px;
        cursor: pointer;
        white-space: nowrap;
    }
    .acd-btn:hover { background: var(--brand-tint); }

    /* Tabel */
    .acd-table-wrap { overflow: auto; max-height: 420px; border: 1px solid var(--line); border-radius: 10px; }
    .acd table.table-bordered { border-collapse: collapse; width: 100%; font-size: 13px; }
    .acd table.table-bordered th {
        position: sticky; top: 0;
        padding: 10px 10px;
        background: #f6f9f8;
        color: var(--muted);
        font-size: 11px; font-weight: 700; letter-spacing: .5px; text-transform: uppercase;
        text-align: center;
        border: 0; border-bottom: 1px solid var(--line);
        white-space: nowrap;
    }
    .acd table.table-bordered td {
        padding: 8px 10px;
        border: 0; border-bottom: 1px solid #eef2f1;
        text-align: center;
        font-variant-numeric: tabular-nums;
    }
    .acd table.table-bordered tbody tr:last-child td { border-bottom: 0; }
    .acd table.table-bordered tbody tr:hover td { background: #f8fbfa; }
    .acd table.table-bordered a { color: var(--brand); font-weight: 600; }

    .badge {
        display: inline-flex; align-items: center; gap: 5px;
        padding: 3px 9px;
        border-radius: 999px;
        font-size: 12px; font-weight: 600;
        white-space: nowrap;
    }
    .badge-good { background: var(--good-bg); color: var(--good); }
    .badge-info { background: var(--info-bg); color: var(--info); }
    .badge-warn { background: var(--warn-bg); color: var(--warn); }
    .badge-bad  { background: var(--bad-bg); color: var(--bad); }

    .acd-chart { position: relative; }
    .acd-heat-legend { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 10px; font-size: 12px; color: var(--muted); }
    .acd-heat-legend span { display: inline-flex; align-items: center; gap: 6px; }
    .acd-heat-legend i { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }
</style>

<div class="acd">
    <%
        Dim bln, thn
        bln = Month(Date)
        thn = Year(Date)
        dim userlogins1
        userlogins1 = Session("CodeName")
        if userlogins1="STAFF_OFFICER_PURCH" OR userlogins1="STAFF_OFFICER_PURCH_02" OR userlogins1="DIV_HEAD_COMM" OR userlogins1="IT_HO6" OR userlogins1="IT_HO2" OR userlogins1="SISKA_MEILINA_SEMBIRING" OR userlogins1="RINTO_PURNAMASIDI" OR userlogins1="DEPT_HEAD_PURCH" THEN
    %>
        <input type="hidden" id="sessionfcba" value="-" />
    <%
        else
    %>
        <input type="hidden" id="sessionfcba" value="<%=Session("fcba")%>" />
    <%
        end if
    %>

    <div class="acd-toolbar">
        <div>
            <h2>Dashboard Closing Accounting</h2>
            <p>Ketepatan waktu closing modul per PT pada periode yang dipilih.</p>
        </div>
        <div class="acd-filters">
            <div>
                <label for="month-selected">Bulan</label>
                <select name="month" id="month-selected">
                    <option value="1" <%if bln = 1 then%> selected <% end if %>>Januari</option>
                    <option value="2" <%if bln = 2 then%> selected <% end if %>>Februari</option>
                    <option value="3" <%if bln = 3 then%> selected <% end if %>>Maret</option>
                    <option value="4" <%if bln = 4 then%> selected <% end if %>>April</option>
                    <option value="5" <%if bln = 5 then%> selected <% end if %>>Mei</option>
                    <option value="6" <%if bln = 6 then%> selected <% end if %>>Juni</option>
                    <option value="7" <%if bln = 7 then%> selected <% end if %>>Juli</option>
                    <option value="8" <%if bln = 8 then%> selected <% end if %>>Agustus</option>
                    <option value="9" <%if bln = 9 then%> selected <% end if %>>September</option>
                    <option value="10" <%if bln = 10 then%> selected <% end if %>>Oktober</option>
                    <option value="11" <%if bln = 11 then%> selected <% end if %>>November</option>
                    <option value="12" <%if bln = 12 then%> selected <% end if %>>Desember</option>
                </select>
            </div>
            <div>
                <label for="year-selected">Tahun</label>
                <select name="year" id="year-selected">
                    <option value="2024" <%if thn = 2024 then%> selected <% end if %>>2024</option>
                    <option value="2025" <%if thn = 2025 then%> selected <% end if %>>2025</option>
                    <option value="2026" <%if thn = 2026 then%> selected <% end if %>>2026</option>
                    <option value="2027" <%if thn = 2027 then%> selected <% end if %>>2027</option>
                    <option value="2028" <%if thn = 2028 then%> selected <% end if %>>2028</option>
                    <option value="2029" <%if thn = 2029 then%> selected <% end if %>>2029</option>
                    <option value="2030" <%if thn = 2030 then%> selected <% end if %>>2030</option>
                </select>
            </div>
        </div>
    </div>

    <!-- SUMMARY -->
    <div class="acd-kpis">
        <div class="acd-kpi">
            <div class="acd-kpi-head"><span class="acd-kpi-icon tone-brand"><i class="fa fa-bar-chart-o" aria-hidden="true"></i></span>Summary Bulanan</div>
            <div id="summary-bulanan" class="acd-kpi-value">0</div>
            <div class="acd-kpi-note">Rata-rata skor semua PT</div>
        </div>
        <div class="acd-kpi">
            <div class="acd-kpi-head"><span class="acd-kpi-icon tone-good"><i class="fa fa-check-circle" aria-hidden="true"></i></span>Tepat Waktu</div>
            <div class="acd-kpi-value" id="total-above">0</div>
            <div class="acd-kpi-note">Jumlah PT dengan skor &gt; 90</div>
        </div>
        <div class="acd-kpi">
            <div class="acd-kpi-head"><span class="acd-kpi-icon tone-bad"><i class="fa fa-exclamation-circle" aria-hidden="true"></i></span>Telat</div>
            <div class="acd-kpi-value" id="total-below-60">0</div>
            <div class="acd-kpi-note">Jumlah PT dengan skor &lt; 60</div>
        </div>
        <div class="acd-kpi">
            <div class="acd-kpi-head"><span class="acd-kpi-icon tone-good"><i class="fa fa-arrow-up" aria-hidden="true"></i></span>Teratas</div>
            <div class="acd-kpi-value small"><span id="top-rank">-</span></div>
            <div class="acd-kpi-note">PT dengan skor tertinggi</div>
        </div>
        <div class="acd-kpi">
            <div class="acd-kpi-head"><span class="acd-kpi-icon tone-bad"><i class="fa fa-arrow-down" aria-hidden="true"></i></span>Terendah</div>
            <div class="acd-kpi-value small"><span id="loss-rank">-</span></div>
            <div class="acd-kpi-note">PT dengan skor terendah</div>
        </div>
    </div>

    <div class="acd-section-title">Closing Accounting</div>

    <!-- TABLE + BAR CHART -->
    <div class="acd-grid wide-left">
        <div class="acd-card">
            <div class="acd-card-head">
                <div>
                    <h4>Skor Modul per PT</h4>
                    <p>Skor tiap modul (maks. 100), diurutkan dari skor tertinggi</p>
                </div>
            </div>
            <div class="acd-table-wrap">
                <table id="tablePT" class="display table-bordered">
                    <thead>
                        <tr>
                            <th>PT</th>
                            <th>ARESTA</th>
                            <th>STOK</th>
                            <th>WIP</th>
                            <th>WSVRA</th>
                            <th>Kas Bank</th>
                            <th>Score</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody id="bodytablePT">
                    </tbody>
                </table>
            </div>
        </div>

        <div class="acd-card">
            <div class="acd-card-head">
                <div>
                    <h4>Grafik Skor per PT</h4>
                    <p>Skor closing bulan terpilih</p>
                </div>
            </div>
            <div class="acd-chart" style="height: 380px;"><canvas id="barPT"></canvas></div>
        </div>
    </div>

    <div class="acd-grid">
        <!-- HEAT MAP -->
        <div class="acd-card">
            <div class="acd-card-head">
                <div>
                    <h4>Trend Modul 12 Bulan</h4>
                    <p>Skor per PT per bulan pada tahun terpilih</p>
                </div>
            </div>
            <div class="acd-chart" id="heatmapWrap" style="height: 200px;"><canvas id="heatmapPT"></canvas></div>
            <div class="acd-heat-legend">
                <span><i style="background:#e9ecea"></i>0 / belum ada</span>
                <span><i style="background:#c0392b"></i>&le; 50</span>
                <span><i style="background:#ec835a"></i>51&ndash;60</span>
                <span><i style="background:#fab219"></i>61&ndash;70</span>
                <span><i style="background:#8cc97a"></i>71&ndash;80</span>
                <span><i style="background:#1f8f4e"></i>&gt; 80</span>
            </div>
        </div>
        <!-- TREND LINE -->
        <div class="acd-card">
            <div class="acd-card-head">
                <div>
                    <h4>Trend Closing per PT</h4>
                    <p>Klik nama PT di legenda untuk menyembunyikan / menampilkan garis</p>
                </div>
            </div>
            <div class="acd-chart" style="height: 420px;"><canvas id="linePT"></canvas></div>
        </div>
    </div>

    <div class="acd-section-title">Procurement</div>

    <div class="acd-grid">
        <div class="acd-card">
            <div class="acd-card-head">
                <div>
                    <h4>Top 10 PR/PO</h4>
                    <p>Dokumen dengan durasi approval terlama</p>
                </div>
                <button id="btnDownloadPR" class="acd-btn" onclick="downloadExcel('top5pr', 'Top10_PR_PO')">
                    <i class="fa fa-download" aria-hidden="true"></i> Download Excel
                </button>
            </div>
            <div class="acd-table-wrap">
                <table id="top5pr" class="display table-bordered">
                    <thead>
                        <tr>
                            <th>PT</th>
                            <th>ESTATE</th>
                            <th>DOC NO</th>
                            <th>Durasi</th>
                            <th>Doc. Class</th>
                        </tr>
                    </thead>
                    <tbody id="bodytablePR">
                    </tbody>
                </table>
            </div>
        </div>

        <div class="acd-card">
            <div class="acd-card-head">
                <div>
                    <h4>Top 10 Belum GRN</h4>
                    <p>PO terlama yang belum diterima barangnya</p>
                </div>
                <button id="btnDownloadGRN" class="acd-btn" onclick="downloadExcel('top5po', 'Top10_Belum_GRN')">
                    <i class="fa fa-download" aria-hidden="true"></i> Download Excel
                </button>
            </div>
            <div class="acd-table-wrap">
                <table id="top5po" class="display table-bordered">
                    <thead>
                        <tr>
                            <th>PT</th>
                            <th>ESTATE</th>
                            <th>NO. PO</th>
                            <th>Durasi</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody id="bodytableGRN">
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<link rel="stylesheet" href="css/jquery-ui.css" />
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-chart-matrix@1.3.0/dist/chartjs-chart-matrix.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-chart-matrix@1.3.0"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
<script type="text/javascript">
    /* Gaya grafik: font, warna teks & grid yang tenang */
    Chart.defaults.font.family = "'Plus Jakarta Sans', 'Segoe UI', Roboto, Arial, sans-serif";
    Chart.defaults.font.size = 12;
    Chart.defaults.color = '#5d6d69';
    Chart.defaults.borderColor = '#e9eeec';
    Chart.defaults.plugins.tooltip.backgroundColor = '#17262a';
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.cornerRadius = 8;

    /* Palet kategori (urutan tetap). Lebih dari 8 PT: warna diulang dengan pola garis berbeda */
    var SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"];
    var SERIES_DASHES = [[], [6, 4], [2, 3]];

    function statusBadge(avg) {
        if (avg === 100) return "<span class='badge badge-good'><i class='fa fa-check-circle'></i> Tepat Waktu</span>";
        if (avg >= 80)   return "<span class='badge badge-info'><i class='fa fa-clock-o'></i> Telat Ringan</span>";
        if (avg >= 60)   return "<span class='badge badge-warn'><i class='fa fa-exclamation-triangle'></i> Telat Sedang</span>";
        return "<span class='badge badge-bad'><i class='fa fa-times-circle'></i> Telat Berat</span>";
    }

    /* Datatable init */
    $(document).ready(function () {
        let tahun = $('#year-selected').val();
        let month = $('#month-selected').val();
        tablePT(tahun, month);
        LineChart(tahun);
        tablePR();
        tableGRN();
    });

    $('#year-selected').on('change', function(){
        let month = $('#month-selected').val();
        let tahun = $(this).val();
        tablePT(tahun, month);
        LineChart(tahun);
    });

    $('#month-selected').on('change', function(){
        let month = $(this).val();
        let tahun = $('#year-selected').val();
        tablePT(tahun, month);
    });

    function downloadExcel(tableId, fileName) {
        var table = document.getElementById(tableId);

        // Ambil header
        var headers = [];
        table.querySelectorAll('thead tr th').forEach(function(th) {
            headers.push(th.innerText.trim());
        });

        // Ambil rows
        var rows = [];
        table.querySelectorAll('tbody tr').forEach(function(tr) {
            var row = [];
            tr.querySelectorAll('td').forEach(function(td) {
                // Ambil teks saja (tanpa tag HTML seperti <a>)
                row.push(td.innerText.trim());
            });
            rows.push(row);
        });

        // Gabungkan header + rows
        var wsData = [headers].concat(rows);

        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet(wsData);

        // Auto column width
        var colWidths = headers.map(function(h, i) {
            var maxLen = h.length;
            rows.forEach(function(r) {
                if (r[i] && r[i].length > maxLen) maxLen = r[i].length;
            });
            return { wch: maxLen + 4 };
        });
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, fileName.substring(0, 31));
        XLSX.writeFile(wb, fileName + '.xlsx');
    }

    function tablePR() {
        $.post("datadashboard.asp?idview=4&fcba=" + $("#sessionfcba").val(), function (data) {

            $('#bodytablePR').empty();

            var json = JSON.parse(data);

            if (json.results && json.results.length > 0) {
                var mapped = json.results;
                mapped.forEach(row => {
                    var docketlist = '';
                    docketlist += '<tr>';
                    docketlist += '<td align="center">' + row.fccompanycode + '</td>';
                    docketlist += '<td style="text-align: left">' + row.fcba + '</td>';
                    docketlist += '<td style="text-align: left"><a href="PrintPR.Asp?PrCode='+row.prcode+'&fcba='+row.fcba+'&Mode=Print" target="_blank">' + row.prcode + '</a></td>';
                    docketlist += '<td style="text-align: right">' + row.days_pr_to_approve_900 + ' Hari</td>';
                    docketlist += '<td align="center">' + row.doc_classification_name + '</td>';
                    docketlist += '</tr>';

                    $('#bodytablePR').append(docketlist);
                });
            } else {
                $('#bodytablePR').append(
                    '<tr><td colspan="5" style="text-align: center; color: #5d6d69; padding: 24px;">Data Tidak Ada</td></tr>'
                );
            }
        });
    }

    function tableGRN() {
        $.post("datadashboard.asp?idview=5&fcba=" + $("#sessionfcba").val(), function (data) {

            $('#bodytableGRN').empty();

            var json = JSON.parse(data);

            if (json.results && json.results.length > 0) {
                var mapped = json.results;
                mapped.forEach(row => {
                    var docketlist = '';
                    docketlist += '<tr>';
                    docketlist += '<td align="center">' + row.fccompanycode + '</td>';
                    docketlist += '<td style="text-align: left">' + row.fcba + '</td>';
                    docketlist += '<td style="text-align: left"><a href="printingpo.Asp?pocode='+row.pocode+'&fcba='+row.fcba+'&purchaseby='+row.purchaseby+'&deliverydestination='+row.deliverydestination+'" target="_blank">' + row.pocode + '</a></td>';
                    docketlist += '<td align="center">' + row.days_po_to_grn + ' Hari</td>';
                    docketlist += '<td align="center">' + row.status_grn + '</td>';
                    docketlist += '</tr>';

                    $('#bodytableGRN').append(docketlist);
                });
            } else {
                $('#bodytableGRN').append(
                    '<tr><td colspan="5" style="text-align: center; color: #5d6d69; padding: 24px;">Data Tidak Ada</td></tr>'
                );
            }
        });
    }

    function tablePT(tahun, month) {
        $.post("datadashboard.asp?idview=1&tahun=" + tahun + "&month=" + month, function (data) {

            $('#bodytablePT').empty();

            var labels = [];
            var scores = [];

            var json = JSON.parse(data);

            var totalScore = 0;

            if (json.results && json.results.length > 0) {

                // mapping total nilai
                var mapped = json.results.map(row => {

                var total = row.score;
                var avg = Math.min(total, 100);

                return {
                    fccompanycode: row.fccompanycode,
                    aresta_closing: row.aresta_closing,
                    stock_closing: row.stock_closing,
                    wip_closing: row.wip_closing,
                    wsvra_closing: row.wsvra_closing,
                    kasbank_closing: row.kasbank_closing,
                    avg: avg,
                    status: statusBadge(avg)
                };
            });

                var totalOnTime = mapped.filter(row => {
                    return row.avg > 90;
                }).length;

                $('#total-above').text(totalOnTime);

                var totalTelatBerat = mapped.filter(row => {
                    return row.avg < 60;
                }).length;

                $('#total-below-60').text(totalTelatBerat);

                // =====================
                // SORT DESC BY AVG
                // =====================
                mapped.sort((a, b) => b.avg - a.avg);

                // =====================
                // TOP & LOSS RANK
                // =====================
                $('#top-rank').html(mapped[0].fccompanycode);
                $('#loss-rank').html(mapped[mapped.length - 1].fccompanycode);

                // =====================
                // TABLE + CHART DATA
                // =====================
                mapped.forEach(row => {
                    totalScore += row.avg;
                    labels.push(row.fccompanycode);
                    scores.push(row.avg);

                    var docketlist = '';
                    docketlist += '<tr>';
                    docketlist += '<td align="center"><b>' + row.fccompanycode + '</b></td>';
                    docketlist += '<td align="center">' + Math.min(parseFloat(row.aresta_closing), 100).toFixed(0) + '</td>';
                    docketlist += '<td align="center">' + Math.min(parseFloat(row.stock_closing), 100).toFixed(0) + '</td>';
                    docketlist += '<td align="center">' + Math.min(parseFloat(row.wip_closing), 100).toFixed(0) + '</td>';
                    docketlist += '<td align="center">' + Math.min(parseFloat(row.wsvra_closing), 100).toFixed(0) + '</td>';
                    docketlist += '<td align="center">' + Math.min(parseFloat(row.kasbank_closing), 100).toFixed(0) + '</td>';
                    docketlist += '<td align="center"><b>' + row.avg + '</b></td>';
                    docketlist += '<td align="center">' + row.status + '</td>';
                    docketlist += '</tr>';

                    $('#bodytablePT').append(docketlist);
                });

                var summaryBulanan = mapped.length > 0
                ? totalScore / mapped.length
                : 0;

                $('#summary-bulanan').text(summaryBulanan.toFixed(2));
            }

            if (window.barChart) window.barChart.destroy();

            window.barChart = new Chart(
                document.getElementById("barPT"),
                {
                    type: "bar",
                    data: {
                        labels: labels,
                        datasets: [{
                            label: "Score",
                            backgroundColor: "#1B7863",
                            hoverBackgroundColor: "#12584a",
                            borderRadius: { topLeft: 4, topRight: 4 },
                            borderSkipped: 'bottom',
                            maxBarThickness: 32,
                            data: scores
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { display: false }
                        },
                        scales: {
                            x: {
                                grid: { display: false }
                            },
                            y: {
                                beginAtZero: true,
                                max: 100,
                                border: { display: false },
                                ticks: { stepSize: 20 }
                            }
                        }
                    }
                }
            );
        });
    }

    function LineChart(tahun) {
        $.post("datadashboard.asp?idview=2&tahun=" + tahun, function (data) {

            var labelsMonth = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OKT","NOV","DEC"];
            var ptMap = {};

            var json = JSON.parse(data);
            // =====================
            // LINE CHART (PER BULAN)
            // =====================
            json.results.forEach(function (row) {

                var pt = row.fccompanycode;
                var bulan = parseInt(row.month) - 1;

                var avg = Math.min(row.score, 100);

                if (!ptMap[pt]) {
                    var idx = Object.keys(ptMap).length;
                    var color = SERIES_COLORS[idx % SERIES_COLORS.length];
                    ptMap[pt] = {
                        label: pt,
                        borderColor: color,
                        backgroundColor: color,
                        borderDash: SERIES_DASHES[Math.floor(idx / SERIES_COLORS.length) % SERIES_DASHES.length],
                        borderWidth: 2,
                        pointRadius: 3,
                        pointHoverRadius: 5,
                        fill: false,
                        tension: 0.3,
                        cubicInterpolationMode: 'monotone',
                        data: new Array(12).fill(null)
                    };
                }

                ptMap[pt].data[bulan] = avg;
            });

            var datasets = Object.values(ptMap);

            if (window.lineChart) window.lineChart.destroy();

            window.lineChart = new Chart(
                document.getElementById("linePT"),
                {
                    type: "line",
                    data: {
                        labels: labelsMonth,
                        datasets: datasets
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        interaction: { mode: 'index', intersect: false },
                        plugins: {
                            legend: {
                                position: 'bottom',
                                labels: { usePointStyle: true, pointStyle: 'line', boxWidth: 24, padding: 14 }
                            }
                        },
                        scales: {
                            x: { grid: { display: false } },
                            y: {
                                beginAtZero: true,
                                max: 100,
                                border: { display: false }
                            }
                        }
                    }
                }
            );

            var ptList = Object.keys(ptMap);

            // susun data heatmap
            var heatmapData = [];

            ptList.forEach(function (pt) {
                labelsMonth.forEach(function (bulan) {
                    heatmapData.push({
                        x: bulan,
                        y: pt,
                        v: 0
                    });
                });
            });

            json.results.forEach(function (row) {

                var pt = row.fccompanycode;
                var bulanIndex = parseInt(row.month) - 1;

                var avg = Math.min(row.score, 100);

                var cell = heatmapData.find(d =>
                    d.x === labelsMonth[bulanIndex] &&
                    d.y === pt
                );

                if (cell) cell.v = avg;
            });

            // Tinggi heatmap mengikuti jumlah PT agar baris tidak berdesakan
            document.getElementById("heatmapWrap").style.height = Math.max(200, ptList.length * 30 + 50) + "px";

            // destroy jika sudah ada
            const canvas = document.getElementById("heatmapPT");

            // destroy via registry Chart.js
            const existingChart = Chart.getChart(canvas);
            if (existingChart) {
                existingChart.destroy();
            }

            function heatColor(v) {
                if (v == 0) return '#e9ecea';   // belum ada
                if (v <= 50) return '#c0392b';  // merah
                if (v <= 60) return '#ec835a';  // oranye
                if (v <= 70) return '#fab219';  // kuning
                if (v <= 80) return '#8cc97a';  // hijau muda
                return '#1f8f4e';               // hijau tua
            }

            window.heatmapChart = new Chart(
                document.getElementById("heatmapPT"),
                {
                    type: 'matrix',
                    data: {
                        datasets: [{
                            data: heatmapData,
                            backgroundColor: function (ctx) {
                                return heatColor(ctx.raw.v);
                            },
                            borderColor: '#ffffff',
                            borderWidth: 2,
                            borderRadius: 4,
                            anchor: 'center',
                            align: 'center',
                            width: ctx => {
                                if (!ctx.chart.chartArea) return 20;
                                return (ctx.chart.chartArea.width / labelsMonth.length) - 2; // dibagi 12
                            },
                            height: ctx => {
                                if (!ctx.chart.chartArea || ptList.length === 0) return 30;
                                return (ctx.chart.chartArea.height / ptList.length) - 2; // dinamis sesuai jumlah PT
                            }
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        layout: {
                            padding: {
                                top: 4,
                                bottom: 0
                            }
                        },
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                callbacks: {
                                    title: items => items[0].raw.y + ' - ' + items[0].raw.x,
                                    label: item => 'Skor: ' + item.raw.v
                                }
                            },
                            datalabels: {
                                // teks putih di sel gelap, gelap di sel terang
                                color: ctx => {
                                    var v = ctx.dataset.data[ctx.dataIndex].v;
                                    return (v > 0 && v <= 50) || v > 80 ? '#ffffff' : '#17262a';
                                },
                                font: {
                                    weight: '600',
                                    size: 11
                                },
                                formatter: v => v.v
                            }
                        },
                        scales: {
                            x: {
                                type: 'category',
                                labels: labelsMonth,
                                position: 'top',
                                grid: { display: false },
                                border: { display: false }
                            },
                            y: {
                                type: 'category',
                                labels: ptList,
                                offset: true,       // supaya row rata & tidak numpuk
                                grid: { display: false },
                                border: { display: false },
                                ticks: {
                                    padding: 5
                                }
                            }
                        }
                    },
                    plugins: [ChartDataLabels]
                }
            );
        });
    }

</script>
