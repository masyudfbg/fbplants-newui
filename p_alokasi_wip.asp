<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="id">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>FBPlants Reporting - WIP Listing</title>

<!-- CSS -->
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<style>
    body {
        font-family: "Segoe UI", Tahoma, Arial, sans-serif;
        font-size: 12px;
        color: #333;
        margin: 20px;
        background-color: #f9f9f9;
    }
    h4, h5 {
        margin: 0;
        color: #2c3e50;
    }
    #header-wrapper {
        padding: 10px 0;
        border-bottom: 2px solid #2c3e50;
        margin-bottom: 15px;
    }
    #product-table {
        width: 100%;
        border-collapse: collapse;
        background-color: #fff;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    #product-table thead th {
        background-color: #2c3e50;
        color: #fff;
        font-size: 11px;
        font-weight: bold;
        text-align: center;
        padding: 6px;
        border: 1px solid #ddd;
    }
    
   
    .printhide {
        display: none;
    }
    @media print {
        body { background: #fff; }
        #product-table { box-shadow: none; }
        #header-wrapper { border: none; }
    }

/* freeze baris 1 */
.product-table thead tr:first-child th {
  background: #2c3e50;
  color: #fff;
  font-weight: 600;
  position: sticky;
  top: 0;
  z-index: 10;
  text-align: center;
}

/* freeze baris 2 */
.product-table thead tr:nth-child(2) th {
  background: #34495e; /* beda warna biar jelas */
  color: #fff;
  position: sticky;
  top: 28px;  /* sesuaikan dengan tinggi baris pertama */
  z-index: 9;
  text-align: center;
}

.product-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}

.product-table th, 
.product-table td {
  border: 1px solid #ddd;
  padding: 6px 8px;
  text-align: right;
}

.product-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}

.product-table th, 
.product-table td {
  border: 1px solid #ddd;
  padding: 6px 8px;
  text-align: right;
}
.export-btn {
    position: fixed;
    top: 20px;   /* jarak dari atas */
    right: 20px; /* jarak dari kanan */
    z-index: 9999; /* supaya tidak ketutup table */
}
.ratakiri{
	text-align:left;
}
.ratatengah{
	text-align:center;
}
.ratakanan{
	text-align:right;
}
.tebalhuruf{
	font-weight: bold;
}
.warnabaristotal{
	background: #f3e745ff; 
}
</style>

<!-- JS -->
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script src="/js/x.js"></script>
<script src="/js/xtableheaderfixed.js"></script>
<script>
  $(function() {
    new xTableHeaderFixed('product-table', window);
});


var tableToExcel = (function () {
  var encabezado = '<html><head><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/><style> table, td {border:thin solid black} table {border-collapse:collapse} td {mso-number-format:"\\@";}</style></head><body><table>';
  var dataTable = document.getElementById('producttable').innerHTML
  var piePagina = "</table></body></html>";
  var tabla = encabezado + dataTable + piePagina;
  var myBlob =  new Blob( [tabla] , {type:'text/html'});
  var url = window.URL.createObjectURL(myBlob);
  var a = document.createElement("a");
  document.body.appendChild(a);
  a.href = url;
  a.download = "export_excel.xls";
  a.click();
  setTimeout(function() {window.URL.revokeObjectURL(url);},0);
});

</script>
</head>

<%
	Dim pcompany,fcba,txtmonth,txtyear,q,qq,no,rsdata,filter,filteragrementcode,filtercontractor,fcbajudul
    pcompany  = Request("companywip")
    fcba      = Request("businessunitwip")
    txtmonth  = Request("txtmonth")
    txtyear    = Request("txtyear")
    agreementcode    = Request("agreementcode")
    contractorcode    = Request("contractorcode")

    q = "SELECT fccode, fccompanycode, FCNAME, FCCOMPANYNAME, FCMANAGER, FCKTU, FCCROP, FCTYPE FROM BusinessUnit WHERE fcCode='" & fcba & "' "
    set rsCat = connStr.Execute(q)
    if not rsCat.eof then
        CompanyNamedb = rsCat("FCCOMPANYNAME")
        fcbaname    = rsCat("FCNAME")
		
    end if
	if CompanyNamedb = "" then 
			CompanyName = pcompany
		ELSE 
			CompanyName = CompanyNamedb
	end if
	if fcba = "-" then 
			fcbajudul = "ALL"
		ELSE 
			fcbajudul = fcbaname
	end if

	if trim(fcba) = "-" then 
		filter = " AND W.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "' and isactive='TRUE') "
	else
	    filter = " AND W.FCBA ='" & fcba & "' "
	end if

	if trim(agreementcode) <> "" then 
	    filteragrementcode = " AND W.AGREEMENTCODE ='" & agreementcode & "' "
	end if
	if trim(contractorcode) <> "" then 
	    filtercontractor = " AND C.FCCODE ='" & contractorcode & "' "
	end if


	' query alokasi wip
	qq = " Select B.FCCOMPANYCODE, W.FCBA, W.AGREEMENTCODE,C.FCCODE AS CODECONTRACTOR, C.FCNAME AS NAMACONTRACTOR,A.FEECONTRACTOR, " & _ 
	" W.JOBCODE,J.FCNAME AS COANAME,J.JOB_FORFIELDSTATUS AS JOBSTATUS, " & _ 
	" W.LOCATIONTYPE, W.LOCATIONCODE,SUM(W.QTY_FINAL)QTY_FINAL, SUM(W.OUTPUT_FINAL)OUTPUT_FINAL, SUM(W.UNIT_FINAL)UNIT_FINAL, W.RATE_FINAL, " & _ 
	" SUM(W.QTY_FINAL) * W.RATE_FINAL AS AMOUNT " & _ 
	" From  WORKINPROGRESSFINAL_NEW W " & _ 
	" LEFT JOIN CONTRACTAGREEMENT A ON W.AGREEMENTCODE=A.AGREEMENTCODE AND W.FCBA=A.FCBA " & _ 
	" LEFT JOIN JOBMASTER J ON W.JOBCODE=J.FCCODE " & _ 
	" LEFT JOIN CONTRACTOR C ON C.FCCODE=A.CONTRACTORCODE AND C.FCBA=A.FCBA " & _ 
	" LEFT JOIN BUSINESSUNIT B ON B.FCCODE=W.FCBA " & _ 
	" WHERE W.YEAR='" & txtyear & "' " & _ 
	" AND W.MONTH='" & txtmonth & "' AND W.ISPROCESSING='TRUE' " & _ 
	filter  & _ 
	filteragrementcode  & _ 
	filtercontractor  & _ 
	" GROUP BY B.FCCOMPANYCODE, W.FCBA, W.AGREEMENTCODE,C.FCCODE, C.FCNAME, " & _ 
	" W.JOBCODE,J.FCNAME,J.JOB_FORFIELDSTATUS, " & _ 
	" W.LOCATIONTYPE, W.LOCATIONCODE,W.RATE_FINAL,A.FEECONTRACTOR" 
	'response.write qq
	Set rsdata = Server.CreateObject("ADODB.Recordset")
	rsdata.CursorType = 3
	rsdata.CursorLocation = 3
	rsdata.Open qq, connstr
%>

<body>
    <div id="wrapper">
        <!-- Header -->
       <div class="export-btn">
        
			 <input class="tableToExcel" onClick="tableToExcel()" value="Export to Excel" type="button">
        </div>

        <!-- Table -->
        <table id="producttable" class="product-table">
            <thead>
				<tr>
					<th colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none">
						 <div id="header-wrapper">
							<h4><%=CompanyName%></h4>
							<h5>Estate: <%=fcbajudul%> | Periode: <%=textMonth(txtmonth) & "-" & txtyear%></h5>
							<h5>Laporan Bulanan Work in Progress (Final)</h5>
						</div>
						
					</th>
				</tr>
                <tr>
                    <th rowspan="2">No</th>
                    <th rowspan="2">Company</th>
                    <th rowspan="2">FCBA</th>
                    <th rowspan="2">Agreement Code</th>
                    <th rowspan="2">Contractor Code</th>
                    <th rowspan="2">Contractor Name</th>
                    <th rowspan="2">Activity Code</th>
                    <th rowspan="2">Activity Name</th>
                    <th rowspan="2">Status</th>
                    <th rowspan="2">Target Code</th>
                    <th rowspan="2">Rate</th>
                  
                    <th colspan="3">Realisasi</th>
                    <th rowspan="2">Deskripsi</th>
                </tr>
                <tr>
                    <th style="background-color: #2c3e50;color: #fff;">Output (Ha)</th>
                    <th style="background-color: #2c3e50;color: #fff;">Unit (Qty)</th>
                    <th style="background-color: #2c3e50;color: #fff;">Amount (Rp)</th>
                </tr>
            </thead>
            <tbody>
                <!-- contoh dummy row -->
				 <%
				 	dim totaloutput,totalunit,totalamount,feecontractor,qtyfinal,outputfinal,feecontractorname
					totaloutput = 0
					totalunit   = 0  
					totalamount = 0
					if Not rsdata.eof then
					no=1
					While Not rsdata.EOF

                    feecontractor = rsdata("FEECONTRACTOR")
                    if feecontractor = "1" then 
                        qtyfinal = 0 
                        outputfinal = 0
                        feecontractorname = "FEE CONTRACTOR"
                    Else 
                        qtyfinal = rsdata("OUTPUT_FINAL")
                        outputfinal = rsdata("QTY_FINAL")
                        feecontractorname = ""
                    end if

				%>
                <tr>
                    <td style="text-align:center;"><%=no%></td>
                    <td style="text-align:left;width:150px"><%=rsdata("FCCOMPANYCODE")%></td>
                    <td style="text-align:center;"><%=rsdata("FCBA")%></td>
                    <td style="text-align:left;"><%=rsdata("agreementcode")%></td>
                    <td style="text-align:left;"><%=rsdata("CODECONTRACTOR")%></td>
                    <td style="text-align:left;"><%=rsdata("NAMACONTRACTOR")%></td>
                    <td style="text-align:center;"><%=rsdata("JOBCODE")%></td>
                    <td style="text-align:left;"><%=rsdata("COANAME")%></td>
                    <td style="text-align:center;"><%=rsdata("JOBSTATUS")%></td>
                    <td style="text-align:center;"><%=rsdata("LOCATIONCODE")%></td>
                    <td style="text-align:right;"><%=FormatNumber(rsdata("RATE_FINAL"),0)%></td>
                  
                    <td class="ratakanan"><%=FormatNumber(outputfinal,2)%></td>
                    <td class="ratakanan"><%=FormatNumber(qtyfinal,2)%></td>
                    <td class="ratakanan"><%=FormatNumber(rsdata("amount"),0)%></td>
                    <td style="text-align:left;"><%=feecontractorname%></td>
                </tr>
				<%
					totaloutput = cdbl(totaloutput) + cdbl(rsdata("OUTPUT_FINAL"))
					totalunit   = cdbl(totalunit) + cdbl(rsdata("QTY_FINAL"))
					totalamount = cdbl(totalamount) + cdbl(rsdata("amount"))

					rsdata.MoveNext
					no=no+1
					Wend
					rsdata.close()
					rsdata=null
					END IF
				%>
            </tbody>
            <tfoot>
                <tr bgcolor="#F4F4A0" class="tebalhuruf">
                    <td colspan="11" align="right">Total</td>
                    <td><%=FormatNumber(totaloutput,2)%></td>
                    <td><%=FormatNumber(totalunit,2)%></td>
                    <td><%=FormatNumber(totalamount,0)%></td>
                </tr>
            </tfoot>
        </table>

        <!-- Footer -->
        <!--#include file="Bottom_Report.asp"-->
    </div>
</body>
</html>
