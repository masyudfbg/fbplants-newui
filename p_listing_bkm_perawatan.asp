<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,gangcode
Dim tbl,doc,v_txtFDDATE_From,v_txtFDDATE_To
Dim  MgrName,KTUName,CropName,FcbaType,txtSQL,txtSelect
Dim cmdstyle,cmdstyleTop,cmdstyleLeft
dim v_ALLOWANCE_AMT,v_bunch_weight,v_kg_brondolan
dim nCountofField,no,v_amount_allowded,v_jumlah_janjang
			
businessunit= trim(session("fcba"))
gangcode	= trim(request("gangcode"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtFDDATE_From=trim(REQUEST("bkmfrom"))
v_txtFDDATE_To	=trim(REQUEST("bkmto"))
doc				=trim(REQUEST("doc"))

businessunit= trim(session("fcba"))
'rowstate= trim(request("rowstate"))

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

	txtSQL = "Select H.BKMNO, TO_char(H.BKMDATE,'DD/MM/YYYY') BKMDATE, H.SECTIONCODE, H.GANGCODE, H.SUPERVISI1, " & _
	" H.SUPERVISI2, H.SUPERVISI3, H.SUPERVISI4,H.VEHICLECODE,D.EMPLOYEECODE, " & _
	" E.FCNAME || '-' || D.EMPLOYEECODE EMPLOYEENAME, D.ATTENDANCE,D.ACTCODE, C.FCNAME || '-' || D.ACTCODE ACTNAME, " & _
	" D.LOCATIONCODE, NVL(D.MANDAYS,0)MANDAYS, NVL(D.RATE,0)RATE, NVL(D.UNIT,0)UNIT, NVL(D.OUTPUT,0)OUTPUT, NVL(D.ABW,0)ABW, " & _
	" NVL(D.BUNCH_WEIGHT,0)BUNCH_WEIGHT, NVL(D.KG_BRONDOLAN,0) KG_BRONDOLAN, NVL(D.OVERTIME_HOURS,0) OVERTIME_HOURS, D.TYPE_OVERTIME, NVL(D.OTHCONVERTION,0)OTHCONVERTION," & _
	" D.ALLOWANCE_CODE, NVL(D.ALLOWANCE_AMT,0)ALLOWANCE_AMT,H.FCBA,D.CHARGEUNIT,D.CHARGEJOB,D.CHARGECODE, " & _
	" NVL(to_char(D.IN_FIN,'HH24:MI:SS'),'-')TIMEIN,NVL(to_char(D.OUT_FIN,'HH24:MI:SS'),'-')TIMEOUT,H.ROWSTATE, H.POSTINGFLG,D.REMARKS " & _
	" From   BKMHDR H " & _
	" LEFT JOIN BKMUPAHDTL D ON (H.BKMNO=D.BKMNO AND H.FCBA=D.FCBA) " & _
	" LEFT JOIN EMPLOYEE E ON (D.EMPLOYEECODE=E.FCCODE AND D.FCBA=E.FCBA) " & _
	" LEFT JOIN ACTMASTER C ON (C.FCCODE=D.ACTCODE) " & _
	" WHERE (H.FCBA='" & businessunit & "' or D.CHARGEUNIT='" & businessunit & "')  AND document_classification LIKE '" & doc & "%'"


if  len(trim(v_txtFDDATE_From)) > 0 then
    txtSQL = txtSQL & " and  H.BKMDATE >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
	txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
end if
if  len(trim(v_txtFDDATE_To)) > 0 then
    txtSQL =txtSQL &  " and H.BKMDATE <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
end if

txtsql	=	txtsql & " order by H.FCBA, H.BKMDATE"



'response.Write(txtsql)
'response.Flush()
%>

<html>
	<head>
		<title>FBPlants - BKM PERAWATAN</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
		<!-- Tampilan baru: tabel ala DataTables, kartu ringkasan (data & Export Excel tidak berubah) -->
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
		<link rel="stylesheet" href="/css/fbp-bkmlist.css?v=20261013b" type="text/css" />
		<script src="/js/fbp-table.js?v=20261009c" type="text/javascript"></script>
	</head>
<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: 15px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 45px;
	}
	.lebarlokasi{
		width: 50px;
	}
	.lebartc{
		width: 60px;
	}
	.lebarovtime{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		width: 110px;
	}
	.lebarallow{
		width: 90px;
	}
	.lebarupah{
		width: 120px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
	}
	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	
</style>



<body>

<br>

<div id="wrapper">
        <!--#include file="print_excel_blob.asp"-->
	<table border="0" width="100%" id="product-table" class=''>
			<thead>
				<tr>
					
					<th colspan="" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=COmpanyName%></h4>
								<h5><%=fcbaname%></h5>
							</div>
							
						</div>
					</th>
				</tr>
				<tr>
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">BKM PERAWATAN DAN BAHAN</th>
				</tr>
				<tr>
						<th align="left" colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
			<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH PERAWATAN</li>
						<li class="TabbedPanelsTab">BKM BAHAN</li>
					</ul>

					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1200"  id="bkm-upah" class="xthf-site bkm-list">
								<thead>
								<tr id="trheader2">
									<th style="color:#fff;">No</th>
									<th style="color:#fff;">No.BKM</th>
									<th style="color:#fff;">Fcba</th>
									<th style="color:#fff;">Charge Unit</th>
									<th style="color:#fff;">Date</th>
									<th style="color:#fff;">Gang Code</th> 
									<th style="color:#fff;">Supervisi 1</th> 
									<th style="color:#fff;">Supervisi 2</th> 
									<th style="color:#fff;">Supervisi 3</th> 
									<th style="color:#fff;">Supervisi 4</th> 
									<th style="color:#fff;">Section Name</th> 
									<th style="color:#fff;">Employee Name</th> 
									<th style="color:#fff;">Attendance</th> 
									<th style="color:#fff;">Job Code</th> 
									<th style="color:#fff;">Target Code</th> 
									<th style="color:#fff;">Charge Job</th> 
									<th style="color:#fff;">Charge Code</th> 
									<th style="color:#fff;">Allowded</th> 
									<th style="color:#fff;">Amount</th> 
									<th style="color:#fff;">Mandays</th>
									<th style="color:#fff;">Over Times</th>
									<th style="color:#fff;">Conv. Times</th>
									<th style="color:#fff;">Unit</th> 
									<th style="color:#fff;">Bunch Weight</th> 
									<th style="color:#fff;">KG Brondolan</th>
									<th style="color:#fff;">Output</th>
									<th style="color:#fff;">Rate</th>
									<th style="color:#fff;">Posting</th>
									<th style="color:#fff;">Finger Time</th>
									<th style="color:#fff;">Remarks</th>
									
								</tr>
								</thead>
								<tbody>
								<%
									Set rsCat = Server.CreateObject("ADODB.Recordset")
										rsCat.CursorType = 3
										rsCat.CursorLocation = 3
										rsCat.Open txtSql, connstr
								if Not rsCat.eof then
										nCountofField=rsCat.fields.count-1
										no=1
									
									Dim v_mandays,v_unit,v_Rate,v_Output,v_othrs, v_bucket,v_othconvertion,posting
										v_mandays=0
										v_unit=0
										v_Rate=0
										v_Output=0
										v_othrs=0
										v_othconvertion=0
										v_bucket=0
										v_amount_allowded= 0
										v_kg_brondolan = 0
										v_jumlah_janjang = 0

									While Not rsCat.EOF

										posting = rscat("POSTINGFLG")
										if  posting = "0" then 
											posting = "TIDAK"
										ELSE 
											posting = "YA"

										end if


										%>
								<tr id='tru'>
												
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMNO")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCBA")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGEUNIT")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMDATE")%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("GANGCODE")%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI1")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI2")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI3")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI4")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SECTIONCODE")%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EMPLOYEENAME")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ATTENDANCE")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ACTNAME")%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGEJOB")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGECODE")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ALLOWANCE_CODE")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("ALLOWANCE_AMT"),2)%></td>	
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("MANDAYS"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OVERTIME_HOURS"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OTHCONVERTION"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNIT"), 2)%></td> 
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BUNCH_WEIGHT")%></td> 
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_BRONDOLAN")%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OUTPUT"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("RATE"), 2)%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=posting%>/<%=rscat("ROWSTATE")%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIMEIN") & "-" & rscat("TIMEOUT")%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%></td>
								</tr>
								  <%
										
										v_ALLOWANCE_AMT	= v_ALLOWANCE_AMT + Cdbl(rsCat("ALLOWANCE_AMT"))
										v_mandays	= v_mandays + Cdbl(rsCat("mandays"))
										v_othrs	= v_othrs + Cdbl(rsCat("OVERTIME_HOURS"))
										v_othconvertion	= v_othconvertion + Cdbl(rsCat("OTHCONVERTION"))
										v_unit	= v_unit + Cdbl(rsCat("unit"))
										v_bunch_weight	= v_bunch_weight + Cdbl(rsCat("bunch_weight"))
										v_Output	= v_Output + Cdbl(rsCat("Output"))
										v_rate	= v_rate + Cdbl(rsCat("rate"))
										v_kg_brondolan	= v_kg_brondolan + Cdbl(rsCat("kg_brondolan"))
										
										' v_bucket	= v_bucket + Cdbl(to_nol(rsCat("bucket")))
										rsCat.MoveNext
										'response.Flush()
										no=no+1
										Wend
										rsCat.close()
										rsCat=null
										END IF
									%>
									</tbody>
									<tfoot>
									<tr bgcolor="#F4F4A0">
										<td align="Center" style="<%=cmdstyle%>" colspan= "18"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_ALLOWANCE_AMT, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_mandays, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_othrs, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_othconvertion, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_unit, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_bunch_weight, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_kg_brondolan, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Output, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_rate, 2)%></strong></td>
										<td colspan="3"></td>
									</tr>
									</tfoot>
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0"  width="1050" class='xthf-site bkm-list' id='bkm-bahan'>
									<thead>
									<tr id="trheader3">
										<th>No BKM</th>
										<th>Date</th>
										<th>Activity</th>
										<th>TC</th>
										<th>Item Bahan</th>
										<th>Satuan</th>
										<th>Qty</th>
									
									</tr>
									</thead>
									<tbody>


									<%
									
										
										dim sqlbarang,rsbarang 
										sqlbarang = "select h.bkmdate,b.BKMNO,b.ACTCODE,b.ACTCODE || '-' ||  a.fcname as activityname,b.ITEMCODE,s.FCCODE || '-' || s.FCNAME ITEMBARANG, " & _
										" s.UNITOFMEASUREMENT,b.LOCATIONCODE,NVL(b.qty,0)qty " & _
										" from BKMBAHANDTL b " & _
										" LEFT JOIN BKMHDR h on b.bkmno=h.bkmno and b.fcba=h.fcba" & _
										" LEFT join STOCK s on (b.ITEMCODE=s.FCCODE AND b.fcba=s.fcba) " & _
										" LEFT join ACTMASTER a on (b.ACTCODE=a.FCCODE)" & _ 
										" WHERE h.FCBA='" & businessunit & "'"
										 

										if  len(trim(v_txtFDDATE_From)) > 0 then
											sqlbarang = sqlbarang & " AND  h.BKMDATE >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
											txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
										end if
										if  len(trim(v_txtFDDATE_To)) > 0 then
											sqlbarang =sqlbarang &  " AND h.BKMDATE <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
											txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
										end if

										sqlbarang	=	sqlbarang & " order by h.BKMDATE"

										'response.write sqlbarang
										set rsbarang = connStr.Execute(sqlbarang)
										
									%>

									<% 
										dim novh,ttlqty
										novh = 0
										ttlqty = 0
									
										While Not rsbarang.EOF 
									%>
										<tr id='trv<%=novh%>'>
											<td class="ratakiri">
												<%=rsbarang("bkmno")%>
											</td>
											<td class="ratakiri">
												<%=rsbarang("bkmdate")%>
											</td>
											<td class="ratakiri">
												<%=rsbarang("activityname")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("LOCATIONCODE")%>
											</td>
											<td class="ratakiri">
												<%=rsbarang("ITEMCODE")%>&nbsp;<%=rsbarang("ITEMBARANG")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("UNITOFMEASUREMENT")%>
											</td>
											<td class="ratatengah">
												<%=FormatNumber(rsbarang("qty"))%>
											</td>

										</tr>

											
										
									<%
										ttlqty = ttlqty + cdbl(rsbarang("qty"))
										nobr = nobr + 1
										rsbarang.MoveNext
										Wend
									
									%>
									</tbody>
									<tfoot>
									<tr bgcolor="#F4F4A0">
										<td align="Center" style="<%=cmdstyle%>" colspan= "6"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlqty, 2)%></strong></td>
									
									</tr>
									</tfoot>
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>

					</div>
				</div>
			</td>
		</tr>

	</table>

</form>
</body>

</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">
	//window.onload = getcountEmployee();
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});
</script>
<script src="/js/fbp-bkmlist.js?v=20261013b" type="text/javascript"></script>
<script type="text/javascript">



</script>