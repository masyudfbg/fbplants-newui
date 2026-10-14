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
dim jjgun,jjgba
			
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
	" NVL(to_char(D.IN_FIN,'HH:MM:SS'),'-')TIMEIN,NVL(to_char(D.OUT_FIN,'HH:MM:SS'),'-')TIMEOUT,H.ROWSTATE, H.POSTINGFLG,D.REMARKS, " & _
	" NVL(D.bunchun,0)bunchun,NVL(D.bunchba,0)bunchba,NVL(D.abw,0)abw,nvl(D.kg_brondolan,0)kg_brondolan" & _
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
		<title>FBPlants - BKM PANEN</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
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

	.warnarowbg{
		background-color: #ffffe0;
	}
	.warnacell{
		background-color: #f8f811;
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
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">BKM PANEN DAN MUTU</th>
				</tr>
				<tr>
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
			<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM PANEN</li>
						<li class="TabbedPanelsTab">BKM MUTU</li>
						<li class="TabbedPanelsTab">BKM DENDA</li>
					</ul>

					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1200"  id="product-table" class="xthf-site">
								
								<tr id="trheader2">
									<th style="color:#fff;">No</th>
									<th style="color:#fff;">No.BKM</th>
									<th style="color:#fff;">Fcba</th>
									<th style="color:#fff;">Charge Unit</th>
									<th style="color:#fff;">Date</th>
									<th style="color:#fff;">Finger Time</th>
									<th style="color:#fff;">Gang Code</th> 
									<th style="color:#fff;">Supervisi 1</th> 
									<th style="color:#fff;">Supervisi 2</th> 
									<th style="color:#fff;">Supervisi 3</th> 
									<th style="color:#fff;">Supervisi 4</th> 
									<th style="color:#fff;">Section Name</th> 
									<th style="color:#fff;">Employee</th>
									
									<th style="color:#fff;">Lokasi</th>
									<th style="color:#fff;">Charge Job</th>
									<th style="color:#fff;">Charge Code</th>
									<th style="color:#fff;">Absen</th>
									<th style="color:#fff;">Activity</th>
									<th style="color:#fff;">TC</th>
									<th style="color:#fff;">HK</th>
									
									<th style="color:#fff;">JJG</th>
									<th style="color:#fff;">JJG UN</th>
									<th style="color:#fff;">JJG BA</th>
									<th style="color:#fff;">BJR</th>
									<th style="color:#fff;">Brdln</th>
									<th style="color:#fff;">KG TBS</th>
									<th style="color:#fff;">Output</th>
									<th style="color:#fff;">Rate</th>
									<th style="color:#fff;">Allowed Code</th>
									<th style="color:#fff;">Rate Allowed</th>
									<th style="color:#fff;">Remark</th>
									<th style="color:#fff;">Posting</th>
									
									
								</tr>
								<thead>
								<tbody>
								<%
									Set rsCat = Server.CreateObject("ADODB.Recordset")
										rsCat.CursorType = 3
										rsCat.CursorLocation = 3
										rsCat.Open txtSql, connstr
								if Not rsCat.eof then
										nCountofField=rsCat.fields.count-1
										no=1
									
									Dim v_mandays,v_unit,v_Rate,v_Output,v_othrs, v_bucket,v_abw,posting,rowClass,cellUnitClass
										v_mandays=0
										v_unit=0
										v_Rate=0
										v_Output=0
										v_othrs=0
										v_abw=0
										v_bucket=0
										v_amount_allowded= 0
										v_kg_brondolan = 0
										v_jumlah_janjang = 0
										jjgun = 0
										jjgba = 0
									While Not rsCat.EOF

										posting = rscat("POSTINGFLG")
										if  posting = "0" then 
											posting = "TIDAK"
										ELSE 
											posting = "YA"

										end if

										if cdbl(rscat("unit")) >= 180 AND rscat("ATTENDANCE")="KJ" AND rscat("ACTCODE")="6502101" then 
											rowClass = "warnarowbg"
											cellUnitClass = "warnacell"
											'response.write "tes 80"
										else
											rowClass = ""
											cellUnitClass = ""
										end if


										%>
								<tr class="" id='tru'>
												
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMNO")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCBA")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGEUNIT")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMDATE")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIMEIN") & "-" & rscat("TIMEOUT")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("GANGCODE")%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI1")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI2")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI3")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPERVISI4")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SECTIONCODE")%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EMPLOYEENAME")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGEUNIT")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGEJOB")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGECODE")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ATTENDANCE")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ACTNAME")%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%></td>		
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("MANDAYS"), 2)%></td>
									<td align="Right" class="" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNIT"), 2)%></td> 
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bunchun")%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bunchba")%></td>
												<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("abw")%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_BRONDOLAN")%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BUNCH_WEIGHT")%></td> 
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OUTPUT"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("RATE"), 2)%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ALLOWANCE_CODE")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("ALLOWANCE_AMT"),2)%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%></td>
								
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=posting%>/<%=rscat("ROWSTATE")%></td>
									
					
						
								</tr>
								  <%
										
										jjgun	= jjgun + Cdbl(rsCat("bunchun"))
										v_mandays	= v_mandays + Cdbl(rsCat("mandays"))
										jjgba	= jjgba + Cdbl(rsCat("bunchba"))
										v_abw	= v_abw + Cdbl(rsCat("abw"))
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
									<tr bgcolor="#F4F4A0">
										<td align="Center" style="<%=cmdstyle%>" colspan= "19"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_mandays, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_unit, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(jjgun, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(jjgba, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_abw, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_kg_brondolan, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_bunch_weight, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Output, 2)%></strong></td>
										
										<td></td>
									</tr>
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0"  width="1050" class='xthf-site' id='table3'>
									<tr id="trheader3">
										<th>Date</th>
										<th>Blok</th>
										<th>B. Mateng</th>
										<th>B. LMateng</th>
										<th>B. KMateng</th>
										<th>B. Mentah</th>
										<th>B. Busuk</th>
									</tr>


									<%
									
										
										dim sqlbarang,rsbarang 
										sqlbarang = "select * " & _
										" from BKMMUTU U " & _
										" LEFT JOIN BKMHDR H ON H.BKMNO=U.BKMNO AND H.FCBA=U.FCBA" & _
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
											<td class="ratatengah">
												
												<%=rsbarang("blok")%>
										
											</td>
											<td class="ratatengah">
												<%=rsbarang("bmateng")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("blmateng")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("bkmateng")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("bmentah")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("bbusuk")%>
											</td>

										</tr>

											
										
									<%
										
										nobr = nobr + 1
										rsbarang.MoveNext
										Wend
									
									%>
									</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>

						<!-- view bkm denda-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="detaildenda">
								<table align=center border="0" width="950" class='data' id='table3'>
									<tr >
										
											<th rowspan="2">Date</th>
											<th rowspan="2">Employee</th>
											<th colspan="7">Kualitas TBS</th>
											<th colspan="7">Kualitas Ancak</th>
											
									</tr>
									<tr id="trheader3">
										
										<th>(A)</th>
										<th>(BA)</th>
										<th>(TP)</th>
										<th>(BK)</th>
										<th>(O)</th>
										<th>(AB)</th>
										<th>(E)</th>

										<th>(BT)</th>
										<th>(PS)</th>
										<th>(SP)</th>
										<th>(BRD)</th>
										<th>(OP)</th>
										<th>(BM)</th>
										<th>(BDT)</th>
									</tr>

									<%
									
										dim sqldenda,rsdenda,nodenda 
										sqldenda = "Select H.BKMDATE, B.EMPLOYEECODE,E.FCNAME, B.LINENOKEY, B.REMARKS, " & _
										" B.TBSA, B.TBSBA, B.TBSTP, B.TBSBK, B.TBSO, " & _
										" B.TBSAB, B.TBSE, B.ANCAKBT, B.ANCAKPS, B.ANCAKSP, " & _
										" B.ANCAKBRD, B.ANCAKOP, B.ANCAKBM, B.ANCAKBDT " & _
										" From   BKMDENDA B " & _
										" LEFT JOIN BKMHDR H ON H.BKMNO=B.BKMNO AND H.FCBA=B.FCBA" & _
										" LEFT JOIN EMPLOYEE E ON (B.EMPLOYEECODE=E.FCCODE AND B.FCBA=E.FCBA) " & _
										" WHERE H.FCBA='" & businessunit & "'"

										if  len(trim(v_txtFDDATE_From)) > 0 then
											sqldenda = sqldenda & " AND  H.BKMDATE >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
											txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
										end if
										if  len(trim(v_txtFDDATE_To)) > 0 then
											sqldenda =sqldenda &  " AND H.BKMDATE <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
											txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
										end if

										 set rsdenda = connStr.Execute(sqldenda)
										 'response.write sqlbarang
									%>

									<% 
										 
										nodenda = 0
										While Not rsdenda.EOF 
									%>
										<tr id='trd<%=nodenda%>'>
											<td class="ratakiri" >
												<%=rsdenda("BKMDATE")%>
											</td>
											<td class="ratakiri">
													<%=rsdenda("EMPLOYEECODE")%>><%=rsdenda("FCNAME")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSA")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSBA")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSTP")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSBK")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSO")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSAB")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("TBSE")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKBT")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKPS")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKSP")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKBRD")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKOP")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKBM")%>
											</td>
											<td class="ratatengah">
												<%=rsdenda("ANCAKBDT")%>
											</td>
										</tr>

											
										
									<%
										nodenda = nodenda + 1
										rsdenda.MoveNext
										Wend
									
									%>
									
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