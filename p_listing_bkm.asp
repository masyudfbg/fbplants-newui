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
	" D.CHARGEUNIT,D.CHARGEJOB,D.CHARGECODE,D.ALLOWANCE_CODE, NVL(D.ALLOWANCE_AMT,0)ALLOWANCE_AMT, " & _
	" NVL(to_char(D.IN_FIN,'HH:MM:SS'),'-')TIMEIN,NVL(to_char(D.OUT_FIN,'HH:MM:SS'),'-')TIMEOUT,H.ROWSTATE, H.POSTINGFLG,D.REMARKS,NVL(p.AMOUNT,0)AMOUNT " & _
	" From   BKMHDR H " & _
	" LEFT JOIN BKMUPAHDTL D ON (H.BKMNO=D.BKMNO AND H.FCBA=D.FCBA) " & _
	" LEFT JOIN EMPLOYEE E ON (D.EMPLOYEECODE=E.FCCODE AND D.FCBA=E.FCBA) " & _
	" LEFT JOIN BKMPROSES  p ON (D.EMPLOYEECODE=p.EMPLOYEECODE AND p.BKMNO = D.BKMNO AND p.LINENOKEY = D.LINENOKEY AND D.fcba=p.fcba) " & _ 
	" LEFT JOIN ACTMASTER C ON (C.FCCODE=D.ACTCODE) " & _
	" LEFT JOIN ALLOWDED al ON (al.FCCODE = p.ALLOWDEDCODE and al.fcba=p.fcba) " & _ 
	" WHERE H.FCBA='" & businessunit & "' AND document_classification LIKE '" & doc & "%'"


if  len(trim(v_txtFDDATE_From)) > 0 then
    txtSQL = txtSQL & " and  H.BKMDATE >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
	txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
end if
if  len(trim(v_txtFDDATE_To)) > 0 then
    txtSQL =txtSQL &  " and H.BKMDATE <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
end if

txtsql	=	txtsql & " order by H.BKMDATE"



'response.Write(txtsql)
'response.Flush()
%>

<html>
	<head>
		<title>FBPlants - BKM TRAKSI</title>
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
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">BKM LISTING DAN VEHICLE/WORKSHOP ACTIVITY</th>
				</tr>
				<tr>
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
			<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM VEHICLE/WORKSHOP ACTIVITY</li>
					</ul>

					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1200"  id="product-table" class="xthf-site">
								
								<tr id="trheader2">
									<th style="color:#fff;">No</th>
									<th style="color:#fff;">No BKM</th>
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
									<!--<th style="color:#fff;">Allowded</th> 
									<th style="color:#fff;">Allowed Amount</th> -->
									<th style="color:#fff;">Mandays</th>
									<th style="color:#fff;">Over Times</th>
									<th style="color:#fff;">Conv. Times</th>
									<th style="color:#fff;">Unit</th> 
									<th style="color:#fff;">Bunch Weight</th> 
									<th style="color:#fff;">KG Brondolan</th>
									<th style="color:#fff;">Output</th>
									<th style="color:#fff;">Rate</th>
									<th style="color:#fff;">Amount</th>
									<th style="color:#fff;">Posting</th>
									<th style="color:#fff;">Finger Time</th>
									<th style="color:#fff;">Remarks</th>
									
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
									
									Dim v_mandays,v_unit,v_Rate,v_Output,v_othrs, v_bucket,v_othconvertion,posting,amtall
									Dim ALLOWANCE_CODE,ALLOWANCE_amt,mandays
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
										amtall = 0

									While Not rsCat.EOF

										posting = rscat("POSTINGFLG")
										ALLOWANCE_amt = rscat("ALLOWANCE_AMT")
										mandays = rscat("mandays")
										if  posting = "0" then 
											posting = "TIDAK"
										ELSE 
											posting = "YA"

										end if
										if cdbl(mandays)> 0 then 
											mandays = mandays
										ELSE 
											mandays = 0
										end if

										%>
								<tr id='tru'>
												
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMNO")%></td>			
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
								
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(mandays, 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OVERTIME_HOURS"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OTHCONVERTION"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNIT"), 2)%></td> 
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BUNCH_WEIGHT")%></td> 
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_BRONDOLAN")%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OUTPUT"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("RATE"), 2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("AMOUNT"), 2)%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=posting%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIMEIN") & "-" & rscat("TIMEOUT")%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%></td>
								</tr>
								  <%
										
									
										v_mandays	= v_mandays + Cdbl(rsCat("mandays"))
										v_othrs	= v_othrs + Cdbl(rsCat("OVERTIME_HOURS"))
										v_othconvertion	= v_othconvertion + Cdbl(rsCat("OTHCONVERTION"))
										v_unit	= v_unit + Cdbl(rsCat("unit"))
										v_bunch_weight	= v_bunch_weight + Cdbl(rsCat("bunch_weight"))
										v_Output	= v_Output + Cdbl(rsCat("Output"))
										v_rate	= v_rate + Cdbl(rsCat("rate"))
										v_kg_brondolan	= v_kg_brondolan + Cdbl(rsCat("kg_brondolan"))
										amtall	= amtall + Cdbl(rsCat("amount"))
										
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
										<td align="Center" style="<%=cmdstyle%>" colspan= "14"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
								
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_mandays, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_othrs, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_othconvertion, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_unit, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_bunch_weight, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_kg_brondolan, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Output, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_rate, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(amtall, 2)%></strong></td>
										<td></td>
									</tr>
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0"  width="1050" class='xthf-site' id='table3'>
									<tr id="trheader3">
										<th>No BKM</th>
										<th>Vehicle Code</th>
										<th>Date</th>
										<th>Opr/Driver</th>
										<th>Helper</th>
										<th>Lokasi</th>
										<th>Time</th>
										<th>Activity</th>
										<th>TC</th>
										<th>Category Transport</th>
										<th>Trip</th>
										<th>Hours</th>
										<th>KM/HM</th>
										<th>Output</th>
										<th>Remark</th>
										
									</tr>


									<%
									
										dim sqlvehicle,rsvehicle
										sqlvehicle = "Select h.bkmno, h.bkmdate, bv.EMPLOYEECODE, e.FCNAME || '-' || e.FCCODE EMPLOYEENAME, " & _ 
										" bv.HELPERCODE,(SELECT hp.FCNAME FROM EMPLOYEE hp WHERE bv.HELPERCODE=hp.FCCODE and hp.fcba=bv.fcba) AS HELPERNAME, " & _
										" bv.fcba, bv.VEHICLECODE,v.FCCODE || ' - ' || v.FCNAME VEHICLENAME, bv.ACTCODE,c.FCNAME || ' - ' || c.FCCODE COANAME, " & _ 
										" bv.LOCATIONCODE,bv.CATEGORYTRANSPORT,t.FCCODE || ' - ' || t.FCNAME TRANSPORTKATNAME, " & _
										" bv.TIME, nvl(bv.HOURS,0)HOURS, nvl(bv.UNIT,0) UNIT, NVL(bv.OUTPUT,0)OUTPUT, NVL(bv.TRIP,0)trip,bv.remarks,bv.chargeunit, h.vehiclecode kodekendaraan " & _
										" From   BKMVEHICLEDTL bv " & _
										" INNER JOIN BKMHDR h ON (h.bkmno=bv.bkmno and h.fcba=bv.fcba) " & _
										" INNER JOIN EMPLOYEE e on (e.FCCODE=bv.EMPLOYEECODE and e.fcba=bv.fcba)" & _
										" INNER JOIN ACTMASTER c on (bv.ACTCODE=c.FCCODE) " & _
										" LEFT JOIN VEHICLEACTIVITYCATEGORY t on (bv.CATEGORYTRANSPORT=t.FCCODE AND bv.fcba=t.fcba) " & _
										" INNER JOIN VEHICLE v on (v.FCCODE=bv.VEHICLECODE AND v.FCBA=bv.FCBA) " & _ 
										" WHERE h.fcba='" & businessunit & "' "

										if  len(trim(v_txtFDDATE_From)) > 0 then
											sqlvehicle = sqlvehicle & " AND  h.BKMDATE >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
											txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
										end if
										if  len(trim(v_txtFDDATE_To)) > 0 then
											sqlvehicle =sqlvehicle &  " AND h.BKMDATE <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
											txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
										end if

										sqlvehicle	=	sqlvehicle & " order by h.BKMDATE"

										'response.write sqlvehicle
										 set rsvehicle = connStr.Execute(sqlvehicle)
										
									%>

									<% 
										dim novh,ttltrip,ttlhours,ttlkm,ttloutput
										novh = 0
										ttltrip = 0
										ttlhours = 0
										ttlkm = 0
										ttloutput = 0
										While Not rsvehicle.EOF 
									%>
										<tr id='trv<%=novh%>'>
											<td class="ratakiri">
											<%=rsvehicle("bkmno")%>
											</td>
											<td class="ratatengah">
											<%=rsvehicle("kodekendaraan")%>
											</td>
											<td class="ratakiri">
											<%=rsvehicle("bkmdate")%>
											</td>
											<td class="ratakiri">
											<%=rsvehicle("EMPLOYEENAME")%>
											</td>
											<td class="ratakiri">
												<%=rsvehicle("HELPERNAME")%>
											</td>
											
											<td class="ratakiri">
												<%=rsvehicle("fcba")%>
											</td>
											<td class="ratatengah">
												<%=rsvehicle("TIME")%>
											</td>

											<td class="ratakiri">
												<%=rsvehicle("ACTCODE")%>&nbsp;<%=rsvehicle("COANAME")%>
											</td>

											<td class="ratatengah">
												<%=rsvehicle("LOCATIONCODE")%>
											</td>

											<td class="ratakiri">
												<%=rsvehicle("CATEGORYTRANSPORT")%>&nbsp;<%=rsvehicle("TRANSPORTKATNAME")%>
												
											</td>

											<td class="ratatengah">
												<%=rsvehicle("trip")%>
											</td>

											<td class="ratatengah">
												<%=rsvehicle("HOURS")%>
											</td>


											<td class="ratatengah">
												<%=rsvehicle("UNIT")%>
											</td>

											<td class="ratatengah">
												<%=rsvehicle("OUTPUT")%>
											</td>

											<td class="ratakiri">
												<%=rsvehicle("remarks")%>
											</td>

										</tr>

											
										
									<%
										ttltrip = cdbl(ttltrip) + cdbl(rsvehicle("trip"))
										ttlhours =  cdbl(ttlhours) + cdbl(rsvehicle("HOURS"))
										ttlkm =  cdbl(ttlkm) + cdbl(rsvehicle("unit"))
										ttloutput =  cdbl(ttloutput) + cdbl(rsvehicle("output"))
										novh = novh + 1
										rsvehicle.MoveNext
										Wend
									
									%>
									</tr>
									<tr bgcolor="#F4F4A0">
										<td align="Center" style="<%=cmdstyle%>" colspan= "10"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttltrip, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlhours, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlkm, 2)%></strong></td> 
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttloutput, 2)%></strong></td> 
										
										<td></td>
									</tr>
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