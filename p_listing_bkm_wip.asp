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

	txtSQL = "Select h.bkmno, TO_char(H.BKMDATE,'DD/MM/YYYY') BKMDATE, wi.CONTRACTORCODE,ct.FCCODE || ' - ' || ct. FCNAME CONTRACTORNAME,wi.AGREEMENTCODE,wi.ACTCODE,c.FCCODE || ' - ' || c.FCNAME COANAME, " & _
	" wi.LOCATIONCODE, wi.QTY, wi.OUTPUT, nvl(wi.UNIT,0)UNIT,wi.fcba,wi.CHARGEUNIT,wi.BKMAGREEMENTLINENO,wi.RATE " & _
	" From   BKMWIPDTL wi " & _
	" LEFT join bkmhdr H on H.BKMNO=wi.BKMNO AND H.FCBA=wi.FCBA " & _
	" LEFT JOIN JOBMASTER c on (wi.ACTCODE=c.FCCODE) " & _
	" LEFT JOIN CONTRACTOR  ct on (wi.CONTRACTORCODE=ct.FCCODE AND wi.FCBA=ct.FCBA)" & _
	" WHERE (wi.FCBA='" & businessunit & "' or wi.CHARGEUNIT='" & businessunit & "') AND document_classification LIKE '" & doc & "%'"


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
		<title>FBPlants - BKM WIP</title>
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
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">BKM LISTING WIP</th>
				</tr>
				<tr>
						<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
			<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM WIP</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1200"  id="product-table" class="xthf-site">
								
								<tr id="trheader2">
									<th style="color:#fff;">No</th>
									<th style="color:#fff;">Date</th>
									<th style="color:#fff;">No BKM</th>
									<th style="color:#fff;">Contractor Name</th> 
									<th style="color:#fff;">Fcba</th> 
									<th style="color:#fff;">Charge unit</th> 
									<th style="color:#fff;">SPK No</th> 
									<th style="color:#fff;">Activity</th> 
									<th style="color:#fff;">Line No</th> 
									<th style="color:#fff;">Rate</th> 
									<th style="color:#fff;">TC</th> 
									<th style="color:#fff;">Qty</th> 
									<th style="color:#fff;">Unit</th> 
									<th style="color:#fff;">Output</th> 
																	
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
									
									Dim vqty,vunit,voutput
										vqty=0
										vunit=0
										voutput=0
									

									While Not rsCat.EOF


										%>
								<tr id='tru'>
												
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMDATE")%></td>			
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMNO")%></td>			
									<td align="LEFT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CONTRACTORNAME")%></td>	
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCBA")%></td>				
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHARGEUNIT")%></td>				
									<td align="LEFT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AGREEMENTCODE")%></td>				
									<td align="LEFT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("COANAME")%></td>		
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BKMAGREEMENTLINENO")%></td>		
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("RATE"))%></td>
									<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%></td>				
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("QTY"))%></td>	
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("UNIT"))%></td>			
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("OUTPUT"))%></td>		

								</tr>
								  <%
										
										vqty	= vqty + Cdbl(rsCat("QTY"))
										vunit	= vunit + Cdbl(rsCat("unit"))
										voutput	= voutput + Cdbl(rsCat("output"))
										
										rsCat.MoveNext
										'response.Flush()
										no=no+1
										Wend
										rsCat.close()
										rsCat=null
										END IF
									%>
									<tr bgcolor="#F4F4A0">
										<td align="Center" style="<%=cmdstyle%>" colspan= "11"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(vqty, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(vunit, 2)%></strong></td>
										<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(voutput, 2)%></strong></td> 
								
									</tr>
							</table>
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