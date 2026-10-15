<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim ScriptName,nSpanTotal,v_txttodaydate,PRODUCTCODE,v_txtDivision,v_txtBlok,v_txtyear,v_txtmonth,txtyear,txtmonth,txttodaydate,txtDivision, txtBlok, formatdatevalue0, formatdatevalue
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
Dim rsST, strST, rsProductCode, strProductCode, rsBlok,rsDivision,no
Dim rsHdr, StrHdr, TotalCost, TotalHa, CostPerHa, TotalWorkDay,TotalWorkDayShi , CostPerHaPerHK, CostPerHaHkPeriod
Dim rsDtl, StrDtl,CostPerBlock,RevenuePerBlock
Dim rsPL,  StrPL, arrayPL , arrayPL_eof, i , fromdate_ , todate_
Dim RsMW, StrMW, nMillWeight, nPrice 
Dim whereCons,orderCons
arrayPL_eof = True	

    Dim rsDefaultBa, strDefaultBa, varDefaultBa
	strDefaultBa = " SELECT" & _
					"     FCCODE as FCBA" & _
					" FROM BUSINESSUNIT BU" & _
					" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
					" WHERE" & _
					"     ISDEFAULT = 'TRUE'" & _
					"     AND FCTYPE = 'E'" & _
					"     AND FCUSERCODE = '" & Session("CodeName") & "'"

	Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
	rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

	If Not rsDefaultBa.EOF Then
		varDefaultBa = rsDefaultBa("FCBA")
	Else
		varDefaultBa = ""
	End If
	rsDefaultBa.Close
	set rsDefaultBa = Nothing
	
    v_txtFcba = Request("txtFcba")
	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If

If Len(Trim(Request("txttodaydate"))) = 0 Then
	v_txttodaydate		=	date
else
	v_txttodaydate		=	Request("txttodaydate")
End If

txtyear 			= year(v_txttodaydate)
txtmonth 			= month(v_txttodaydate)
v_txtyear			=	txtyear
v_txtmonth			=	txtmonth


v_txtDivision		= Request("txtDivision")
v_txtBlok			= Request("txtBlok")






ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
         q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtFcba & "'"
		 set rsCat = connStr.Execute(q)
		 if not rsCat.eof then
		    fcbaname = rsCat("fcName")
		    CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
		    KTUName=rsCat("FCKTU")
         end if
 formatdatevalue = lpad(Day(FormatDateTime(v_txttodaydate,2)),"0",2) & "-" & UCase(MonthName(Month(FormatDateTime(v_txttodaydate,2)),1)) & "-" & Year(FormatDateTime(v_txttodaydate,2))  
 formatdatevalue0 = "01-" & UCase(MonthName(Month(FormatDateTime(v_txttodaydate,2)),1)) & "-" & Year(FormatDateTime(v_txttodaydate,2))  
		 

 
  
	if v_txtDivision <>  "--ALL--" then
	   whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
	end if
	'selectCons = selectCons & " , DIVISION "
	'orderCons  = orderCons & ",DIVISION "

 
 
	if v_txtBlok <>  "--ALL--" then
	   whereCons = whereCons & " AND FCCODE = '" & v_txtBlok & "'"
	end if
	'selectCons = selectCons & " , FCCODE "
	'orderCons  = orderCons & ",FCCODE "

 
					 
%>
<html>
<head>
<title>ePlaS Reporting - Report Profit And Loss </title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>


<style type="text/css">
@media print
{
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>
</head>
<body onload="javascript:filtering()">
	<div id="wrapper">
      <!--#include file="../print_excel.asp"-->
	  <div id="header-wrapper">
	  <!--#include file="menutop.asp"-->
	  
	  <form action="<%=ScriptName%>" method="post" name="frmEntry">
	      <!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
			<table id="form-table" class="centered-margin">
				    <tr>
						<th>YEAR</th>
						<th>MONTH</th>
						<th>DATE</th>
						<th>AFDELING</th>
						<th>BLOK</th>
						<th>&nbsp;</th>
					</tr>
					<tr>
						<!--%response.write v_txtfcba %-->
						<!--%response.write v_txtyear %-->
						<!--%response.write v_txtmonth %-->
						<!--%response.write len(trim(request("txtYear"))) %-->
						<td align="right"><input type="text" name="txtyear" size="4" value="<%=v_txtyear%>" readonly="True"  style="background:#999999"></td>
						<td align="right"><input type="text" name="txtmonth" size="4" value="<%=v_txtmonth%>" readonly="True"  style="background:#999999"></td>
						<td align="center">
						<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
							<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
						<% else%>
							<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						
						<td align="center">
							<select name="txtDivision">	
								<option value="--ALL--">--ALL--</option>
								<%
									
									Set rsDivision = Server.CreateObject("ADODB.Recordset")
									rsDivision.CursorLocation = adUseClient
									

							
									if len(trim(v_txtDivision)) > 0 then 
										strsql =  " SELECT DISTINCT Division as Division FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba& "'"
									else
										strsql =  " SELECT DISTINCT Division as Division FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'"
									end if
									
									
									rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsDivision.eof then
										DO while not rsDivision.eof
										if v_txtDivision = rsDivision("Division") then
								%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
										<%else%>
											<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
										<%end if%>
								<%			rsDivision.movenext
										loop
									End if
									rsDivision.Close
									set rsDivision = nothing
								%>
							</select>		
						</td>
						
						<td align="center">
							<select name="txtBlok">	
								<option value="--ALL--">--ALL--</option>
								<%
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.CursorLocation = adUseClient

							
									if len(trim(v_txtBlok)) > 0 then 
										strsql =  " SELECT DISTINCT fccode as Blok FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "' Order by fccode"
									else
										strsql =  " SELECT DISTINCT fccode as Blok FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "' Order by fccode"
									end if
									
									rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsBlok.eof then
										DO while not rsBlok.eof
										if v_txtBlok = rsBlok("Blok") then
								%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
										<%else%>
											<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
										<%end if%>
								<%			rsBlok.movenext
										loop
									End if
									rsBlok.Close
									set rsBlok = nothing
								%>
							</select>
						 </td>	
						<td><input type="submit" value="View" /></td>
					</tr>				
			</table>
	  	 </form>
 
	<table width="100%" id="product-table" class='xthf-site'>
			<thead>
			<tr>
			    <th colspan="8" style="background:#FFFFFF; border-top:1px #ffffff solid; border-bottom:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				
				<div id="header-wrapper">
					<% if  CompanyName <>  "" then %>
						<h4>PT.<%=CompanyName%></h4> 
					<% End if%>
					<% If  FcbaName <> ""  then %>
						   <h5><%=FcbaName%></h5>   
					<% End if%>
				</div>  
				  
   				<div id="subheader" class="centered">
						<h2>PROFIT AND LOSS PER BLOCK</h3>
						<br/>
						<h4>PERIOD&nbsp;:&nbsp;<%=formatdatevalue0 %> to <%=formatdatevalue %></h4>
				</div>
			</tr>
			
			<% 
			
			'response.write " v_txttodaydate " & v_txttodaydate & Request("txttodaydate")  & formatdatevalue0
			StrHdr = " Select cost, totha, case totha when 0 then 0 else  cost / totha end  costperha , budget_transaction_workday_api.GetQtyWorkingDay_SHI('" & v_txtFcba & "'," & v_txtyear & ",1,1,TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')) WDshi, budget_transaction_workday_api.GetQtyWorkingDay_Month('" & v_txtFcba & "'," & v_txtyear & ",1,1,add_months(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),-1)) totWD" & _
					 " From ( select PNL_PER_BLOCK_API.GetTotalCost_PM('" & v_txtFcba & "', TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')) Cost," & _
					 "	    field_api.GetValid_HaEstate('" & v_txtFcba & "', add_months(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),-1)) TotHa" & _
					 "		  from dual) "
            Set rsHdr = Server.CreateObject("ADODB.Recordset")
			'response.write StrHdr
			rsHdr.Open StrHdr, connstr, adOpenStatic


			if Not rsHdr.eof then
				TotalCost = rsHdr("cost") 
				TotalHa   = rsHdr("totha")
				CostPerHa = rsHdr("costperha")
				TotalWorkDay = rsHdr("totWD")
				TotalWorkDayShi = rsHdr("WDshi")
			end if
			rsHdr.close()
			if cdbl(TotalWorkDay) > 0 then
			   CostPerHaPerHK = cdbl(CostPerHa) / cdbl(TotalWorkDay)
			else
			   CostPerHaPerHK = 0
            end if 		

        
			CostPerHaHkPeriod = cdbl(CostPerHaPerHK) * cdbl(TotalWorkDayShi)
						
			
			%>
<tr  style="background:#FFFFFF;border:0;">
<th colspan="8" border="0" style="background:#FFFFFF;border:0;">			
		<table width="60%" border="0" style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;" >
			<tbody>
			  <tr>
				<td  style="background-color:#00A300;color:#fff;border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;">Total Cost for Prev. Month</td>
				<td style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;" align="right"><%=formatnumber(TotalCost,2) %></td>
			  </tr>
			  <tr>
				<td style="background-color:#00A300;color:#fff;border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;">Total Ha Prev. Month</td>
				<td style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;"  align="right"><%=formatnumber(TotalHa,2) %></td>
			  </tr>
			  <tr>
				<td style="background-color:#00A300;color:#fff;border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;">Total Working Days for Prev. Month</td>
				<td style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;"  align="right"><%=formatnumber(TotalWorkDay,2) %></td>
			  </tr>
			  <tr>
				<td style="background-color:#00A300;color:#fff;border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;">Cost per Ha. for Prev. Month</td>
				<td style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;"  align="right"><%=formatnumber(CostPerHa,2) %></td>
			  </tr>
			  <tr>
				<td style="background-color:#00A300;color:#fff;border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;">Total Cost per Ha. / HK for Prev. Month</td>
				<td style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;"  align="right"><%=formatnumber(CostPerHaPerHK,2) %></td>
			  </tr>
			  <tr>
				<td style="background-color:#00A300;color:#fff;border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;">Total Working Days for this Period </td>
				<td style="border: 1px solid #3A5F0B;padding: 5px; font-size: 12px;"  align="right"><%=formatnumber(TotalWorkDayShi,2) %></td>
			  </tr>
			  </tbody>
		</table>	  
			  
</th>
<tr>	


  <tr>
    <th align="center" style="background-color:#00A300;color:#fff;">Block</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Ha</th>
	<th align="center" style="background-color:#00A300;color:#fff;">Planting Year</th>
    <th align="center" style="background-color:#00A300;color:#fff;">&nbsp;</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Period</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Mill Weight (Kg)</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Price ( Kg)</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Amount</th>
  </tr>
  </thead>
<tbody>
  <%
    Dim AdaRevenue, RevenueKe
	
	StrPL  = " Select p.productcode,p.plantingyear,p.price ,to_char(p.ondate,'DD-MON-YYYY') fromdate, " & _
			 " to_char(LEAD(p.ondate, 1) over(partition by p.plantingyear ORDER BY p.ondate) - 1, 'DD-MON-YYYY') todate " & _
			 " from PriceListTBS p " & _
			 " where p.fcba='" & v_txtFcba & "'" & _
			 " and to_char(p.ondate,'MMYYYY') = to_char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),'MMYYYY') " & _
			 " and trunc(p.ondate) <= trunc(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'))"
	
	'response.write StrPL
	Set rsPL = Server.CreateObject("ADODB.Recordset")
	rsPL.Open StrPL, connstr, adOpenStatic
	if not rsPL.EOF then
		arrayPL = rsPL.getrows()
		arrayPL_eof= false
	end if
	rsPL.Close
	set rsST = Nothing
	
 IF arrayPL_eof = false then
	
	
	StrDtl = " select f.fccode, field_api.getValid_HA(f.fccode, '" & v_txtFcba & "', TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')) Ha " & _
	         " , field_api.getValid_Plantingyear(f.fccode, '" & v_txtFcba & "', TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')) Plantyear " & _ 
			 " from field f " & _
			 " where f.fcba = '" & v_txtFcba & "' " & whereCons & " " & _
			 " and field_api.GetValid_Activation(f.fccode, f.fcba, TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')) = 'Y' " & _
			 " and field_api.GetValid_ValidFrom(f.fccode, f.fcba, TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')) <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
			 " order by f.fccode "
    
    Set rsDtl = Server.CreateObject("ADODB.Recordset")
	'response.write StrDtl
	rsDtl.Open StrDtl, connstr, adOpenStatic
    IF Not rsDtl.eof Then
		While Not rsDtl.eof	
		
		
		
		    RevenuePerBlock = 0
			RevenueKe = 0
			AdaRevenue = False
			'response.write "LB " & LBound(arrayPL, 2) & " , Ubound" & UBound(arrayPL, 2)
		If IsArray(arrayPL) Then	
			For i = LBound(arrayPL, 2) To UBound(arrayPL, 2)
			        fromdate_ = arrayPL(3, i) 
					todate_   = arrayPL(4, i)
					'response.write fromdate_ & " - " & todate_  & "***************"
                    IF 	IsNull(todate_)=true then
					    todate_ = formatdatevalue
                    End If	
                    					

                    nMillWeight = 0	
					StrMW = " Select nvl(Sum(h.mill_weight),0)  mill_weight from harvesting h" & _
							 " where h.fcba='" & v_txtFcba & "'  and h.fieldcode = '" & rsDtl("fccode") & "' " & _
							 " and field_api.getValid_Plantingyear(h.fieldcode, '" & v_txtFcba & "', TO_DATE ('" & fromdate_ & "', 'DD-MON-YYYY')) = '" & arrayPL(1, i) & "' "& _
							 " and trunc(h.receptiondate) between to_date('" & fromdate_ & "','DD-MON-YYYY')   and to_date( '" & todate_ & "','DD-MON-YYYY')"

					Set RsMW = Server.CreateObject("ADODB.Recordset")
					'response.write StrMW
					RsMW.Open StrMW, connstr, adOpenStatic
					if Not RsMW.eof then
						nMillWeight = RsMW("mill_weight") 		
					end if
					RsMW.close()	
					RevenuePerBlock = cdbl(RevenuePerBlock )+ cdbl(nMillWeight) * cdbl(arrayPL(2, i))
			       ' response.write " nMillWeight " & nMillWeight
			    If cdbl(nMillWeight) > 0 then 	
                   RevenueKe = cdbl(RevenueKe) + 1		
%>	 
					  <tr>
					     <% If cdbl(RevenueKe) = 1 then %>
								<td align="left" ><%=rsDtl("fccode")%></td>
								<td align="right"><%=formatnumber(rsDtl("Ha"),2) %></td>
								<td align="Center"><%=rsDtl("Plantyear") %></td>
								<td>Revenue</td>
						<%  Else   %>
						        <td colspan="4">&nbsp;</td> 
                        <%  End If %>								
							<td align="center"><%=fromdate_ %> - <%=todate_ %></td>
							<td align="right"><%=formatnumber(nMillWeight,2)%></td>
							<td align="right"><%=formatnumber(arrayPL(2, i),2)%></td>
						    <td align="right"><%=formatnumber(cdbl(nMillWeight) * cdbl(arrayPL(2, i)),2)%></td>
					  </tr>
<%              End If   
	        Next
		End if	

		If cdbl(RevenueKe) = 0 then 
%>
			<tr>
				<td align="left" ><%=rsDtl("fccode")%></td>
				<td align="right"><%=formatnumber(rsDtl("Ha"),2) %></td>
				<td align="Center"><%=rsDtl("Plantyear") %></td>
				<td>Revenue</td>							
				<td align="center"><%=formatdatevalue0 %> - <%=formatdatevalue %></td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
			</tr>
<%
        End if	
%>
		
   
              <tr>    
				<td colspan="7">&nbsp;</td>
				<td align="right"><%=formatnumber(RevenuePerBlock,2)%></td>
			  </tr>
			  <tr>
				<td colspan="8">&nbsp;</td>
			  </tr>
			  <tr>
				<td colspan="3">&nbsp;</td>
				<td>Cost</td>
				<td align="center"><%=formatdatevalue0 %> - <%=formatdatevalue %></td>
				<td colspan="2">&nbsp;</td>
				<% CostPerBlock = rsDtl("Ha")* cdbl(CostPerHaHkPeriod)   %>
				<td align="right"><%=formatnumber( CostPerBlock,2) %></td>
			  </tr>
			  <tr>
				<td colspan="8">&nbsp;</td>
			  </tr>
			  <tr>
				<td align="center" colspan="7">P/L Block Code ( <%=rsDtl("fccode")%> )</td>
				<td align="right"><%=formatnumber(cdbl(RevenuePerBlock) - cdbl(CostPerBlock),2)%></td>
			  </tr>
			 <tr>
				<td colspan="8">&nbsp;</td>
			 </tr>
	<%
	       rsDtl.MoveNext
		Wend
	End if
 Else	
    %>			
            <tr>
				<td align="center" colspan="7">NO DATA AVAILABLE ( Please Input TBS (FFB) Price  Period <%=formatdatevalue0 %> - <%=formatdatevalue %> ) </td>
			</tr>
<% End If %>  			
 </tbody>	
</table>
<!--#include file="bottom.Asp"-->