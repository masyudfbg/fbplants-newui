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
<title>ePlaS Reporting - Report Estimate Profit And Loss </title>
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
			    <th colspan="9" style="background:#FFFFFF; border-top:1px #ffffff solid; border-bottom:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				
				<div id="header-wrapper">
					<% if  CompanyName <>  "" then %>
						<h4>PT.<%=CompanyName%></h4> 
					<% End if%>
					<% If  FcbaName <> ""  then %>
						   <h5><%=FcbaName%></h5>   
					<% End if%>
				</div>  
				  
   				<div id="subheader" class="centered">
						<h2> ESTIMATE PROFIT AND LOSS PER BLOCK</h3>
						<br/>
						<h4>PERIOD&nbsp;:&nbsp;<%=formatdatevalue0 %> to <%=formatdatevalue %></h4>
				</div>
			</tr>
			




  <tr>
    <th align="center" style="background-color:#00A300;color:#fff;">Company</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Estate</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Block</th>
	<th align="center" style="background-color:#00A300;color:#fff;">Ha</th>
	<th align="center" style="background-color:#00A300;color:#fff;">Planting Year</th>
    <th align="center" style="background-color:#00A300;color:#fff;">&nbsp;</th>
	<th align="center" style="background-color:#00A300;color:#fff;">Period</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Source</th>
    <th align="center" style="background-color:#00A300;color:#fff;">Amount</th>
  </tr>
  </thead>
<tbody>
  <%
    Dim lastCompany,lastfcba,lastCategory,lastField,lastHa,lastTT,lastdate
	Dim TotCost, TotRevenue
	Dim SubPNLBlock
	StrDtl =""
	StrDtl = " select categorycode," & _
			 "	   companycode," & _
			 "	   fcba," & _
			 "     field," & _ 
             "     field_api.getValid_HA(field,fcba,to_date('" & v_txttodaydate & "', 'MM/DD/YYYY')  ) HA," & _
             "     field_api.getValid_plantingyear(field,fcba,to_date('" & v_txttodaydate & "', 'MM/DD/YYYY')  ) TT," & _ 
			 "	   trunc(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),'MM') formdate," & _
			 "	   TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')   ToDate," & _
			 "	   sourcetype   sourcetypecode," & _
			 "	   LOCATIONTYPE_API.get_Name(sourcetype) sourcetype," & _
			 "	   round(Sum(case when categorycode='COST' then amount * -1 else Amount  end),0)  Amount" & _
			 " from estimatepnl " & _
			 " where fcba = '" & v_txtFcba & "' " & _
			 " And to_char(tdate,'MMYYYY')= to_char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),'MMYYYY') " & _
			 " and   trunc(tdate) <= trunc(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') )" & _
			 " group by orderpriority,categorycode, companycode, fcba,field, sourcetype" & _
			 " order by companycode, fcba,field,orderpriority,sourcetype"
    
    Set rsDtl = Server.CreateObject("ADODB.Recordset")
	'response.write StrDtl
	rsDtl.Open StrDtl, connstr, adOpenStatic
    IF Not rsDtl.eof Then
	   lastCompany = ""
	   lastfcba =""
	   lastCategory =""
	   lastField = ""
       lastdate =""
	   lastHa=""
	   lastTT=""
	   SubPNLBlock=0
	   TotCost=0 
	   TotRevenue=0
      While Not rsDtl.eof	
	
		If lastField <> rsDtl("Field") and lastField <> "" then %>
				  <tr>
						<td align="Center" colspan="8">PNL Block <%=lastField %></td>	
						<td align="Right"><%=formatnumber(SubPNLBlock,0)%></td>
						<%
                            SubPNLBlock=0						
						%>
				  </tr>
				  <tr>
						<td align="Left" colspan="10"></td>	
				  </tr>
				  <tr>
						<% If lastCompany <> rsDtl("companycode") then %>
							<td align="Left" style="border-bottom:0px;"><%=rsDtl("companycode")%></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>
						<% If lastfcba <> rsDtl("fcba") then %>
							<td align="Left" style="border-bottom:0px;"><%=rsDtl("fcba")%></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>	
						
						<% If lastField <> rsDtl("Field") then %>
							 <td align="Left" style="border-bottom:0px;"><%=rsDtl("Field") %></td>
							 <td align="Right" style="border-bottom:0px;"><%=rsDtl("Ha") %></td>
							 <td align="Right" style="border-bottom:0px;"><%=rsDtl("TT") %></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>	
								
					   
						<% If lastCategory <> rsDtl("categorycode") or  lastField <> rsDtl("Field") then %>
							 <td align="Left" style="border-bottom:0px;"><%=rsDtl("categorycode") %></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>
						
						<td align="center"><%=rsDtl("formdate") %> - <%=rsDtl("ToDate") %></td>
						<td align="Left"><%=rsDtl("sourcetype") %></td>	
						<td align="right">
							<a target="blank_" href="p_EstimatePNL_detail.asp?txtcompany=<%=rsDtl("companycode")%>&txtFcba=<%=rsDtl("fcba")%>&txtBlok=<%=rsDtl("Field") %>&txtCategory=<%=rsDtl("categorycode") %>&locationtypecode=<%=rsDtl("sourcetypecode") %>&txtFromDate=<%=rsDtl("formdate")%>&txttodate=<%=rsDtl("ToDate")%>">
								<%=formatnumber(rsDtl("Amount"),0)%></a>
						</td>
					</tr>
		<% Else %>
					<tr>
						<% If lastCompany <> rsDtl("companycode") then %>
							<td align="Left" style="border-bottom:0px;"><%=rsDtl("companycode")%></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>
						<% If lastfcba <> rsDtl("fcba") then %>
							<td align="Left" style="border-bottom:0px;"><%=rsDtl("fcba")%></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>			
						
						<% If lastField <> rsDtl("Field") then %>
							 <td align="Left" style="border-bottom:0px;"><%=rsDtl("Field") %></td>
							 <td align="Right" style="border-bottom:0px;"><%=rsDtl("Ha") %></td>
							 <td align="Right" style="border-bottom:0px;"><%=rsDtl("TT") %></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>	
						
						<% If lastCategory <> rsDtl("categorycode") or  lastField <> rsDtl("Field") then %>
							 <td align="Left" style="border-bottom:0px;"><%=rsDtl("categorycode") %></td>
						<% Else %>
							<td align="Left" style="border-top:0px;border-bottom:0px;"></td>
						<% End If %>
						
						<td align="center"><%=rsDtl("formdate") %> - <%=rsDtl("ToDate") %></td>
						<td align="Left"><%=rsDtl("sourcetype") %></td>	
						<td align="right">
						   <a target="blank_" href="p_EstimatePNL_detail.asp?txtcompany=<%=rsDtl("companycode")%>&txtFcba=<%=rsDtl("fcba")%>&txtBlok=<%=rsDtl("Field") %>&txtCategory=<%=rsDtl("categorycode") %>&locationtypecode=<%=rsDtl("sourcetypecode") %>&txtFromDate=<%=rsDtl("formdate")%>&txttodate=<%=rsDtl("ToDate")%>">
							  <%=formatnumber(rsDtl("Amount"),0)%>
						   </a>
						</td>
					</tr>

		<%	
		       
		  End If 
		  
		     SubPNLBlock = cdbl(SubPNLBlock) + cdbl(rsDtl("Amount"))
			 IF trim(rsDtl("categorycode")) = "COST" Then
			     TotCost= cdbl(TotCost) + cdbl(rsDtl("Amount"))
			 Else
	             TotRevenue= cdbl(TotRevenue) + cdbl(rsDtl("Amount"))
		     End if		
		     lastCompany = rsDtl("companycode")
			 lastfcba = rsDtl("fcba") 
			 lastCategory = rsDtl("categorycode") 
			 lastField = rsDtl("Field") 
			 lastHa = rsDtl("Ha")
			 lastTT =rsDtl("TT") 
			rsDtl.MoveNext
		Wend
%>		
		<tr>				
			<td align="Center" colspan="8">PNL Block <%=lastField %></td>	
			<td align="Right"><%=formatnumber(SubPNLBlock,0)%></td>
		</tr>
		<tr>	
			<td align="Left" colspan="10"></td>	
		</tr>
        <tr>
			<td align="Left" colspan="8">Total Revenue </td>	
			<td align="Right"><%=formatnumber(TotRevenue,0)%></td>
		</tr>
		<tr>
			<td align="Left" colspan="8">Total Cost </td>	
			<td align="Right"><%=formatnumber(TotCost,0)%></td>
		</tr>
		<tr>
			<td align="Left" colspan="8">Profit and Loss </td>	
			<td align="Right"><%=formatnumber(cdbl(TotRevenue)+ cdbl(TotCost),0)%></td>
		</tr>
<%		
	Else	
		%>			
        <tr>
			<td align="center" colspan="9">NO DATA AVAILABLE ( Please Genarate Data Period <%=formatdatevalue0 %> - <%=formatdatevalue %> ) </td>
		</tr>
		
<%  End If %>  	
				  
		
 </tbody>	
</table>
<!--#include file="bottom.Asp"-->