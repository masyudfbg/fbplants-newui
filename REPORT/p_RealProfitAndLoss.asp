<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%

Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonthto,v_txtmonth
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_txtjob,v_txtjobgroup
Dim tempyear,tempmonth,no
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname, Sprevblock
Dim sprevPlantingYear, sprevFieldCode
Dim tempjob
Dim MM,MF, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob
Dim costsbi_,Revsbi_,Pnlsbi_
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim str1,str2,str3,strMain,whereCons,whereCons1,selectCons,orderCons
Dim var_ColSpan : var_ColSpan = 0
Dim GrandCost(36)
Dim GrandRev(36)
Dim Grandpnl(36)
Dim GrandSbiCost, GrandSbiRev, GrandSbiPNL


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
           
   		  
			v_txtfcba 			  = Request("txtfcba")
			'v_txtPlantingDate	  = Request("txtPlantingDate")
			v_txtBlok			  = Request("txtBlok")
			v_txtyear             = Request("txtYear")
			'v_txtmonthto          = Request("txtMonthto")
			v_txtmonth            = Request("txtMonth")
			'v_vstatus             = Request("vStatus")
			'v_txtDivision		  = Request("txtDivision")
			'v_txtjob			  = Request("txtjob")
			'v_txtjobgroup		  = Request("txtjobgroup")
			
			if v_txtmonth ="" then
			   v_txtmonth = Month(now)
			end if
			
			nSpanTotal = 9
			
			strcurrentdate = " select currentdate" & _
                 " from periodcontrol" & _                 
                 " where fcba = '" & v_txtFcba & "'"

			Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
			rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

			If Not rscurrentdate.EOF Then
				varDefaultYear	=	Year(rscurrentdate("currentdate"))
				varDefaultMonth	=	Month(rscurrentdate("currentdate"))
			Else
				varDefaultYear	=	Year(Session("v_Employee"))
				varDefaultMonth	=	Month(Session("v_Employee"))
			End If
			rscurrentdate.Close
			set rscurrentdate = Nothing

			If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
			If (v_txtMonthto = "" Or Len(Trim(v_txtMonthto)) = 0) Then v_txtMonthto = varDefaultMonth
			'If (v_txtMonthfrom = "" Or Len(Trim(v_txtMonthfrom)) = 0) Then v_txtMonthfrom = varDefaultMonth
	        
			
            
			if  v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""  then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & " , PLANTINGYEAR "
			else
			      nSpanTotal = nSpanTotal - 1 	  
			end if
			 	
			if  v_txtDivision <> "--SUMMARY--" and v_txtDivision <> ""  then
				   if v_txtDivision <>  "--ALL--" then
					  whereCons = whereCons & " AND Afdeling = '" & v_txtDivision & "'"
				   end if			    
			end if
			          			 
		     if v_txtBlok <>  "--ALL--" and v_txtBlok <>  "" then
				whereCons1 = whereCons1 & " AND targetcode = '" & v_txtBlok & "'"
			end if
			
		    if v_txtjob <>  "--ALL--" and v_txtBlok <>  "" then
				whereCons1 = whereCons1 & " AND jobcode = '" & v_txtjob & "'"
			end if

       q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtFcba & "'"
		 set rsCat = connStr.Execute(q)
		 if not rsCat.eof then
		    fcbaname = rsCat("fcName")
		    CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
		    KTUName=rsCat("FCKTU")
         end if

	   
%>			

<html>

<head>
<title>iPla'S Reporting - Real Profit And Loss</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->
<style type="text/css">
@media print
{
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
		<div class="printhide">
		
		<form action="<%=ScriptName%>" method="post">
		<!--#include file="menutop.asp"-->
		<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
		<table id="form-table" class="centered-margin">
			<tr>
				<th>TAHUN</th>
				<th>BULAN</th>
				<th>BLOCK</th>
				<th>&nbsp;</th>
			</tr>
			<tr>

					<td>
						<font size="2"  color="#0000FF" ></font>
						<% 'dim v_txtMonthfrom
							MF=v_txtmonth
							'MM=v_txtMonthto
							yyyy=v_txtYear

							mon = Month(Session("v_Costbook"))
							If len(MM)=0 then MM=mon

							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(yyyy)= 0 then yyyy = Yr
							'response.write MF
						%>
						<select name="txtYear">
								<%
									For i = MinYr to MaxYr
								%>
										<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
								<%
									Next
								%>
						</select>
					</td>
			
					<td>
						<font size="2" color="#0000FF" ></font>
							<select name="txtMonth">
								<%
									For i = 1 to 12
								%>
										<Option Value="<%=i%>" <%If cint(i) = cint(MF) then%> Selected <%End If%>><%=MonthName(i)%></option>
								<%
									Next
								%>
							</select>
					</td>		
					<td align="center">
						<select name="txtBlok">	
							<option value="--ALL--">--ALL--</option>
							<%
								Dim rsBlok
								Set rsBlok = Server.CreateObject("ADODB.Recordset")
								rsBlok.CursorLocation = adUseClient

						
								if len(trim(v_txtBlok)) > 0 then 
									strsql =  " SELECT DISTINCT fccode as Blok,fcname FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'  order by fccode"
								else
									strsql =  " SELECT DISTINCT fccode as Blok,fcname FROM FIELDDETAIL   order by fccode"
								end if
								
								rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsBlok.eof then
									DO while not rsBlok.eof
									if v_txtBlok = rsBlok("Blok") then
							%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%> - <%=rsBlok("fcname")%> </option>
									<%else%>
										<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%> - <%=rsBlok("fcname")%> </option>
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
		</div>
</div>   
<table border="0" id="product-table" class="xthf-site" width="100%">
	<thead>
        <tr>
          <th colspan="42" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">	    
		<div id="header"  class="centered">
		   <% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			
		</div>
							<% Dim ia 
							    If len(trim(v_txtyear))	< 1 then
								   v_txtyear= Year(now)
								   v_txtmonth= Month(now)
								   v_txtWEEK="1"
								end if 
							    ia = v_txtmonth
								
								
							 %>
		<div id="subheader" class="center">
    	 <h3>LAPORAN REAL PROFIT AND LOSS</h3>   
		 <h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=v_txtyear %></h5>
		</div>
		
        </th>
    </tr>
             
     		 
			  <tr>
                <th  style="background-color:#00A300;color:#fff;" rowspan="2"><div align="center">Block</div></th>
				<th  style="background-color:#00A300;color:#fff;" rowspan="2"><div align="center">Hectare (Ha)</div></th>
				<th  style="background-color:#00A300;color:#fff;" rowspan="2"><div align="center">Planting Year</div></th>
				<% FOR i = 1 to  Cint(v_txtmonth)  %>      
						<th  style="background-color:#00A300;color:#fff;" colspan="3"><div align="center"><%=UCASE(MonthName(cint(i)))%></div></th>
				<% Next %>	
				<th  style="background-color:#00A300;color:#fff;" colspan="3"><div align="center">SBI</div></th>
              </tr>
              <tr>
				<% FOR i = 1 to  Cint(v_txtmonth)  %>   	   
					<th style="background-color:#00A300;color:#fff;"><div align="center">Cost</div></th>
					<th style="background-color:#00A300;color:#fff;"><div align="center">Revenue</div></th>
					<th style="background-color:#00A300;color:#fff;"><div align="center">PNL</div></th>
				<% Next %>  
			    <th style="background-color:#00A300;color:#fff;"><div align="center">Cost</div></th>
                <th style="background-color:#00A300;color:#fff;"><div align="center">Rev</div></th>
                <th style="background-color:#00A300;color:#fff;"><div align="center">PNL</div></th>
           </tr>
 </thead>
 
    <tbody>
		<% 
		
		Dim sqlcentre,rssqlcentre,xi,strwhr
		


		sqlcentre =  " SELECT  DynamicQuery_Api.dynamic_GetSqlRealPNL('" & v_txtfcba &"'," & Cint(v_txtmonth) & "," & Cint(v_txtyear) & ") str1 " & _
					 " FROM DUAL "
		'response.write sqlcentre

		Set rssqlcentre = Server.CreateObject("ADODB.Recordset")
		rssqlcentre.Open sqlcentre, connstr, adOpenStatic

			if not rssqlcentre.eof then
				strMain=rssqlcentre("str1")
			END IF 
			
	   ' response.write  v_txtBlok
		if v_txtBlok <>  "--ALL--" then
		     strwhr = strwhr & " where  FIELDCODE = '" & v_txtBlok & "'"
		end if
		 strMain = " Select * from ( " & strMain & ") " & strwhr & " order by fieldcode "
		'response.write strMain
		dim rsreal,cClass,subtotal, statuspnl, bgcolor
		Set rsreal = Server.CreateObject("ADODB.Recordset")
		rsreal.Open strMain, connstr, adOpenStatic

		'Dim Rp_Per_Ha_RowReal : Rp_Per_Ha_RowReal = 0
		'Dim Rp_Per_Ha_RowBudget : Rp_Per_Ha_RowBudget = 0
		'Dim RowPercentage : RowPercentage = 0

		if rsreal.bof and rsreal.eof then
		%>
			<tr>
				<td colspan="<%=var_colspan+42%>">NO DATA AVAILABLE</td>
			</tr>
		<%
		else
				No=0
				Sprevblock=""
				while not rsreal.eof
					if no  mod 2 = 1 then
						cClass = "odd"
					else
						cClass = "default"
					end if
					No = No + 1 
		%>


<tr>
						<tr class="<%=cClass%>">	
							    <td align="center"><%=rsreal("FIELDCODE")%></td>
								<td align="Right"><%=rsreal("Ha")%></td>
								<td align="center"><%=rsreal("tt")%></td>
							<%	Costsbi_ = 0
								Revsbi_ =0
								Pnlsbi_ = 0
								

							    for xi = 1 to Cint(v_txtmonth)   %>
									<td align="Right"><%=FormatNumber(rsreal("COST_" & Lpad (xi,"0",2) & ""), 0)	%></td>	
									<td align="Right"><%=FormatNumber(rsreal("REVENUE_" & Lpad (xi,"0",2)  & ""), 0)	%></td>	
									<td align="Right"><%=FormatNumber(rsreal("PNL_" & Lpad (xi,"0",2)  & ""), 0)	%></td>	
									<%
									  
									  GrandRev(xi) = cdbl(GrandRev(xi)) + cdbl(rsreal("COST_" & Lpad (xi,"0",2) & ""))
									  GrandCost(xi) = cdbl(GrandCost(xi)) + cdbl(rsreal("REVENUE_" & Lpad (xi,"0",2) & ""))
									  Grandpnl(xi) = cdbl(Grandpnl(xi)) + cdbl(rsreal("PNL_" & Lpad (xi,"0",2) & ""))
									  
									  Costsbi_ = cdbl(Costsbi_) + cdbl(rsreal("COST_" & Lpad (xi,"0",2) & ""))
									  Revsbi_ = cdbl(Revsbi_) + cdbl(rsreal("REVENUE_" & Lpad (xi,"0",2) & ""))
									  Pnlsbi_ = cdbl(Pnlsbi_) + cdbl(rsreal("PNL_" & Lpad (xi,"0",2) & ""))
							  Next 
							  
							  GrandSbiCost = cdbl(GrandSbiCost) + cdbl(Costsbi_) 
							  GrandSbiRev  = cdbl(GrandSbiRev) + cdbl(Revsbi_) 
							  GrandSbiPNL  = cdbl(GrandSbiPNL) + cdbl(Pnlsbi_)
							  
							  %>	
								<td align="Right"><%=FormatNumber(Costsbi_, 0)	%></td>	
								<td align="Right"><%=FormatNumber(Revsbi_, 0)	%></td>	
								<td align="Right"><%=FormatNumber(Pnlsbi_, 0)	%></td>																																																																												
						</tr>
		<% 
					rsreal.MoveNext
					Wend	
		rsreal.Close
		set rsreal = Nothing
		%>
		
		<tr bgcolor="#F4F4A0">
			<td colspan="3" align="center"><strong>GRAND TOTAL</strong></td>	
			<% for xi = 1 to Cint(v_txtmonth) %>	
					 <td align="right"><strong><%=FormatNumber(GrandRev(xi), 0)	%></strong></td> 
					 <td align="right"><strong><%=FormatNumber(GrandCost(xi), 0)%></strong></td> 
					 <td align="right"><strong><%=FormatNumber(Grandpnl(xi), 0)	%></strong></td> 
			<%  next%>
			<td align="right"><strong><%=FormatNumber(GrandSbiCost, 0)%></strong></td>  
			<td align="right"><strong><%=FormatNumber(GrandSbiRev, 0)%></strong></td>  
			<td align="right"><strong><%=FormatNumber(GrandSbiPNL, 0)%></strong></td>  
		</tr>
		
		<%End if%>
	</tbody>	
</table>
<!--#include file="bottom.Asp"-->
