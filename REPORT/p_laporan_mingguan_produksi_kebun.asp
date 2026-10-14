<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	
	
	
	
	
	'Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom
	Dim strsql,rsCount, v_txtRegion, nSpanTotal
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, v_txtyear,v_txtmonth,v_txtWEEK
	Dim ScriptName,visibleR,visibleGR
	Dim strMain, strHarianLC, strHarianLC2, rsHarianLC
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	dim CountFcba, rsCountFcba, coloumnname , rscoloumn, rscoloumnname
	
	
	Dim v_txtGroupRegion, txtyear, txtmonth
	
	
	
    'v_txtGroupPlantation = Request("txtGroupPlantation")
	'v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	'v_txtRegion = Request("txtRegion")
	'v_txtcompany = Request("txtCompany")
	'v_txtfcba = Request("txtfcba")
	
	'v_txtBudgetYear = Request("txtBudgetYear")
	'v_txtBudgetVersion = Request("txtBudgetVersion")
	'v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtyear             = Request("txtYear")
	v_txtmonth            = Request("txtMonth") 
	v_txtWEEK            = Request("txtWEEK")
	
	
	'If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textDate(Date)
	
	Dim var_colspan : var_colspan = 1
	
	strsql = " SELECT count(*) ncount from ( SELECT DISTINCT a.region fccode  " & _
			 " FROM Businessunit a, region b  " & _
			 " WHERE a.region = b.fccode  " & _
			 " AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', a.fccode) =  'TRUE'  ) " 
	Set rsCount = Server.CreateObject("ADODB.Recordset")
	rsCount.Open strsql, connStr
	
	if not rsCount.eof then
     visibleR = rsCount("ncount")
	 rsCount.Close
	 if cint(visibleR) = 1 then 
	 
	strsql = " SELECT DISTINCT a.region fccode  " & _
			 " FROM Businessunit a, region b  " & _
			 " WHERE a.region = b.fccode  " & _
			 " AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', a.fccode) =  'TRUE'  " 
		Set rsCount = Server.CreateObject("ADODB.Recordset")
		 rsCount.Open strsql, connStr
		 txtRegion = rsCount("fccode")
		 rsCount.Close
	 end if	 
  else
     visibleR = 0
	 rsCount.Close
  end if
	'response.write visibleR

	' Query Main Grid 	
	'v_txtyear 		= Request("txtyear")
	'v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	'v_txtRegion  	= Request("txtRegion")
	nSpanTotal 		= 8
	
	Dim wherecons, selectCons, orderCons
	wherecons = "where region is not null"
	
	if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
		if v_txtRegion <> "--ALL--" then
			whereCons = whereCons & " And  region = '" & v_txtRegion & "' "
		end if
			selectCons = selectCons & " ,region"
			orderCons  = orderCons & ", REGION_API.get_OrderData(region) "
		else
			nSpanTotal = nSpanTotal - 1 
	end if
	
	
%>









<%
	' Query Main Grid 	
	'v_txtyear 		= Request("txtyear")
	'v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	'v_txtRegion  	= Request("txtRegion")
	'nSpanTotal 		= 8
	
	
	
%>


<%


If len(trim(v_txtGroupRegion)) < 1 then
   v_txtGroupRegion ="BATAM"
end if


coloumnname = "select FCCODE from businessunit a where a.GROUPREGION = '"& v_txtGroupRegion &"' and fctype='E' and isactive='TRUE' ORDER BY FCCODE ASC"

	
CountFcba = " select count (*) total from businessunit a where a.fctype='E' and a.GROUPREGION = '"& v_txtGroupRegion &"'"	

Set rsCountFcba = Server.CreateObject("ADODB.Recordset")
rsCountFcba.Open CountFcba, connstr, adOpenStatic

	dim kjTot
	dim cols
	if NOT rsCountFcba.eof then
   	   kjTot = cdbl (rsCountFcba("total"))
	   cols=2
	   cols=(cols+9)*kjTot
	end if
	rsCountFcba.Close
	set rsCountFcba = Nothing
	'response.write cols
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Mingguan Kebun</title>
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
            $( "#datepicker" ).datepicker();
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
				@media print{.printhide{display:none;}}
			</style>
       </head>


	<body>
		<div id="wrapper">
		<div class="printhide">
		<!--#include file="../print_excel.asp"-->
		</div>
			<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include file="menutop.asp"-->
				
				
					<table id="form-table" class="centered-margin">
						<tr>
										
								
							<td><font face="verdana" size="1">GROUP REGION</font></td>
							<td>
								<select name="txtGroupRegion">
																		<%
									Dim rsRegion
									Set rsRegion = Server.CreateObject("ADODB.Recordset")
									rsRegion.CursorLocation = adUseClient
									strsql =  " SELECT DISTINCT a.groupregion fccode  " & _
											  " FROM Businessunit a, region b  " & _
											  " WHERE a.region = b.fccode  " & _
											  " AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', a.fccode) =  'TRUE'  " 

									rsRegion.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsRegion.eof then
										DO while not rsRegion.eof
											if v_txtGroupRegion = rsRegion("fccode") then
										%>			
											<option value="<%=rsRegion("fccode")%>" Selected><%=rsRegion("fccode")%></option>
										<% Else %>	
											<option value="<%=rsRegion("fccode")%>"><%=rsRegion("fccode")%></option>
										<% End If
										
											rsRegion.movenext
										loop
									End if
									rsRegion.Close
									set rsRegion = Nothing
								%>
								</select>		

							</td>
							
							<td><font face="verdana" size="1">TAHUN</font></td>
							<td align="center">
								<font size="2"  color="#0000FF" ></font>
								 <% 
									MM=v_txtMonth
									yyyy=v_txtYear

									mon = Month(Session("v_Costbook"))
									If len(MM)=0 then MM=mon

									Yr = Year(Session("v_Costbook"))
									MinYr = 2015
									MaxYr = Yr  + 2
									If len(yyyy)= 0 then yyyy = Yr
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
							<td><font face="verdana" size="1">BULAN</font></td>
							<td align="center">
							  <font size="2" color="#0000FF" ></font>
									<select name="txtMonth">
										<%
											For i = 1 to 12
										%>
												<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
										<%
											Next
										%>
									</select>
							</td>
							<td><font face="verdana" size="1">MINGGU</font></td>
							<td>
								<select name="txtWEEK">
										<option value="1" <%if request("v_txtWEEK")="1" then%> selected <%End if%>>1</option>
										<%
											Dim ListWeek
											ListWeek = Array( "2", "3","4","5","6")
										For i =0 to 4
											if v_txtWEEK = ListWeek(i) then
										%>
											<Option Value="<%=ListWeek(i)%>" Selected><%=ListWeek(i)%></option>
										<%
											ELSE
										%>
											<Option Value="<%=ListWeek(i)%>"><%=ListWeek(i)%></option>
										<%
											end if
											Next
										%>
								</select>
							</td>
							<td align="center">
								<input type="submit" value="View" />
							</td>
						</tr>
						
					</table>
				</div>
			</form>

			<table border="0" width="100%" id="product-table" class='xthf-site'>
				<thead>
					<tr>
						<th colspan="350" style="background-color:#FFFFFF; border:none;">
							<div id="header-wrapper">
							 <% Dim ia 
							    If len(trim(v_txtyear))	< 1 then
								   v_txtyear= Year(now)
								   v_txtmonth= Month(now)
								   v_txtWEEK="1"
								end if 
							    ia = v_txtmonth
								
							 %>
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN MINGGUAN PRODUKSI KEBUN</h3>
									<br/>
								</div>
								<div style="text-align:centered;">
									<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=v_txtyear %></h5>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th rowspan="4" colspan="2" width="25%" nowrap>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp DESCRIPTION &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>
						<th rowspan="4" >UNIT</th>
						<!-- <th rowspan="4">Luas (Ha) TM</th> -->
                        
                        <%
							Set rscoloumnname = Server.CreateObject("ADODB.Recordset")
							rscoloumnname.Open coloumnname, connstr, adOpenStatic
							if not rscoloumnname.eof then
							DO while not rscoloumnname.eof
							
						
						 if kjTot > 0 then
						%>
							 <th colspan="9"><%=rscoloumnname("FCCODE")%></th>
                             
						<%						
						else
						%>
						<th style="color:#FFFFFF" colspan="<%=kjTot %>" width="10%">BUSINESSUNIT</th>
						<%
						end if
						%>
                        <%
							rscoloumnname.movenext
							LOOP 
							END IF
							
							
						%>
                        
                        
						
                        
                        <th colspan="18">TOTAL</th>
						
					</tr>
					<tr>
						<%
							Set rscoloumnname = Server.CreateObject("ADODB.Recordset")
							rscoloumnname.Open coloumnname, connstr, adOpenStatic
							if not rscoloumnname.eof then
							DO while not rscoloumnname.eof
						
                        if kjTot > 0 then
						%>
							<th rowspan="3" colspan="1">SBL</th>
                            <th colspan="3">MG INI</th>
                            <th colspan="3">s/d MG INI</th>
                            <th colspan="2">SBI</th>
                             
						<%						
						else
						%>
						<th style="color:#FFFFFF" colspan="<%=kjTot %>" width="10%">BUSINESSUNIT</th>
						<%
						end if
						%>
                        <%
							rscoloumnname.movenext
							LOOP 
							END IF
						%>
                        <th rowspan="3" colspan="1">SBL</th>
                        <th colspan="3">MG INI</th>
                        <th colspan="3">s/d MG INI</th>
                        <th colspan="2">SBI</th>
						
					</tr>
					<tr>
						<%
							Set rscoloumnname = Server.CreateObject("ADODB.Recordset")
							rscoloumnname.Open coloumnname, connstr, adOpenStatic
							if not rscoloumnname.eof then
							DO while not rscoloumnname.eof
						
                        if kjTot > 0 then
						%>
                        <th>Akt</th>
						<th>Bgt</th>
                        <th>Acv</th>
                        <th>Akt</th>
						<th>Bgt</th>
                        <th>Acv</th>
                        <th>Akt</th>
						<th>Bgt</th>
                        <%						
						else
						%>
						<th style="color:#FFFFFF" colspan="<%=kjTot %>" width="10%">BUSINESSUNIT</th>
						<%
						end if
						%>
                        <%
							rscoloumnname.movenext
							LOOP 
							END IF
						%>
                        <th>Akt</th>
						<th>Bgt</th>
                        <th>Acv</th>
                        <th>Akt</th>
						<th>Bgt</th>
                        <th>Acv</th>
                        <th>Akt</th>
						<th>Bgt</th>                        
					</tr>
					<tr>
						
						
					</tr>
				</thead>
				
				<tbody>
<%


Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

Dim CurrRegion : CurrRegion = ""
Dim PrevRegion : PrevRegion = ""

Dim CurrGroupRegion : CurrGroupRegion = ""
Dim PrevGroupRegion : PrevGroupRegion = ""
Dim Xi
Dim sbl,akt,bgt,acv
Dim aktSD, bgtSD, acvSD
Dim aktSBI, bgtSBI
Dim sqlcentre, rssqlcentre,str1,str2, str3, str4,str5,str6,sqlweekly,rssqlweekly,i
Dim ordata1, ordata2, orderdatagrp, rsorderdatagrp, nmgrp,nmsubgrp

	nmgrp=""
	nmsubgrp=""

sqlcentre =  " SELECT  TO_CHAR(DBMS_LOB.substr(DynamicQuery_Api.dynamic_WeeklyReport('"& v_txtGroupRegion &"'),4000,1)) str1, " & _
			 "	TO_CHAR(DBMS_LOB.substr(DynamicQuery_Api.dynamic_WeeklyReport('"& v_txtGroupRegion &"'),4000,4001)) str2,  " & _
			 "	TO_CHAR(DBMS_LOB.substr(DynamicQuery_Api.dynamic_WeeklyReport('"& v_txtGroupRegion &"'),4000,8001)) str3,  " & _
			 "	TO_CHAR(DBMS_LOB.substr(DynamicQuery_Api.dynamic_WeeklyReport('"& v_txtGroupRegion &"'),4000,12001)) str4, " & _
			 "	TO_CHAR(DBMS_LOB.substr(DynamicQuery_Api.dynamic_WeeklyReport('"& v_txtGroupRegion &"'),4000,16001)) str5 " & _
		     " FROM DUAL "
'response.write sqlcentre

Set rssqlcentre = Server.CreateObject("ADODB.Recordset")
rssqlcentre.Open sqlcentre, connstr, adOpenStatic

	if not rssqlcentre.eof then
	    str1=rssqlcentre("str1")
	    str2=rssqlcentre("str2")
	    str3=rssqlcentre("str3")
		str4=rssqlcentre("str4")
		str5=rssqlcentre("str5")
	END IF 
	
	

	sqlweekly = str1 & str2 & str3 & str4 & str5 & "having year =" & v_txtyear & " and month =" & v_txtmonth &" and onweek = " & v_txtWEEK & " order by datagroupcode,orderdata"

    'response.write sqlweekly
	
	Set rssqlweekly = Server.CreateObject("ADODB.Recordset")
	rssqlweekly.Open sqlweekly, connstr, adOpenStatic
	
	
if rssqlweekly.bof and rssqlweekly.eof then
%>
					<tr>
						<td colspan="<%=182%>">NO DATA AVAILABLE</td>
					</tr>
<%
else			
        If rssqlweekly.fields.count >= 9 then
		Dim totalsbl : totalsbl = 0
		Dim totalakt : totalakt =0
		Dim totalbgt : totalbgt = 0
		Dim totalacv : totalacv =0
		Dim totalaktSD : totalaktSD =0
		Dim totalbgtSD : totalbgtSD = 0
		Dim totalacvSD : totalacvSD =0
		Dim totalaktSBI : totalaktSD =0
		Dim totalbgtSBI : totalbgtSD = 0
%>						
					
			
<%			
             while not rssqlweekly.eof
				Dim groupdata, groupdt, subgroupdata, subgroupdt
				
				groupdata = rssqlweekly("DataGroupCode")
				groupdt = Left(groupdata, 2)
				
				subgroupdata = rssqlweekly("DataGroupCode")
				subgroupdt = Mid(subgroupdata, 3,1)
				
%>			 
						<%
							if nmgrp <> rssqlweekly("GROUPCODENAME") then	
						%>
						<tr>
							<td align="left" colspan="<%=cols %>"  width="10%" style=" border-bottom:none;"><%=groupdt &" "%> <%=rssqlweekly("GROUPCODENAME")%></td>
						</tr>
						<% nmgrp = rssqlweekly("GROUPCODENAME")%>
						<% else	%>
						
						<% end if %>
						
						
						<%
							if nmsubgrp <> rssqlweekly("SUBGROUPCODENAME") then	
						%>
						<tr>
							<td align="left" colspan="<%=cols %>"  width="10%" style=" border-top:none;">&nbsp&nbsp&nbsp&nbsp<%=subgroupdt &" " %> <%=rssqlweekly("SUBGROUPCODENAME")%></td>
						</tr>
						<% nmsubgrp = rssqlweekly("SUBGROUPCODENAME")%>
						<% else	%>
						
						<% end if %>
						
						<tr>
							<td align="left"  width="1%" style="border-right:none;">&nbsp&nbsp</td>
							<td align="left"  width="14%" style="border-left:none;">- &nbsp<%=rssqlweekly("DataGroupCodeName")%></td>
							<td align="left" ><%=rssqlweekly("unitofmeasurement")%></td>
					    
							 <%
							 sbl =9
							 akt =10
							 bgt =11
							 acv =12
							 
							 aktSD =13
							 bgtSD =14
							 acvSD =15
							 
							 aktSBI =13
							 bgtSBI =14
							 
							 for xi = 9 to rssqlweekly.fields.count -1
									if len(trim(rssqlweekly(xi).value)) > 0 then%>
									
									   <td align="right"><%=FormatNumber(rssqlweekly(xi),0)%></td>
									   
									<%
									if xi = sbl then
									totalsbl = totalsbl + Cdbl(rssqlweekly(xi))
									sbl = sbl + 9
									end if
									
									'response.write akt&"="&xi&" "
									if xi = akt then
									'response.write "hasil ="&totalakt&"| "
									totalakt = totalakt + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									akt = akt + 9
									end if
									
									if xi = bgt then
									'response.write "hasil ="&totalakt&"| "
									totalbgt = totalbgt + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									bgt = bgt + 9
									end if
									
									if xi = acv then
									'response.write "hasil ="&totalakt&"| "
									totalacv = totalacv + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									acv = acv + 9
									end if
									
									if xi = aktSD then
									'response.write "hasil ="&totalakt&"| "
									totalaktSD = totalaktSD + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									aktSD = aktSD + 9
									end if
									
									if xi = bgtSD then
									'response.write "hasil ="&totalakt&"| "
									totalbgtSD = totalbgtSD + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									bgtSD = bgtSD + 9
									end if
									
									if xi = acvSD then
									'response.write "hasil ="&totalakt&"| "
									totalacvSD = totalacvSD + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									acvSD = acvSD + 9
									end if
									
									if xi = aktSBI then
									'response.write "hasil ="&totalakt&"| "
									totalaktSBI = totalaktSBI + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									aktSBI = aktSBI + 9
									end if
									
									if xi = bgtSBI then
									'response.write "hasil ="&totalakt&"| "
									totalbgtSBI = totalbgtSBI + Cdbl(rssqlweekly(xi))
									'response.write rssqlweekly(xi)
									bgtSBI = bgtSBI + 9
									end if
									
									'response.write sbl
									else%>
									   <td align="right" colspan="2">-</td>
									<%end if%>
							 <%next%>
							<td align="right"><%=FormatNumber(totalsbl,0)%></td>
							<td align="right"><%=FormatNumber(totalakt,0)%></td>
							<td align="right"><%=FormatNumber(totalbgt,0)%></td>
							<td align="right"><%=FormatNumber(totalacv,0)%></td>
							
							<td align="right"><%=FormatNumber(totalaktSD,0)%></td>
							<td align="right"><%=FormatNumber(totalbgtSD,0)%></td>
							<td align="right"><%=FormatNumber(totalacvSD,0)%></td>
							
							<td align="right"><%=FormatNumber(totalaktSBI,0)%></td>
							<td align="right" colspan="2"><%=FormatNumber(totalbgtSBI,0)%></td>
					</tr>
<%		
				rssqlweekly.MoveNext			
				PrevRegion = CurrRegion
				PrevGroupRegion = CurrGroupRegion
				
			totalsbl = 0
			totalakt =0
			totalbgt = 0
			totalacv =0
			totalaktSD =0
			totalbgtSD = 0
			totalacvSD =0
			totalaktSBI =0
			totalbgtSBI = 0
			Wend	
		else
%>		
	   		<tr>
						<td colspan="<%=182%>">NO DATA AVAILABLE</td>
					</tr>
<%		
     	end if
rssqlweekly.Close
set rssqlweekly = Nothing
%>

					
<%End if%>

				</tbody>
			</table>

		</div>
	</body>
</html>
