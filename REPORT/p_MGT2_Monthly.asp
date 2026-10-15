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
								<option value="--ALL--"Selected>--ALL--</option>
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
											<option value="<%=rsRegion("fccode")%>"><%=rsRegion("fccode")%></option>
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
									<h1>LAPORAN MINGGUAN KEBUN</h1>
									<div style="margin-botton:10px;"/>
									<h3></h3>
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
						
                        <th colspan="3">THIS MONTH</th>
                        <th colspan="3">TO DATE</th>
                        <th rowspan="2">&nbspANNUAL BUDGET</th>
						
					</tr>
					<tr>
							<th colspan="1">ACT</th>
                            <th colspan="1">BGT</th>
                            <th colspan="1">%</th>
							<th colspan="1">ACT</th>
                            <th colspan="1">BGT</th>
                            <th colspan="1">%</th>
                             
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
Dim aktSBI, bgtSBI,i,actMTot,actDTot,budDTot,budMTot,budATot,abtot1,abtot2
Dim sql1,rssql1,sql2,rssql2,sql3,rssql3,sqld1,sqld2,sqld3,sqld4,sqld5,sqlc
Dim hs1,hs2,hs3,hs4,hs5
DIM sqlc1,rssqlc1, sqlc2,rssqlc2,sqlc3,rssqlc3,sqlc4,rssqlc4,sqlc5,rssqlc5
Dim ordata1, ordata2, orderdatagrp, rsorderdatagrp, nmgrp,nmsubgrp,links
Dim persentase1 : persentase1  = 0
Dim persentase2 : persentase2 = 0
Dim persentase : persentase = 0
Dim persentased1 : persentased1  = 0
Dim persentased2 : persentased2 = 0
Dim persentased : persentased = 0
Dim persentasea : persentasea = 0
Dim tempAdditionalString
Dim tempAdditionalString_01

sql1 = "select subgroupcode,report,description from MGT2_SETUP_SGRPT_TAB order by orderdata"

Set rssql1 = Server.CreateObject("ADODB.Recordset")
	rssql1.Open sql1, connstr, adOpenStatic
	if not rssql1.eof then
	DO while not rssql1.eof 
%>				
			<tr>
						<td colspan="10" bgcolor="yellow"><%=rssql1("description")%></td>
			</tr>
<%
sql2 = "select datagroupcode,subgroupcode,str_act_thismonth,str_bud_thismonth,str_act_todate,str_bud_todate,str_bud_annual,description,unitofmeasurement,string_link,string_parameter, addwherestmt, addwherestmt_01  from MGT2_SETUP_DATA_TAB where subgroupcode ='"&rssql1("subgroupcode")&"' order by orderdata"

Set rssql2 = Server.CreateObject("ADODB.Recordset")
	rssql2.Open sql2, connstr, adOpenStatic
	actDTot =0
	actMTot =0
	budMTot =0
	budDTot = 0
	budATot = 0
	abtot1 =0
	abtot2 = 0
	if not rssql2.eof then
	DO while not rssql2.eof 

        tempAdditionalString = ""

	    if rssql2("addwherestmt") <> "" then
	        tempAdditionalString = " and " & rssql2("addwherestmt") & " "
	        tempAdditionalString_01 = " and " & rssql2("addwherestmt_01") & " "
	    end if



	    if rssql2("datagroupcode") = "REPGROUP010101" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010102" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010103" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010201" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010202" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010203" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010204" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010301" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtstatus=TBM"
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010302" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtstatus=TM"
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010303" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtstatus=MN"
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010401" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtBudgetYear="&v_txtYear
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010402" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010403" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtBudgetYear="&v_txtYear
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010404" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010405" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010406" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010501" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtBudgetYear="&v_txtYear
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010502" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010601" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&txtBudgetYear="&v_txtYear
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010602" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010701" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&vstatus=TBM&txtPlantingDate="&v_txtYear&"&vCROPTYPE=OP&vOwnership=ALL"
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010702" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&vstatus=TBM&txtPlantingDate="&v_txtYear-1&"&vCROPTYPE=OP&vOwnership=ALL"
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010703" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&vstatus=TM&vCROPTYPE=OP&vOwnership=ALL"
	    end if
	    if rssql2("datagroupcode") = "REPGROUP010704" then
	     links = rssql2("string_link")&"?txtyear="&v_txtYear&"&txtmonth="&v_txtMonth&"&txtGroupRegion="&v_txtGroupRegion&"&vJenisReport=Group"
	    end if
	    'response.write "this is link = "&links 
%>	
			<tr>
						<td colspan="2"><a href =<%= links%>><%=rssql2("description")%></a></td>
						<td colspan="1"><%=rssql2("unitofmeasurement")%></td>
						<%
						
						if  v_txtGroupRegion = "--ALL--" then
						
						sqld1= "select "& rssql2("str_act_thismonth") &" actM,year,month,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and year = "& v_txtyear &" and month = "& v_txtmonth &  tempAdditionalString & " group by datagroupcode, year, month order by Year,Month"
						
						sqld2= "select "& rssql2("str_bud_thismonth") &" budM,year,month,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"''and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by datagroupcode, year, month order by Year,Month"
						
						sqld3= "select "& rssql2("str_act_todate") &" actD,year,month,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by datagroupcode, year, month order by Year,Month"
						
						sqld4= "select "& rssql2("str_bud_todate") &" budD,year,month,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by datagroupcode, year, month order by Year,Month"
						
						sqld5= "select "& rssql2("str_bud_annual") &" budA,year,month,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by datagroupcode, year, month order by Year,Month"
						
						
						
						sqlc1 = "select check_query2('"&sqld1&"') checkid from dual"
						sqlc2 = "select check_query2('"&sqld2&"') checkid from dual "
						sqlc3 = "select check_query2('"&sqld3&"') checkid from dual "
						sqlc4 = "select check_query2('"&sqld4&"') checkid from dual "
						sqlc5 = "select check_query2('"&sqld5&"') checkid from dual "
						
						     
						else
						
						
						sqld1= "select "& rssql2("str_act_thismonth") &" actM,year,month,groupregioncode,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and groupregioncode =''"& v_txtGroupRegion &"''  and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by groupregioncode,datagroupcode, year, month order by Year,Month"
						
						sqld2= "select "& rssql2("str_bud_thismonth") &" budM,year,month,groupregioncode,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and groupregioncode =''"& v_txtGroupRegion &"''  and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by groupregioncode,datagroupcode, year, month order by Year,Month"
						
						sqld3= "select "& rssql2("str_act_todate") &" actD,year,month,groupregioncode,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and groupregioncode =''"& v_txtGroupRegion &"''  and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by groupregioncode,datagroupcode, year, month order by Year,Month"
						
						sqld4= "select "& rssql2("str_bud_todate") &" budD,year,month,groupregioncode,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and groupregioncode =''"& v_txtGroupRegion &"''  and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by groupregioncode,datagroupcode, year, month order by Year,Month"
						
						sqld5= "select "& rssql2("str_bud_annual") &" budA,year,month,groupregioncode,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode=''"& rssql2("datagroupcode") &"'' and groupregioncode =''"& v_txtGroupRegion &"''  and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString & " group by groupregioncode,datagroupcode, year, month order by Year,Month"
						
						
						
						sqlc1 = "select check_query('"&sqld1&"') checkid from dual"
						sqlc2 = "select check_query('"&sqld2&"') checkid from dual "
						sqlc3 = "select check_query('"&sqld3&"') checkid from dual "
						sqlc4 = "select check_query('"&sqld4&"') checkid from dual "
						sqlc5 = "select check_query('"&sqld5&"') checkid from dual "
						
						end if
						
                        'response.Write(sqlc1)

						set rssqlc1 = Server.CreateObject("ADODB.Recordset")
						rssqlc1.open sqlc1,connstr, adOpenStatic
						
						
						if  rssqlc1("checkid") = "TRUE" then
						hs1 = rssql2("str_act_thismonth")
						else
						hs1 = 0
						end if
						rssqlc1.Close
						set rssqlc1 = nothing
						
						
						set rssqlc2 = Server.CreateObject("ADODB.Recordset")
						rssqlc2.open sqlc2,connstr, adOpenStatic
						
						if  rssqlc2("checkid") = "TRUE" then
						hs2 = rssql2("str_bud_thismonth")
						else
						hs2 = 0
						end if
						rssqlc2.Close
						set rssqlc2 = nothing
						
						set rssqlc3 = Server.CreateObject("ADODB.Recordset")
						rssqlc3.open sqlc3,connstr, adOpenStatic
						
						if  rssqlc3("checkid") = "TRUE" then
						hs3 = rssql2("str_act_todate")
						else
						hs3 = 0
						end if
						rssqlc3.Close
						set rssqlc3 = nothing
						
						set rssqlc4 = Server.CreateObject("ADODB.Recordset")
						rssqlc4.open sqlc4,connstr, adOpenStatic
						
						if  rssqlc4("checkid") = "TRUE" then
						hs4 = rssql2("str_bud_todate")
						else
						hs4 = 0
						end if
						rssqlc4.Close
						set rssqlc4 = nothing
						
						set rssqlc5 = Server.CreateObject("ADODB.Recordset")
						rssqlc5.open sqlc5,connstr, adOpenStatic
						
						if  rssqlc5("checkid") = "TRUE" then
						hs5 = rssql2("str_bud_annual")
						else
						hs5 = 0
						end if
						rssqlc5.Close
						set rssqlc5 = nothing
						
						
						
						if v_txtGroupRegion = "--ALL--" then
						
						sql3= "select "& hs1 &" actM , nvl("& hs2 &",0) budM , "& hs3 &" actD , "& hs4 &" budD , "& hs5 &" budA,year,month,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode='"& rssql2("datagroupcode") &"' and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString_01 & " group by datagroupcode, year, month order by Year,Month"
						
						else
						
						sql3= "select "& hs1 &" actM ,  nvl("& hs2 &",0) budM  , "& hs3 &" actD , "& hs4 &" budD , "& hs5 &" budA,year,month,groupregioncode,datagroupcode  from  MGT2_MONTHLY_RPT_RESULT_TAB where datagroupcode='"& rssql2("datagroupcode") &"' and groupregioncode ='"& v_txtGroupRegion &"'  and year = "& v_txtyear &" and month = "& v_txtmonth & tempAdditionalString_01 & " group by groupregioncode,datagroupcode, year, month order by Year,Month"
						
						end if
						
						'response.write sql3
						Set rssql3 = Server.CreateObject("ADODB.Recordset")
							rssql3.Open sql3, connstr, adOpenStatic
							
						if not rssql3.eof then
						
							 if isNull(rssql3("actM")) then 
							 persentase1 = 0
							 persentase2  = 0
							 else
							persentase1 = FormatNumber(Cdbl(rssql3("actM")),2 )
							persentase2 = FormatNumber(Cdbl(rssql3("budM")),2)
							end if
							if persentase1 <> 0  and persentase2 <> 0 then
							persentase = FormatNumber(cdbl(persentase1/persentase2)*100,2)
							else
							persentase = 0
							end if
							
							if isNull(rssql3("budA")) then 
							 persentasea= 0
							 else
							persentasea =FormatNumber(Cdbl(rssql3("budA")),2 )
							end if

							if isNull(rssql3("actD")) then 
							 persentased1 = 0
							 persentased2  = 0
							 else
							persentased1 = FormatNumber(Cdbl(rssql3("actD")),2 )
							persentased2 = FormatNumber(Cdbl(rssql3("budD")),2)
							end if
							
							if persentased1 <> 0 and persentased2 <> 0 then
							persentased = FormatNumber(cdbl(persentased1/persentased2)*100,2)
							else
							persentased = 0
							end if
							
							
						%>
						<td colspan="1" align="right"><%=FbpNum2(persentase1)%></td>
						<td colspan="1" align="right"><%=FbpNum2(persentase2)%></td>
						<td colspan="1" align="right"><%=FbpNum2(persentase)%></td>
						
						<td colspan="1" align="right"><%=FbpNum2(persentased1)%></td>
						<td colspan="1" align="right"><%=FbpNum2(persentased2)%></td>
						<td colspan="1" align="right"><%=FbpNum2(persentased)%></td>
						<td colspan="1" align="right"><%=FbpNum2(persentasea)%></td>
						<%
						if isNull(rssql3("actM")) then 
							actMTot = actMTot + 0
							budMTot = budMTot + 0
						else
							actMTot = actMTot + Cdbl(rssql3("actM"))
							budMTot = budMTot + Cdbl(rssql3("budM"))
							end if
							abtot1 = abtot1 + persentase
							abtot2 = abtot2 + persentased
							if isNull(rssql3("actD")) then 
							actDTot = actDTot + 0
							budDTot = budDTot + 0
						else
							actDTot = actDTot + Cdbl(rssql3("actD"))
							budDTot = budDTot + Cdbl(rssql3("budD"))
							end if
							
							if isNull(rssql3("budA")) then 
							budATot = budATot + 0
							else
							budATot = budATot + Cdbl(rssql3("budA"))
							
							end if
							
							else
							%>
							<td colspan="7" align="center">NO DATA</td>
							<%
						end if
						%>
			</tr>
			<%
			persentase = 0
			persentase1 =0
			persentase2 =0
			persentased =0
			persentased1 = 0
			persentased2 = 0
			rssql2.movenext
			loop
			end if
			rssql2.Close
			set rssql2 = nothing
			
			if rssql1("subgroupcode") = "REPGROUP0101" then
			
						%>
								
	   		<tr>
						<td colspan="2">Total Planted</td>
						<td colspan="1">HA</td>
						<td colspan="1" align="right"><%= FormatNumber(actMTot ,2)%></td>
						<td colspan="1" align="right"><%= FormatNumber(budMTot,2) %></td>
						<td colspan="1" align="right"><%= FormatNumber(abtot1,2) %>%</td>
						<td colspan="1" align="right"><%= FormatNumber(actDTot ,2)%></td>
						<td colspan="1" align="right"><%= FormatNumber(budDTot,2) %></td>
						<td colspan="1" align="right"><%= FormatNumber(abtot2,2)%>%</td>
						<td colspan="1" align="right"><%= FormatNumber(budATot,2)%></td>
					</tr>
						<%
						end if
				rssql1.movenext
				
			loop
			rssql1.Close
			set rssql1 = nothing
			
						else
						%>		
	   		<tr>
						<td colspan="<%=182%>">NO DATA AVAILABLE</td>
					</tr>
				<%
				
						END IF 
				%>

				</tbody>
			</table>

		</div>
	</body>
</html>
