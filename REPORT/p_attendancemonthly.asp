<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,txtDate, txtJob_forfieldstatus
	Dim strMain, nSpanTotal
	Dim rsAttDaily, strAttDaily,  v_txtJob_forfieldstatus, v_txtfcba, strsql
	Dim v_txtDate, NO, cClass
	Dim sprevFcba,SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	
	
	
	v_txtDate				=	request("txtDate")
	v_txtJob_forfieldstatus	=	request("txtJob_forfieldstatus")
	v_txtfcba 			    =   Session("fcba")
	
	nSpanTotal = 6
	
	if v_txtDate ="" then
	   v_txtDate= textdate(Now())
    end if	
	
	v_txtyear 		= Request("txtyear")
    v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	
	if v_txtyear = "" then
		v_txtyear = year(date())
	end if

	 	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
				rsCat.Close
				set rsCat = nothing
			   
	  strMain = " SELECT fcba,division,groupakun,groupname, " & _ 
				" Nvl(jan_khl,0)jan_khl,Nvl(feb_khl,0)feb_khl,Nvl(mar_khl,0)mar_khl,Nvl(apr_khl,0)apr_khl,Nvl(may_khl,0)may_khl,Nvl(jun_khl,0)jun_khl,Nvl(jul_khl,0)jul_khl,Nvl(aug_khl,0)aug_khl,Nvl(sep_khl,0)sep_khl,Nvl(oct_khl,0)oct_khl,Nvl(Nov_khl,0)Nov_khl,Nvl(des_khl,0) des_khl," & _
				" Nvl(jan_kht,0)jan_kht,Nvl(feb_kht,0)feb_kht,Nvl(mar_kht,0)mar_kht,Nvl(apr_kht,0)apr_kht,Nvl(may_kht,0)may_kht,Nvl(jun_kht,0)jun_kht,Nvl(jul_kht,0)jul_kht,Nvl(aug_kht,0)aug_kht,Nvl(sep_kht,0)sep_kht,Nvl(oct_kht,0)oct_kht,Nvl(Nov_kht,0)Nov_kht,Nvl(des_kht,0) des_kht," & _
				" Nvl(jan_kbt,0)jan_kbt,Nvl(feb_kbt,0)feb_kbt,Nvl(mar_kbt,0)mar_kbt,Nvl(apr_kbt,0)apr_kbt,Nvl(may_kbt,0)may_kbt,Nvl(jun_kbt,0)jun_kbt,Nvl(jul_kbt,0)jul_kbt,Nvl(aug_kbt,0)aug_kbt,Nvl(sep_kbt,0)sep_kbt,Nvl(oct_kbt,0)oct_kbt,Nvl(Nov_kbt,0)Nov_kbt,Nvl(des_kbt,0) des_kbt," & _
				" Nvl(jan_total,0)jan_total,Nvl(feb_total,0)feb_total,Nvl(mar_total,0)mar_total,Nvl(apr_total,0)apr_total,Nvl(may_total,0)may_total,Nvl(jun_total,0)jun_total,Nvl(jul_total,0)jul_total,Nvl(aug_total,0)aug_total,Nvl(sep_total,0)sep_total,Nvl(oct_total,0)oct_total,Nvl(Nov_total,0)Nov_total,Nvl(des_total,0) des_total," & _
				" Nvl(jan_khl,0)+Nvl(feb_khl,0)+Nvl(mar_khl,0)+Nvl(apr_khl,0)+Nvl(may_khl,0)+Nvl(jun_khl,0)+Nvl(jul_khl,0)+Nvl(aug_khl,0)+Nvl(sep_khl,0)+Nvl(oct_khl,0)+Nvl(Nov_khl,0)+Nvl(des_khl,0) sbi_khl," & _
				" Nvl(jan_kht,0)+Nvl(feb_kht,0)+Nvl(mar_kht,0)+Nvl(apr_kht,0)+Nvl(may_kht,0)+Nvl(jun_kht,0)+Nvl(jul_kht,0)+Nvl(aug_kht,0)+Nvl(sep_kht,0)+Nvl(oct_kht,0)+Nvl(Nov_kht,0)+Nvl(des_kht,0) sbi_kht," & _
				" Nvl(jan_kbt,0)+Nvl(feb_kbt,0)+Nvl(mar_kbt,0)+Nvl(apr_kbt,0)+Nvl(may_kbt,0)+Nvl(jun_kbt,0)+Nvl(jul_kbt,0)+Nvl(aug_kbt,0)+Nvl(sep_kbt,0)+Nvl(oct_kbt,0)+Nvl(Nov_kbt,0)+Nvl(des_kbt,0) sbi_kbt" & _
				" FROM " & _
				" (" & _
				" (" & _
				" SELECT " & _
				" businessunit_api.get_groupregion(a.fcba ) region," & _
				" businessunit_api.get_companycode(a.fcba ) company," & _
				" a.fcba," & _
				" b.sectioncode division," & _
				" f.YEAR," & _
				" f.month," & _
				" e.fccode groupakun," & _
				" e.fcname groupname," & _
				" c.presentgrade," & _
				" Sum(a.mandays) hk" & _
				" FROM bkmupahdtl a " & _
				" INNER JOIN bkmhdr b ON (a.bkmno=b.bkmno AND a.fcba=b.fcba) " & _
				" INNER JOIN employee_hist_tab c  ON (a.employeecode=c.fccode AND a.fcba=c.fcba AND extract(YEAR FROM bkmdate)=c.YEAR AND  extract(month FROM bkmdate)=c.month)" & _
				" INNER JOIN jobmaster d ON (a.actcode=d.fccode)" & _
				" INNER JOIN jobgroupmaster e ON (d.jobgroupcode=e.fccode)" & _
				" INNER JOIN empalldedstatus f ON (a.fcba=f.fcba and b.bkmdate between start_date AND end_date)" & _ 
				" WHERE f.fcba='" & v_txtfcba & "' AND   f.year =" & v_txtyear & " AND f.month between 1 AND " & v_txtmonth & "" & _
				" AND a.attendance IN ('KJ','DL','TG')" & _
				" AND a.mandays > 0" & _
				" GROUP BY " & _
				" a.fcba," & _
				" b.sectioncode," & _
				" f.YEAR  ," & _
				" f.month ," & _
				" e.fccode," & _
				" e.fcname," & _
				" c.presentgrade" & _
				" )" & _
				" PIVOT" & _
				" (" & _
				"   Sum(CASE WHEN presentgrade = 'KBT'  THEN hk ELSE 0 end) KBT," & _
				"   Sum(CASE WHEN presentgrade = 'KHT'  THEN hk ELSE 0 end) KHT," & _
				"   Sum(CASE WHEN presentgrade = 'KHL'  THEN hk ELSE 0 end) KHL," & _
				"   Sum(hk) Total" & _
				"   FOR MONTH IN " & _
				" 	( " & _
				" 	  1 jan," & _
				" 	  2 feb," & _
				" 	  3 mar," & _
				" 	  4 apr," & _
				" 	  5 may," & _
				" 	  6 jun," & _
				" 	  7 jul," & _
				" 	  8 aug," & _
				" 	  9 sep," & _
				" 	  10 oct," & _
				" 	  11 nov," & _
				" 	  12 des" & _
				" 	)" & _
				" )" & _
				" )s" & _
				" ORDER BY division,groupakun"


	  
	 				
	 

    'response.write strMain
	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strMain, connstr, adOpenStatic 
	
%>


<html>

<head>

<title>FB-Plans Reporting - Laporan Harian Tenaga Kerja</title>
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

<!--<script type="text/javascript" src="http://jqueryjs.googlecode.com/files/jquery-1.3.1.min.js" > </script> -->
<script type="text/javascript">

    function PrintElem(elem)
    {
        Popup($(elem).html());
    }

    function Popup(data) 
    {
        var mywindow = window.open('', 'my div', 'height=400,width=600');
        mywindow.document.write('<html><head><title>my div</title>');
        mywindow.document.write('<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />');
        mywindow.document.write('</head><body >');
        mywindow.document.write(data);
        mywindow.document.write('</body></html>');

        mywindow.print();
        mywindow.close();

        return true;
    }

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
</head>

<body>
<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	</div>
<div class="top_select" style=" position:absolute;  top:0; left:0px; width:100%;"> 
<form name="frmEntry" method="post" action="p_Attendancemonthly.asp">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	
	
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>&nbsp;</th>
		<th>BULAN</th>
		<th>TAHUN</th>
		<th>&nbsp;</th>
	</tr>
	<tr>		
	
		<td align="center">
		<% 
					MM=Request("txtMonth")
					yyyy=Request("txtYear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = 2017
					MaxYr = Yr  + 3
					If len(yyyy)= 0 then yyyy = Yr
				%>
		
		   <td>
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
		   <td>
				<font size="2"  color="#0000FF" ></font>
				
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

		
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
   <tr>
<td colspan="140" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
	<div id="header" class="centered">
			 <% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
			 <% End if%>
			 <% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
			 <% End if%>
	</div>
	<div id="subheader"class="fixed-1 clearfix centered">
			<div id="subheader" class="centered">
				<h2>LAPORAN BULANAN TENAGA KERJA</h3>
				<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> </h5>
			</div>
	</div>
</div>
</td>
</tr>
</table>
</div>



<div id="mydiv" style="margin-top:280px;">
<table border="0" width="100%" id="product-table" class="xthf-site" style="background-color:#FFFFFF;">
<thead>
	<tr>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Divisi</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Aktivitas</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >JAN</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >FEB</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >MAR</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >APR</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >MAY</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >JUN</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >JUL</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >AUG</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >SEP</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >OCT</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >NOV</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >DES</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >SBI</th>
		
	</tr>
	<tr>
	    <% DIM K
		   FOR K = 1 TO 13 %>
			<th style="background-color:#00A300;color:#fff;">KBT</th>
			<th style="background-color:#00A300;color:#fff;">KHT</th>
			<th style="background-color:#00A300;color:#fff;">KHL</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<% NEXT %>	
		
		
	</tr>
        </thead>
 <tbody>
<%
Dim nkbt_01, nkht_01, nkhl_01, nTotal_01,nkbt_02, nkht_02, nkhl_02, nTotal_02,nkbt_03, nkht_03, nkhl_03, nTotal_03,nkbt_04, nkht_04, nkhl_04, nTotal_04
Dim nkbt_05, nkht_05, nkhl_05, nTotal_05,nkbt_06, nkht_06, nkhl_06, nTotal_06,nkbt_07, nkht_07, nkhl_07, nTotal_07,nkbt_08, nkht_08, nkhl_08, nTotal_08
Dim nkbt_09, nkht_09, nkhl_09, nTotal_09,nkbt_10, nkht_10, nkhl_10, nTotal_10,nkbt_11, nkht_11, nkhl_11, nTotal_11,nkbt_12, nkht_12, nkhl_12, nTotal_12
Dim nkbt_SBI, nkht_SBI, nkhl_SBI, nTotal_SBI,nTotal_BUD
Dim Total_SBI
Dim nkbt_1, nkht_1, nkhl_1, nTotal_1,nsbi,nshi
Dim nkbt_2, nkht_2, nkhl_2, nTotal_2
Dim sprevjob_forfieldstatus, tempjob_forfieldstatus
	if not rsAttDaily.EOF then
		NO = 1
		
		
		sprevFcba = ""
		sprevjob_forfieldstatus	=	""
		
		while not rsAttDaily.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="left"><%=rsAttDaily("division")%></td>
				<td align="left"><%=rsAttDaily("groupname")%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jan_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jan_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jan_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jan_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("feb_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("feb_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("feb_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("feb_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("mar_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("mar_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("mar_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("mar_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("apr_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("apr_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("apr_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("apr_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("may_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("may_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("may_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("may_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("jun_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jun_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jun_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jun_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("jul_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jul_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jul_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("jul_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("aug_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("aug_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("aug_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("aug_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("sep_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("sep_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("sep_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("sep_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("oct_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("oct_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("oct_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("oct_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("nov_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("nov_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("nov_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("nov_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("des_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("des_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("des_khl"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("des_total"))%></td>
				
				<td align="right"><%=FbpNum2(rsAttDaily("sbi_kbt"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("sbi_kht"))%></td>
				<td align="right"><%=FbpNum2(rsAttDaily("sbi_khl"))%></td>
				<% total_SBI = cdbl(rsAttDaily("sbi_kbt")) + cdbl(rsAttDaily("sbi_kht")) + cdbl(rsAttDaily("sbi_khl")) %>
				<td align="right"><%=FbpNum2(total_SBI)%></td>
				
			</tr>
<%
			NO = NO + 1			
			nkbt_01 = nkbt_01 + Cdbl(rsAttDaily("jan_kbt"))
			nkht_01 = nkht_01 + Cdbl(rsAttDaily("jan_kht"))
			nkhl_01 = nkhl_01 + Cdbl(rsAttDaily("jan_khl"))
			nTotal_01 = nTotal_01 + Cdbl(rsAttDaily("jan_total"))
			
			nkbt_02 = nkbt_02 + Cdbl(rsAttDaily("feb_kbt"))
			nkht_02 = nkht_02 + Cdbl(rsAttDaily("feb_kht"))
			nkhl_02 = nkhl_02 + Cdbl(rsAttDaily("feb_khl"))
			nTotal_02 = nTotal_02 + Cdbl(rsAttDaily("feb_total"))
			
			nkbt_03 = nkbt_03 + Cdbl(rsAttDaily("mar_kbt"))
			nkht_03 = nkht_03 + Cdbl(rsAttDaily("mar_kht"))
			nkhl_03 = nkhl_03 + Cdbl(rsAttDaily("mar_khl"))
			nTotal_03 = nTotal_03 + Cdbl(rsAttDaily("mar_total"))
			
			nkbt_04 = nkbt_04 + Cdbl(rsAttDaily("apr_kbt"))
			nkht_04 = nkht_04 + Cdbl(rsAttDaily("apr_kht"))
			nkhl_04 = nkhl_04 + Cdbl(rsAttDaily("apr_khl"))
			nTotal_04 = nTotal_04 + Cdbl(rsAttDaily("apr_total"))
			
			nkbt_05 = nkbt_05 + Cdbl(rsAttDaily("may_kbt"))
			nkht_05 = nkht_05 + Cdbl(rsAttDaily("may_kht"))
			nkhl_05 = nkhl_05 + Cdbl(rsAttDaily("may_khl"))
			nTotal_05 = nTotal_05 + Cdbl(rsAttDaily("may_total"))
			
			nkbt_06 = nkbt_06 + Cdbl(rsAttDaily("jun_kbt"))
			nkht_06 = nkht_06 + Cdbl(rsAttDaily("jun_kht"))
			nkhl_06 = nkhl_06 + Cdbl(rsAttDaily("jun_khl"))
			nTotal_06 = nTotal_06 + Cdbl(rsAttDaily("jun_total"))
			
			nkbt_07 = nkbt_07 + Cdbl(rsAttDaily("jul_kbt"))
			nkht_07 = nkht_07 + Cdbl(rsAttDaily("jul_kht"))
			nkhl_07 = nkhl_07 + Cdbl(rsAttDaily("jul_khl"))
			nTotal_07 = nTotal_07 + Cdbl(rsAttDaily("jul_total"))
			
			nkbt_08 = nkbt_08 + Cdbl(rsAttDaily("aug_kbt"))
			nkht_08 = nkht_08 + Cdbl(rsAttDaily("aug_kht"))
			nkhl_08 = nkhl_08 + Cdbl(rsAttDaily("aug_khl"))
			nTotal_08 = nTotal_08 + Cdbl(rsAttDaily("aug_total"))
			
			nkbt_09 = nkbt_09 + Cdbl(rsAttDaily("sep_kbt"))
			nkht_09 = nkht_09 + Cdbl(rsAttDaily("sep_kht"))
			nkhl_09 = nkhl_09 + Cdbl(rsAttDaily("sep_khl"))
			nTotal_09 = nTotal_09 + Cdbl(rsAttDaily("sep_total"))
			
			nkbt_10 = nkbt_10 + Cdbl(rsAttDaily("oct_kbt"))
			nkht_10 = nkht_10 + Cdbl(rsAttDaily("oct_kht"))
			nkhl_10 = nkhl_10 + Cdbl(rsAttDaily("oct_khl"))
			nTotal_10 = nTotal_10 + Cdbl(rsAttDaily("oct_total"))
			
			nkbt_11 = nkbt_11 + Cdbl(rsAttDaily("nov_kbt"))
			nkht_11 = nkht_11 + Cdbl(rsAttDaily("nov_kht"))
			nkhl_11 = nkhl_11 + Cdbl(rsAttDaily("nov_khl"))
			nTotal_11 = nTotal_11 + Cdbl(rsAttDaily("nov_total"))
			
			nkbt_12 = nkbt_12 + Cdbl(rsAttDaily("des_kbt"))
			nkht_12 = nkht_12 + Cdbl(rsAttDaily("des_kht"))
			nkhl_12 = nkhl_12 + Cdbl(rsAttDaily("des_khl"))
			nTotal_12 = nTotal_12 + Cdbl(rsAttDaily("des_total"))
			
			nkbt_SBI = nkbt_SBI + Cdbl(rsAttDaily("sbi_kbt"))
			nkht_SBI = nkht_SBI + Cdbl(rsAttDaily("sbi_kht"))
			nkhl_SBI = nkhl_SBI + Cdbl(rsAttDaily("sbi_khl"))
			nTotal_SBI = nTotal_SBI + cdbl(total_SBI)
			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 2%>">TOTAL</td>
				<td align="right"><%=FormatNumber(nkbt_01,2)%></td>
				<td align="right"><%=FormatNumber(nkht_01,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_01,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_01,2)%></td>	
				
				<td align="right"><%=FormatNumber(nkbt_02,2)%></td>
				<td align="right"><%=FormatNumber(nkht_02,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_02,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_02,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_03,2)%></td>
				<td align="right"><%=FormatNumber(nkht_03,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_03,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_03,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_04,2)%></td>
				<td align="right"><%=FormatNumber(nkht_04,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_04,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_04,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_05,2)%></td>
				<td align="right"><%=FormatNumber(nkht_05,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_05,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_05,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_06,2)%></td>
				<td align="right"><%=FormatNumber(nkht_06,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_06,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_06,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_07,2)%></td>
				<td align="right"><%=FormatNumber(nkht_07,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_07,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_07,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_08,2)%></td>
				<td align="right"><%=FormatNumber(nkht_08,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_08,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_08,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_09,2)%></td>
				<td align="right"><%=FormatNumber(nkht_09,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_09,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_09,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_10,2)%></td>
				<td align="right"><%=FormatNumber(nkht_10,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_10,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_10,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_11,2)%></td>
				<td align="right"><%=FormatNumber(nkht_11,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_11,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_11,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_12,2)%></td>
				<td align="right"><%=FormatNumber(nkht_12,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_12,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_12,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_SBI,2)%></td>
				<td align="right"><%=FormatNumber(nkht_SBI,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_SBI,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_SBI,2)%></td>
				
				<td align="right"><%=FormatNumber(nTotal_BUD,2)%></td>
				
			</tr>
<tbody>
</table>
</form>
</div>
</body>
</html>