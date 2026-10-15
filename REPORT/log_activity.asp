<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
    <title>ePlaS - Chart</title>
    <link rel="StyleSheet" href="dtree.css" type="text/css" />
	<link rel="StyleSheet" href="style-content.css" type="text/css" />
	<link class="include" rel="stylesheet" type="text/css" href="dist/jquery.jqplot.min.css" />
    <link rel="stylesheet" type="text/css" href="dist/examples/examples.min.css" />
	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
    <!--[if lt IE 9]><script language="javascript" type="text/javascript" src="dist/excanvas.js"></script><![endif]-->
    <script class="include" type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>

	 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<style>
.none_style { border:none; padding:none;}
.jqplot-image-button { font-size:11px;}
table.data tr:nth-of-type(odd) {
  background-color:#f5fafa;
}
table.data tr {
  background-color:#fff;
}
.minion {}
</style>
</head>

<body>
        <!--#include file="top.jsp"-->
<% if Session("fcBa") <> "" then %>		
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
        <!--#include file="menu.asp"-->
	</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:100%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	  Module :REPORT LOG-USER ACTIVITY ( WEEKLY )
	</div>

<div class="box_both" style="overflow:scroll;">
  
<!-- Example scripts go here -->
<%
Dim txtJUMAT,txtJUMAT_tgl
Dim txtSABTU,txtSABTU_tgl
Dim txtSENIN,txtSENIN_tgl
Dim txtSELASA,txtSELASA_tgl
Dim txtRABU,txtRABU_tgl
Dim txtKAMIS,txtKAMIS_tgl
Dim txtPRDATE_From,txtPRDATE_to

txtPRDATE_From =Request("PRDATE_From")
txtPRDATE_to =Request("PRDATE_to")
	'---- Data Header
 if Request("PRDATE_to")<>"" then
   txtPRDATE_to   =Request("PRDATE_to")
   txtPRDATE_From   = CDATE(Request("PRDATE_to")) - 6
 else
   txtPRDATE_to   =textDate(date)
   txtPRDATE_From   =textDate(date-6)
 end if

stringFRIDAY = "SELECT count(to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY')) JUMAT, to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd') TANGGAL "&_
			"FROM (select * from historical_user_tab  where username in (select trim(FCCODE)  from ( "&_
			"SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM USERS_TAB A, USER_ROLES_TAB B "&_
			"WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN ('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			"union "&_
			"select * from historical_user_tab  where username in "&_
			"(select trim(FCCODE)from (select FCCODE ,FULLNAME, DEPARTMENT from USERS_TAB A  where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "&_
			"WHERE to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY') LIKE '%FRIDAY%'"

if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringFRIDAY	=	stringFRIDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	else
	stringFRIDAY	=	stringFRIDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	end if
			
 Set rsUoM = Server.CreateObject("ADODB.Recordset")
					  rsUoM.Open stringFRIDAY, connstr
					  IF Not rsUoM.eof then
						 txtJUMAT = rsUoM("JUMAT")
						 txtJUMAT_tgl = rsUoM("TANGGAL")
					  End if
					  rsUoM.Close
					  set rsUoM = Nothing  
					  
stringMONDAY = "SELECT count(to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY')) SENIN, to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd') TANGGAL "&_
			"FROM (select * from historical_user_tab  where username in (select trim(FCCODE)  from ( "&_
			"SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM USERS_TAB A, USER_ROLES_TAB B "&_
			"WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN ('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			"union "&_
			"select * from historical_user_tab  where username in "&_
			"(select trim(FCCODE)from (select FCCODE ,FULLNAME, DEPARTMENT from USERS_TAB A  where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "&_
			"WHERE to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY') LIKE '%MONDAY%'"

 if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringMONDAY	=	stringMONDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	else
	stringMONDAY	=	stringMONDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	end if
	
 Set rsUoM = Server.CreateObject("ADODB.Recordset")
					  rsUoM.Open stringMONDAY, connstr
					  IF Not rsUoM.eof then
					     txtSENIN = rsUoM("SENIN")
						 txtSENIN_tgl = rsUoM("TANGGAL")
					  End if
					  rsUoM.Close
					  set rsUoM = Nothing 
					  
stringSATURDAY = "SELECT count(to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY')) SABTU, to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd') TANGGAL "&_
			"FROM (select * from historical_user_tab  where username in (select trim(FCCODE)  from ( "&_
			"SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM USERS_TAB A, USER_ROLES_TAB B "&_
			"WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN ('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			"union "&_
			"select * from historical_user_tab  where username in "&_
			"(select trim(FCCODE)from (select FCCODE ,FULLNAME, DEPARTMENT from USERS_TAB A  where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "&_
			"WHERE to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY') LIKE '%SATURDAY %'"
 if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringSATURDAY	=	stringSATURDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	else
	stringSATURDAY	=	stringSATURDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	end if
	
 Set rsUoM = Server.CreateObject("ADODB.Recordset")
					  rsUoM.Open stringSATURDAY , connstr
					  IF Not rsUoM.eof then
					     txtSABTU  = rsUoM("SABTU")
						 txtSABTU_tgl = rsUoM("TANGGAL")
					  End if
					  rsUoM.Close
					  set rsUoM = Nothing 
stringTHURSDAY  = "SELECT count(to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY')) KAMIS, to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd') TANGGAL "&_
			"FROM (select * from historical_user_tab  where username in (select trim(FCCODE)  from ( "&_
			"SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM USERS_TAB A, USER_ROLES_TAB B "&_
			"WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN ('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			"union "&_
			"select * from historical_user_tab  where username in "&_
			"(select trim(FCCODE)from (select FCCODE ,FULLNAME, DEPARTMENT from USERS_TAB A  where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "&_
			"WHERE to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY') LIKE '%THURSDAY%'"
			
 if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringTHURSDAY	=	stringTHURSDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	else
	stringTHURSDAY	=	stringTHURSDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	end if
			
 Set rsUoM = Server.CreateObject("ADODB.Recordset")
					  rsUoM.Open stringTHURSDAY  , connstr
					  IF Not rsUoM.eof then
					     txtKAMIS  = rsUoM("KAMIS")
						 txtKAMIS_tgl = rsUoM("TANGGAL")
					  End if
					  rsUoM.Close
					  set rsUoM = Nothing 
stringTUESDAY  = "SELECT count(to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY')) SELASA, to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd') TANGGAL "&_
			"FROM (select * from historical_user_tab  where username in (select trim(FCCODE)  from ( "&_
			"SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM USERS_TAB A, USER_ROLES_TAB B "&_
			"WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN ('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			"union "&_
			"select * from historical_user_tab  where username in "&_
			"(select trim(FCCODE)from (select FCCODE ,FULLNAME, DEPARTMENT from USERS_TAB A  where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "&_
			"WHERE to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY') LIKE '%TUESDAY%'"
			
 if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringTUESDAY	=	stringTUESDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"/02/2013' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	else
	stringTUESDAY	=	stringTUESDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	end if			
	
 Set rsUoM = Server.CreateObject("ADODB.Recordset")
					  rsUoM.Open stringTUESDAY , connstr
					  IF Not rsUoM.eof then
					     txtSELASA = rsUoM("SELASA")
						 txtSELASA_tgl = rsUoM("TANGGAL")
					  End if
					  rsUoM.Close
					  set rsUoM = Nothing 
stringWEDNESDAY  = "SELECT count(to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY')) RABU, to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd') TANGGAL "&_
			"FROM (select * from historical_user_tab  where username in (select trim(FCCODE)  from ( "&_
			"SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM USERS_TAB A, USER_ROLES_TAB B "&_
			"WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN ('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			"union "&_
			"select * from historical_user_tab  where username in "&_
			"(select trim(FCCODE)from (select FCCODE ,FULLNAME, DEPARTMENT from USERS_TAB A  where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "&_
			"WHERE to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'DAY') LIKE '%WEDNESDAY%' "
			
 if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringWEDNESDAY	=	stringWEDNESDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	else
	stringWEDNESDAY	=	stringWEDNESDAY & "and TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' group by to_char(to_date(DATE_ACT,'day,Mon dd,yyyy'),'dd')"
	end if
	
 Set rsUoM = Server.CreateObject("ADODB.Recordset")
					  rsUoM.Open stringWEDNESDAY , connstr
					  IF Not rsUoM.eof then
					     txtRABU = rsUoM("RABU")
						 txtRABU_tgl = rsUoM("TANGGAL")
					  End if
					  rsUoM.Close
					  set rsUoM = Nothing  


stringUSER  = "SELECT USERNAME, COUNT(USERNAME) AKTIFITAS FROM (select * from historical_user_tab  where username in ( "&_
			  "select trim(FCCODE)  from (SELECT A.FCCODE ,A.FULLNAME, A.DEPARTMENT,A.ACCESSRIGHTS, B.FCROLESCODE, B.ISACTIVE FROM "&_
			  "USERS_TAB A, USER_ROLES_TAB B WHERE A.FCCODE=B.FCUSERCODE  AND trim(DEPARTMENT) IN "&_
			  "('ACC','PSD','HO','HP HO','PS','PURCHASING HO','RND','POM Department'))) "&_
			  "union "&_
			  "select * from historical_user_tab where username in (select trim(FCCODE) from (select FCCODE ,FULLNAME, DEPARTMENT "&_
			  "from USERS_TAB A where fccode in ('ANDY','DIDIK','ADITA','SUMANTO')))) "
			  
if  LEN(TRIM(txtPRDATE_to)) > 0 then
	stringUSER	=	stringUSER & "WHERE TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& txtPRDATE_From &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& txtPRDATE_to &"' GROUP BY USERNAME ORDER BY USERNAME"
	else
	stringUSER	=	stringUSER & "WHERE TO_CHAR(LASTUPDATE,'MM/DD/YYYY') >= '"& textDate(date-6) &"' and TO_CHAR(LASTUPDATE,'MM/DD/YYYY')<= '"& textDate(date) &"' GROUP BY USERNAME ORDER BY USERNAME"
	end if			  
	
%>
 
 

 <table style="float:right; position:relative; z-index:1; top:2.5px; right:10px;">
 <tr>
 <form class="none_style" action="table.Asp" method="post" name="eForm">
 <td>
  <font style="font-size:10px;">FROM SUNDAY : </font></td><td><input class="datepicker" type="text" maxlength="10" name="PRDATE_From" size="10" value="<%=txtPRDATE_From%>" disabled>
  </td>
  <td>
  <font style="font-size:10px;">TO SATURDAY : </font></td><td><input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=txtPRDATE_to%>">
  </td>
  <td><input type="Submit" name="View" value="View"></td>
  </form> 
  </tr>
  
  </table>
  
  
  <div id="tabs">
    <ul>
      <li><a href="#tabs-1">Data Online</a></li>
      <li><a href="#tabs-2">Grafik Chart</a></li>
      <li><a href="#tabs-3">Pie Chart</a></li>
    </ul>
    <div id="tabs-1">

     <table border="0" id="product-table" class="xthf-site data" style="font-size: 11px; text-align:left; float:right; width:40%; margin-top:15px;margin-bottom:15px;margin-right:2%;">
	 <thead>
	 <tr >
	 <th style="border-bottom:1px #ccc solid;">HARI</th><th tyle="border-bottom:1px #ccc solid;"></th><th style="border-bottom:1px #ccc solid;" align="center">JUMLAH PEMAKAI</th>
	 </tr>
	 </thead>
	<tbody>
	 <tr>
	 <td>SENIN</td><td>:</td><td align="center"><%=txtSENIN%>&nbsp;User</td>
	 </tr>
	 <tr>
	 <td>SELASA</td><td>:</td><td align="center"><%=txtSELASA%>&nbsp;User</td>
	 </tr>
	 <tr>
	 <td>RABU</td><td>:</td><td align="center"><%=txtRABU%>&nbsp;User</td>
	 </tr>
	 <tr>
	 <td>KAMIS</td><td>:</td><td align="center"><%=txtKAMIS%>&nbsp;User</td>
	 </tr>
	 <tr>
	 <td>JUMAT</td><td>:</td><td align="center"><%=txtJUMAT%>&nbsp;User</td>
	 </tr>
	 <tr>
	 <td>SABTU</td><td>:</td><td align="center"><%=txtSABTU%>&nbsp;User</td>
	 </tr>
	 
	 </table>
	 <table border="0" id="product-table" class="xthf-site data" style="font-size: 11px; text-align:left; float:left; width:50%; margin-top:15px;margin-bottom:15px;margin-left:2%;">
	 <thead>
		<tr>
		<th style="border-bottom:1px #ccc solid; width:15px;">No</th><th style="border-bottom:1px #ccc solid;">Nama</th><th style="width:50px; border-bottom:1px #ccc solid;">Jumlah Aktifitas</th>
		</tr>
	</thead>
	<tbody>
	<%	
	NO=1
	Set rsUSER = connStr.Execute(stringUSER)
	Dim v_Saldo

	v_Saldo  = 0
	While Not rsUSER.EOF
	
	%>
		<tr>
		<td><%=NO%></td><td><%=rsUSER("USERNAME")%></td><td align="right"><%=FbpNum2(rsUSER("AKTIFITAS"))%></td>
		</tr>
	<%
	v_Saldo  =v_Saldo+trim(cdbl(rsUSER("AKTIFITAS")))
	rsUSER.MoveNext
	NO=NO+1
	Wend
	rsUSER.close()
	rsUSER=null
	
	%>
	<tr class="nothover">
	<td align="left" style="background-color:#FFCC00; font-weight:bold;<%=cmdstyleLeft%>" colspan="2"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="background-color:#FFCC00;<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_Saldo%></td>
	</tr>
	
	</tbody>
	 
	 </table>
	 
	 <div style="clear:both; content:' ';"> </div>
    </div>
    
    <div id="tabs-2">
      <div id="chart1" data-height="260px" data-width="480px" style="margin-top:20px; margin-left:20px;">
	   
	   
	  </div>
    </div>
 
    <div id="tabs-3">
       <div id="chart2" style="margin-top:20px; margin-left:20px;"></div>
    </div>
    
  </div> 
   
<% 
dim SPACE
SPACE = "  "
if txtSENIN = "" then
 txtSENIN = "0"
 else
 txtSENIN = txtSENIN
end if

if txtSELASA = "" then
 txtSELASA = "0"
 else
 txtSELASA = txtSELASA
end if

if txtRABU = "" then
 txtRABU = "0"
 else
 txtRABU = txtRABU
end if
if txtKAMIS = "" then
 txtKAMIS = "0"
 else
 txtKAMIS = txtKAMIS
end if
if txtJUMAT = "" then
 txtJUMAT = "0"
 else
 txtJUMAT = txtJUMAT
end if
if txtSABTU = "" then
 txtSABTU = "0"
 
 else
 txtSABTU = txtSABTU
end if
%>


<script type="text/javascript">
    $(document).ready(function() {
        $.jqplot.config.enablePlugins = false;
		var ticks = ['Senin (<%=txtSENIN_tgl%>)', 'Selasa (<%=txtSELASA_tgl%>)', 'Rabu (<%=txtRABU_tgl%>)', 'Kamis (<%=txtKAMIS_tgl%>)', 'Jumat (<%=txtJUMAT_tgl%>)', 'Sabtu (<%=txtSABTU_tgl%>)'];
		var TT = "<%=Day(txtPRDATE_From)%> - <%=Day(txtPRDATE_to)%><%=SPACE%><%=MonthName(Month(txtPRDATE_to))%> <%=SPACE%> <%=Year(txtPRDATE_to)%>";
        var l1 = [<%=txtSENIN%>,<%=txtSELASA%>,<%=txtRABU%>,<%=txtKAMIS%>,<%=txtJUMAT%>,<%=txtSABTU%>];
        var l2 = [[2, 14], [7, 2], [8,5]];
        var l3 = [4, 7, 9, 2, 11, 5, 9, 13, 8, 7];
        var l4 = [<%Set rsUSER = connStr.Execute(stringUSER)
	While Not rsUSER.EOF
	%>
	['<%=rsUSER("USERNAME")%>',<%=rsUSER("AKTIFITAS")%>],
<%
	rsUSER.MoveNext 
	Wend
	rsUSER.close()
	rsUSER=null%>];

        $("#tabs").tabs();
        $("#accordion").accordion();

        var plot1 = $.jqplot('chart1', [l1],  {
          title: "Jumlah User Login, " + TT,
          lengend:{show:true},
		  animate: true,
			// Will animate plot on calls to plot1.replot({resetAxes:true})
			animateReplot: true,
			cursor: {
				show: true,
				zoom: true,
				looseZoom: true,
				showTooltip: false
			},
          seriesDefaults:{
                renderer:$.jqplot.BarRenderer,
                pointLabels: { show: true }
            },
            axes: {
                xaxis: {
                    renderer: $.jqplot.CategoryAxisRenderer,
                    ticks: ticks
                }
            },
          cursor:{show:true, zoom:true},
          axesDefaults:{useSeriesColor:true, rendererOptions: { alignTicks: true}}
        });
		var plot2 = jQuery.jqplot ('chart2', [l4], 
    { 
	  title: "Aktifitas User, " + TT,
	  height: 300,
      width: 500,
      seriesDefaults: {
        // Make this a pie chart.
        renderer: jQuery.jqplot.PieRenderer, 
        rendererOptions: {
          // Put data labels on the pie slices.
          // By default, labels show the percentage of the slice.
          showDataLabels: true
        }
      }, 
      legend: { show:true, location: 'e' }
    });

        

        $('#tabs').bind('tabsactivate', function(event, ui) {
          if (ui.newTab.index() === 1 && plot1._drawCount === 0) {
            plot1.replot();
          }
          else if (ui.newTab.index() === 2 && plot2._drawCount === 0) {
            plot2.replot();
          }
        });

        $('#accordion').bind('accordionactivate', function(event, ui) {
          var index = $(this).find("h3").index ( ui.newHeader[0] );
          if (index === 1) {
            plot3.replot();
          }
        });
    
    });
</script> 

<!-- End example scripts -->

<!-- Don't touch this! -->


    <script class="include" type="text/javascript" src="dist/jquery.jqplot.min.js"></script>
    <script type="text/javascript" src="dist/examples/syntaxhighlighter/scripts/shCore.min.js"></script>
    <script type="text/javascript" src="dist/examples/syntaxhighlighter/scripts/shBrushJScript.min.js"></script>
    <script type="text/javascript" src="dist/examples/syntaxhighlighter/scripts/shBrushXml.min.js"></script>
<!-- End Don't touch this! -->

<!-- Additional plugins go here -->
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.cursor.min.js"></script>
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.pieRenderer.min.js"></script>
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.ohlcRenderer.min.js"></script>
  <script class="include" type="text/javascript" src="dist/plugins/jqplot.categoryAxisRenderer.min.js"></script>
  <link class="include" type="text/css" href="css/jquery-ui_tab.css" rel="Stylesheet" /> 
  <script class="include" type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.10.0/jquery-ui.min.js"></script>

<!-- End additional plugins -->

  
    <script type="text/javascript" src="dist/capture.min.js"></script>




</div>
</div>
<%else%>

<%end if%>
<!--#include file="bottom.jsp"-->


    </body>
</html>
