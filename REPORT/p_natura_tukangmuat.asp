<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtFcba = Request("txtFcba")
	v_txtMonth = Request("txtMonth")
	v_txtYear = Request("txtYear")
	v_txtsectionname = Request("txtsectionname")
	v_txtempcode = Request("txtempcode")

	txttdate = txtMonth & "" & txtYear
	If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
	If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
       if  v_txtfcba <> "" then
	       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
		end if
	    ' selectCons = selectCons & ",fcba"
	    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   fcbaname = rsCat("fcName")
		   CompanyName=rsCat("FCCOMPANYNAME")
		   MgrName = rsCat("FCMANAGER")
		   KTUName = rsCat("FCKTU")
		   CropName=rsCat("FCCROP")
		   FcbaType=rsCat("fcType")
	else
	    nSpanTotal = nSpanTotal - 1 
	end if
	 
	
	strGrnInv		=	" select hk_beras_pemuat.fcba,hk_beras_pemuat.employeecode, " & _
						" employee_api.get_employeename( hk_beras_pemuat.fcba,hk_beras_pemuat.employeecode) employeename," & _
						" employee_api.get_sectionname( hk_beras_pemuat.fcba,hk_beras_pemuat.employeecode) division," & _
						" hk_beras_pemuat.HK_Lebih_Basis_1,hk_beras_pemuat.HK_Lebih_Basis_2,hk_beras_pemuat.HK_Lebih_Basis_3,hk_beras_pemuat.HK_Lebih_Basis_4,hk_beras_pemuat.HK_Lebih_Basis_5,hk_beras_pemuat.HK_Lebih_Basis_6,hk_beras_pemuat.HK_Lebih_Basis_7,hk_beras_pemuat.HK_Lebih_Basis_8,hk_beras_pemuat.HK_Lebih_Basis_9,hk_beras_pemuat.HK_Lebih_Basis_10," & _
						" hk_beras_pemuat.HK_Lebih_Basis_11,hk_beras_pemuat.HK_Lebih_Basis_12,hk_beras_pemuat.HK_Lebih_Basis_13,hk_beras_pemuat.HK_Lebih_Basis_14,hk_beras_pemuat.HK_Lebih_Basis_15,hk_beras_pemuat.HK_Lebih_Basis_16,hk_beras_pemuat.HK_Lebih_Basis_17,hk_beras_pemuat.HK_Lebih_Basis_18,hk_beras_pemuat.HK_Lebih_Basis_19,hk_beras_pemuat.HK_Lebih_Basis_20," & _
						" hk_beras_pemuat.HK_Lebih_Basis_21,hk_beras_pemuat.HK_Lebih_Basis_22,hk_beras_pemuat.HK_Lebih_Basis_23,hk_beras_pemuat.HK_Lebih_Basis_24,hk_beras_pemuat.HK_Lebih_Basis_25,hk_beras_pemuat.HK_Lebih_Basis_26,hk_beras_pemuat.HK_Lebih_Basis_27,hk_beras_pemuat.HK_Lebih_Basis_28,hk_beras_pemuat.HK_Lebih_Basis_29,hk_beras_pemuat.HK_Lebih_Basis_30,hk_beras_pemuat.HK_Lebih_Basis_31," & _
						" hk_beras_pemuat.total_HKE," & _
						" hk_beras_pemuat.total_HK_Lebih_Basis," & _
						" case" & _
						"    when hk_beras_pemuat.total_HK_Lebih_Basis < 11 then 0" & _
						"    when hk_beras_pemuat.total_HK_Lebih_Basis < 22 then (hk_beras_pemuat.total_HK_Lebih_Basis*pm.rate_6)" & _
						"    else ((hk_beras_pemuat.total_HK_Lebih_Basis*pm.rate_6)+pm.rate_7)" & _
						"    end as kg_beras" & _
						" from (" & _
						"     select hk_beras.fcba,hk_beras.employeecode," & _
						"     hk_beras.HK_Lebih_Basis_1,hk_beras.HK_Lebih_Basis_2,hk_beras.HK_Lebih_Basis_3,hk_beras.HK_Lebih_Basis_4,hk_beras.HK_Lebih_Basis_5,hk_beras.HK_Lebih_Basis_6,hk_beras.HK_Lebih_Basis_7,hk_beras.HK_Lebih_Basis_8,hk_beras.HK_Lebih_Basis_9,hk_beras.HK_Lebih_Basis_10," & _
						"     hk_beras.HK_Lebih_Basis_11,hk_beras.HK_Lebih_Basis_12,hk_beras.HK_Lebih_Basis_13,hk_beras.HK_Lebih_Basis_14,hk_beras.HK_Lebih_Basis_15,hk_beras.HK_Lebih_Basis_16,hk_beras.HK_Lebih_Basis_17,hk_beras.HK_Lebih_Basis_18,hk_beras.HK_Lebih_Basis_19,hk_beras.HK_Lebih_Basis_20," & _
						"     hk_beras.HK_Lebih_Basis_21,hk_beras.HK_Lebih_Basis_22,hk_beras.HK_Lebih_Basis_23,hk_beras.HK_Lebih_Basis_24,hk_beras.HK_Lebih_Basis_25,hk_beras.HK_Lebih_Basis_26,hk_beras.HK_Lebih_Basis_27,hk_beras.HK_Lebih_Basis_28,hk_beras.HK_Lebih_Basis_29,hk_beras.HK_Lebih_Basis_30,hk_beras.HK_Lebih_Basis_31," & _
						"     (" & _
						"     hk_beras.HKE_1+hk_beras.HKE_2+hk_beras.HKE_3+hk_beras.HKE_4+hk_beras.HKE_5+hk_beras.HKE_6+hk_beras.HKE_7+hk_beras.HKE_8+hk_beras.HKE_9+hk_beras.HKE_10+" & _
						"     hk_beras.HKE_11+hk_beras.HKE_12+hk_beras.HKE_13+hk_beras.HKE_14+hk_beras.HKE_15+hk_beras.HKE_16+hk_beras.HKE_17+hk_beras.HKE_18+hk_beras.HKE_19+hk_beras.HKE_20+" & _
						"     hk_beras.HKE_21+hk_beras.HKE_22+hk_beras.HKE_23+hk_beras.HKE_24+hk_beras.HKE_25+hk_beras.HKE_26+hk_beras.HKE_27+hk_beras.HKE_28+hk_beras.HKE_29+hk_beras.HKE_30+hk_beras.HKE_31" & _
						"     ) as total_HKE," & _
						"     (" & _
						"     hk_beras.HK_Lebih_Basis_1+hk_beras.HK_Lebih_Basis_2+hk_beras.HK_Lebih_Basis_3+hk_beras.HK_Lebih_Basis_4+hk_beras.HK_Lebih_Basis_5+hk_beras.HK_Lebih_Basis_6+hk_beras.HK_Lebih_Basis_7+hk_beras.HK_Lebih_Basis_8+hk_beras.HK_Lebih_Basis_9+hk_beras.HK_Lebih_Basis_10+" & _
						"     hk_beras.HK_Lebih_Basis_11+hk_beras.HK_Lebih_Basis_12+hk_beras.HK_Lebih_Basis_13+hk_beras.HK_Lebih_Basis_14+hk_beras.HK_Lebih_Basis_15+hk_beras.HK_Lebih_Basis_16+hk_beras.HK_Lebih_Basis_17+hk_beras.HK_Lebih_Basis_18+hk_beras.HK_Lebih_Basis_19+hk_beras.HK_Lebih_Basis_20+" & _
						"     hk_beras.HK_Lebih_Basis_21+hk_beras.HK_Lebih_Basis_22+hk_beras.HK_Lebih_Basis_23+hk_beras.HK_Lebih_Basis_24+hk_beras.HK_Lebih_Basis_25+hk_beras.HK_Lebih_Basis_26+hk_beras.HK_Lebih_Basis_27+hk_beras.HK_Lebih_Basis_28+hk_beras.HK_Lebih_Basis_29+hk_beras.HK_Lebih_Basis_30+hk_beras.HK_Lebih_Basis_31" & _
						"     ) as total_HK_Lebih_Basis" & _
						"     from (" & _
						"         select" & _
						"         fcba,employeecode," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),1,HKE,0)) as HKE_1," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),2,HKE,0)) as HKE_2," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),3,HKE,0)) as HKE_3," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),4,HKE,0)) as HKE_4," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),5,HKE,0)) as HKE_5," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),6,HKE,0)) as HKE_6," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),7,HKE,0)) as HKE_7," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),8,HKE,0)) as HKE_8," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),9,HKE,0)) as HKE_9," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),10,HKE,0)) as HKE_10," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),11,HKE,0)) as HKE_11," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),12,HKE,0)) as HKE_12," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),13,HKE,0)) as HKE_13," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),14,HKE,0)) as HKE_14," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),15,HKE,0)) as HKE_15," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),16,HKE,0)) as HKE_16," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),17,HKE,0)) as HKE_17," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),18,HKE,0)) as HKE_18," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),19,HKE,0)) as HKE_19," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),20,HKE,0)) as HKE_20," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),21,HKE,0)) as HKE_21," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),22,HKE,0)) as HKE_22," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),23,HKE,0)) as HKE_23," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),24,HKE,0)) as HKE_24," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),25,HKE,0)) as HKE_25," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),26,HKE,0)) as HKE_26," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),27,HKE,0)) as HKE_27," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),28,HKE,0)) as HKE_28," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),29,HKE,0)) as HKE_29," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),30,HKE,0)) as HKE_30," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),31,HKE,0)) as HKE_31," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),1,HK_Lebih_Basis,0)) as HK_Lebih_Basis_1," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),2,HK_Lebih_Basis,0)) as HK_Lebih_Basis_2," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),3,HK_Lebih_Basis,0)) as HK_Lebih_Basis_3," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),4,HK_Lebih_Basis,0)) as HK_Lebih_Basis_4," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),5,HK_Lebih_Basis,0)) as HK_Lebih_Basis_5," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),6,HK_Lebih_Basis,0)) as HK_Lebih_Basis_6," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),7,HK_Lebih_Basis,0)) as HK_Lebih_Basis_7," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),8,HK_Lebih_Basis,0)) as HK_Lebih_Basis_8," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),9,HK_Lebih_Basis,0)) as HK_Lebih_Basis_9," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),10,HK_Lebih_Basis,0)) as HK_Lebih_Basis_10," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),11,HK_Lebih_Basis,0)) as HK_Lebih_Basis_11," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),12,HK_Lebih_Basis,0)) as HK_Lebih_Basis_12," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),13,HK_Lebih_Basis,0)) as HK_Lebih_Basis_13," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),14,HK_Lebih_Basis,0)) as HK_Lebih_Basis_14," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),15,HK_Lebih_Basis,0)) as HK_Lebih_Basis_15," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),16,HK_Lebih_Basis,0)) as HK_Lebih_Basis_16," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),17,HK_Lebih_Basis,0)) as HK_Lebih_Basis_17," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),18,HK_Lebih_Basis,0)) as HK_Lebih_Basis_18," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),19,HK_Lebih_Basis,0)) as HK_Lebih_Basis_19," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),20,HK_Lebih_Basis,0)) as HK_Lebih_Basis_20," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),21,HK_Lebih_Basis,0)) as HK_Lebih_Basis_21," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),22,HK_Lebih_Basis,0)) as HK_Lebih_Basis_22," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),23,HK_Lebih_Basis,0)) as HK_Lebih_Basis_23," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),24,HK_Lebih_Basis,0)) as HK_Lebih_Basis_24," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),25,HK_Lebih_Basis,0)) as HK_Lebih_Basis_25," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),26,HK_Lebih_Basis,0)) as HK_Lebih_Basis_26," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),27,HK_Lebih_Basis,0)) as HK_Lebih_Basis_27," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),28,HK_Lebih_Basis,0)) as HK_Lebih_Basis_28," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),29,HK_Lebih_Basis,0)) as HK_Lebih_Basis_29," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),30,HK_Lebih_Basis,0)) as HK_Lebih_Basis_30," & _
						"         sum(decode(to_number(to_char(fddate,'dd')),31,HK_Lebih_Basis,0)) as HK_Lebih_Basis_31 " & _
						"         from (" & _
						"             select" & _
						"             fcba," & _
						"             fddate," & _
						"             employeecode," & _
						"             HK_Lebih_Basis," & _
						"             count(HK_Lebih_Basis) as HKE" & _
						"             from (" & _
						"                 select" & _
						"                 fcba," & _
						"                 fddate," & _
						"                 employeecode," & _
						"                 Jumlah_SPB," & _
						"                 case" & _
						"                     when Jumlah_SPB > 1 then 1" & _
						"                     else 0" & _
						"                     end as HK_Lebih_Basis" & _
						"                 from (" & _
						"                     select" & _
						"                     fcba," & _
						"                     fddate," & _
						"                     employeecode," & _
						"                     count(spbno) as Jumlah_SPB" & _
						"                     from (" & _
						"                         select distinct" & _
						"                         fcba," & _
						"                         fddate," & _
						"                         employeecode," & _
						"                         spbno" & _
						"                         from iplasprod.attendance_gad" & _
						"                         where fcba = '"& v_txtfcba &"'" & _
						"                         and to_number(to_char(fddate,'mm')) = "& v_txtMonth &"" & _
						"                         and to_number(to_char(fddate,'yyyy')) = "& v_txtYear &"" & _
						"                         and jobcode in (" & _
						"                             select distinct " & _
						"                             iplasprod.parameterdetail.param_1 " & _
						"                             from iplasprod.parameterdetail " & _
						"                             where iplasprod.parameterdetail.parcode = " & _
						"                             'JOBCODE_PENGANGKUTAN' and fcba = '"& v_txtfcba &"'" & _
						"                         )" & _
						"                     )" & _
						"                     group by" & _
						"                     fcba," & _
						"                     fddate," & _
						"                     employeecode" & _
						"                 )" & _
						"             )" & _
						"             group by" & _
						"             fcba," & _
						"             fddate," & _
						"             employeecode," & _
						"             HK_Lebih_Basis" & _
						"         )" & _
						"         group by" & _
						"         fcba,employeecode" & _
						"     ) hk_beras" & _
						" ) hk_beras_pemuat" & _
						" inner join (" & _
						"     select fcba, rate_6,rate_7 from iplasprod.parameter" & _
						"     where parcode = 'E0006'" & _
						"     and parname = 'MARITAL STATUS'" & _
						"     and fcba = '"& v_txtfcba &"'" & _ 
						"	  and fccode = 'TK0'" & _
						" ) pm" & _
						" on (hk_beras_pemuat.fcba = pm.fcba)" 

	 
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	end if 
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND employee_api.get_sectionname(hk_beras_pemuat.fcba,hk_beras_pemuat.employeecode) ='"& v_txtsectionname &"'"
	end if 


	' if  cdbl(to_nol(compare1)) = 1 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != vehicleactivity.vh_employeecode"
	' end if
	' if  cdbl(to_nol(compare1)) = 2 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != harvestingspb.harvesting_employeecode" 
	' end if
	' if  cdbl(to_nol(compare1)) = 3 then
	'     strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	' end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY employeecode"

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Natura Tukang Muat</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
		
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
            	<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
					<tr> 
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">KEBUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">TAHUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">BULAN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DIVISION</font>
						</td> 
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">EMPLOYEE CODE</font>
						</td>
					</tr>
					<tr>	 
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<option></option>
							<%
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba  = rsFcba("fcba") then
									%>			
										<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
									<% Else %>	
										<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
									<% End If
									
									rsFcba.movenext
								loop
							End if
							rsFcba.Close
							set rsFcba = Nothing
							%>
							</select>	
						</td> 
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
						<td width="4">
							<select name="txtYear">
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td> 
						<td>
							<select name="txtsectionname" >
								<option value="">-- ALL --</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	"select fccode, fcname from sectionname where fccode in ('AFD-01','AFD-02','AFD-03','AFD-04','AFD-05','AFD-06','AFD-07','AFD-08','AFD-09','AFD-10','AFD-11','AFD-12','AFD-13') order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtsectionname then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from employee where fcba = '"& v_txtfcba &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td>
						<!--
						<td>
							<font face="verdana" size="1">VEHICLE CODE</font>
						</td>
						<td width="2">
							<input type="text" name="txtvehiclecode" value="<%=txtvehiclecode%>">
						</td> -->
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="37" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>LAPORAN NATURA TUKANG MUAT</h2>
				<h4>PERIODE : Year <%=v_txtyear%>, Month <%=v_txtmonth%></h4>
                <br/>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th> 
			<th style="background-color:#00A300;color:#fff;">Division</th>
			<th style="background-color:#00A300;color:#fff;">1</th> 
			<th style="background-color:#00A300;color:#fff;">2</th> 
			<th style="background-color:#00A300;color:#fff;">3</th> 
			<th style="background-color:#00A300;color:#fff;">4</th> 
			<th style="background-color:#00A300;color:#fff;">5</th> 
			<th style="background-color:#00A300;color:#fff;">6</th> 
			<th style="background-color:#00A300;color:#fff;">7</th> 
			<th style="background-color:#00A300;color:#fff;">8</th> 
			<th style="background-color:#00A300;color:#fff;">9</th> 
			<th style="background-color:#00A300;color:#fff;">10</th> 
			<th style="background-color:#00A300;color:#fff;">11</th> 
			<th style="background-color:#00A300;color:#fff;">12</th> 
			<th style="background-color:#00A300;color:#fff;">13</th> 
			<th style="background-color:#00A300;color:#fff;">14</th> 
			<th style="background-color:#00A300;color:#fff;">15</th> 
			<th style="background-color:#00A300;color:#fff;">16</th> 
			<th style="background-color:#00A300;color:#fff;">17</th> 
			<th style="background-color:#00A300;color:#fff;">18</th> 
			<th style="background-color:#00A300;color:#fff;">19</th> 
			<th style="background-color:#00A300;color:#fff;">20</th> 
			<th style="background-color:#00A300;color:#fff;">21</th> 
			<th style="background-color:#00A300;color:#fff;">22</th> 
			<th style="background-color:#00A300;color:#fff;">23</th> 
			<th style="background-color:#00A300;color:#fff;">24</th> 
			<th style="background-color:#00A300;color:#fff;">25</th> 
			<th style="background-color:#00A300;color:#fff;">26</th> 
			<th style="background-color:#00A300;color:#fff;">27</th> 
			<th style="background-color:#00A300;color:#fff;">28</th> 
			<th style="background-color:#00A300;color:#fff;">29</th> 
			<th style="background-color:#00A300;color:#fff;">30</th> 
			<th style="background-color:#00A300;color:#fff;">31</th> 
			<th style="background-color:#00A300;color:#fff;">Total HKE</th>
			<th style="background-color:#00A300;color:#fff;">Total HK Lebih Basis</th> 
			<th style="background-color:#00A300;color:#fff;">KG Beras</th> 
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="37" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_beras_1 = 0
		v_beras_2 = 0
		v_beras_3 = 0
		v_beras_4 = 0
		v_beras_5 = 0
		v_beras_6 = 0
		v_beras_7 = 0
		v_beras_8 = 0
		v_beras_9 = 0
		v_beras_10 = 0
		v_beras_11 = 0
		v_beras_12 = 0
		v_beras_13 = 0
		v_beras_14 = 0
		v_beras_15 = 0
		v_beras_16 = 0
		v_beras_17 = 0
		v_beras_18 = 0
		v_beras_19 = 0
		v_beras_20 = 0
		v_beras_21 = 0
		v_beras_22 = 0
		v_beras_23 = 0
		v_beras_24 = 0
		v_beras_25 = 0
		v_beras_26 = 0
		v_beras_27 = 0
		v_beras_28 = 0
		v_beras_29 = 0
		v_beras_30 = 0
		v_beras_31 = 0
		v_total = 0
		v_total_hk_lebih_basis = 0
		v_kg_beras = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("employeecode")%></td>
			<td align="left"><%=rsGrnInv("employeename")%></td> 
			<td align="center"><%=rsGrnInv("division")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_1")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_2")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_3")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_4")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_5")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_6")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_7")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_8")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_9")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_10")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_11")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_12")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_13")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_14")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_15")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_16")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_17")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_18")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_19")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_20")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_21")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_22")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_23")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_24")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_25")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_26")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_27")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_28")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_29")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_30")%></td>
			<td align="center"><%=rsGrnInv("HK_Lebih_Basis_31")%></td>
			<td align="center"><%=rsGrnInv("total_hke")%></td>
			<td align="center"><%=rsGrnInv("total_HK_Lebih_Basis")%></td>
			<td align="center"><%=rsGrnInv("kg_beras")%></td>
		</tr>
<%			
			v_HK_Lebih_Basis_1 = v_HK_Lebih_Basis_1 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_1"))) 
			v_HK_Lebih_Basis_2 = v_HK_Lebih_Basis_2 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_2"))) 
			v_HK_Lebih_Basis_3 = v_HK_Lebih_Basis_3 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_3"))) 
			v_HK_Lebih_Basis_4 = v_HK_Lebih_Basis_4 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_4"))) 
			v_HK_Lebih_Basis_5 = v_HK_Lebih_Basis_5 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_5"))) 
			v_HK_Lebih_Basis_6 = v_HK_Lebih_Basis_6 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_6"))) 
			v_HK_Lebih_Basis_7 = v_HK_Lebih_Basis_7 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_7"))) 
			v_HK_Lebih_Basis_8 = v_HK_Lebih_Basis_8 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_8"))) 
			v_HK_Lebih_Basis_9 = v_HK_Lebih_Basis_9 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_9"))) 
			v_HK_Lebih_Basis_10 = v_HK_Lebih_Basis_10 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_10"))) 
			v_HK_Lebih_Basis_11 = v_HK_Lebih_Basis_11 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_11"))) 
			v_HK_Lebih_Basis_12 = v_HK_Lebih_Basis_12 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_12"))) 
			v_HK_Lebih_Basis_13 = v_HK_Lebih_Basis_13 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_13"))) 
			v_HK_Lebih_Basis_14 = v_HK_Lebih_Basis_14 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_14"))) 
			v_HK_Lebih_Basis_15 = v_HK_Lebih_Basis_15 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_15"))) 
			v_HK_Lebih_Basis_16 = v_HK_Lebih_Basis_16 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_16"))) 
			v_HK_Lebih_Basis_17 = v_HK_Lebih_Basis_17 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_17"))) 
			v_HK_Lebih_Basis_18 = v_HK_Lebih_Basis_18 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_18"))) 
			v_HK_Lebih_Basis_19 = v_HK_Lebih_Basis_19 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_19"))) 
			v_HK_Lebih_Basis_20 = v_HK_Lebih_Basis_20 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_20"))) 
			v_HK_Lebih_Basis_21 = v_HK_Lebih_Basis_21 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_21"))) 
			v_HK_Lebih_Basis_22 = v_HK_Lebih_Basis_22 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_22"))) 
			v_HK_Lebih_Basis_23 = v_HK_Lebih_Basis_23 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_23"))) 
			v_HK_Lebih_Basis_24 = v_HK_Lebih_Basis_24 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_24"))) 
			v_HK_Lebih_Basis_25 = v_HK_Lebih_Basis_25 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_25"))) 
			v_HK_Lebih_Basis_26 = v_HK_Lebih_Basis_26 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_26"))) 
			v_HK_Lebih_Basis_27 = v_HK_Lebih_Basis_27 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_27"))) 
			v_HK_Lebih_Basis_28 = v_HK_Lebih_Basis_28 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_28"))) 
			v_HK_Lebih_Basis_29 = v_HK_Lebih_Basis_29 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_29"))) 
			v_HK_Lebih_Basis_30 = v_HK_Lebih_Basis_30 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_30"))) 
			v_HK_Lebih_Basis_31 = v_HK_Lebih_Basis_31 + cdbl(to_nol(rsGrnInv("HK_Lebih_Basis_31"))) 
			v_total = v_total + cdbl(to_nol(rsGrnInv("total_hke"))) 
			v_total_hk_lebih_basis = v_total_hk_lebih_basis + cdbl(to_nol(rsGrnInv("total_HK_Lebih_Basis"))) 
			v_kg_beras = v_kg_beras + cdbl(to_nol(rsGrnInv("kg_beras")))

			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="3"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_1)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_2)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_3)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_4)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_5)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_6)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_7)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_8)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_9)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_10)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_11)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_12)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_13)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_14)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_15)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_16)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_17)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_18)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_19)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_20)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_21)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_22)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_23)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_24)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_25)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_26)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_27)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_28)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_29)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_30)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_HK_Lebih_Basis_31)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_hk_lebih_basis)),2)%></font></td>
			<td align="center"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_beras)),2)%></font></td>
		</tr>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"45%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");
	
</script>