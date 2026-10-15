<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
ScriptName			= Request.ServerVariables("SCRIPT_NAME")
v_txtCompany		= Request("company")
v_txtFcba			= Request("txtFcba")
v_txtYear			= Request("txtYear")
v_txtMonthFrom		= Request("txtMonthFrom")
v_txtMonthTo		= Request("txtMonthTo")
v_txtJobcode		= Request("txtJobcode")
whereCons			=""
whereConsti			=""
%>
<!DOCTYPE html>
<!--[if IE 8]> <html lang="en" class="ie8"> <![endif]-->
<!--[if !IE]><!-->
<html lang="en">
<!--<![endif]-->
<head>
	<meta charset="utf-8" />
	<title>ePlaS | Integrated Plantation System</title>
	<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" name="viewport" />
	<meta content="" name="description" />
	<meta content="" name="author" />
	
	<!-- ================== BEGIN BASE CSS STYLE ================== -->
	<link href="http://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/jquery-ui/jquery-ui.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/bootstrap/4.1.0/css/bootstrap.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/font-awesome/5.0/css/fontawesome-all.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/animate/animate.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/css/default/style.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/css/default/style-responsive.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/css/default/theme/default.css" rel="stylesheet" id="theme" />
	<!-- ================== END BASE CSS STYLE ================== -->
	
	<!-- ================== BEGIN PAGE LEVEL STYLE ================== -->
	<link href="../PRINT/assets/plugins/jquery-jvectormap/jquery-jvectormap.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/bootstrap-datepicker/css/bootstrap-datepicker.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/bootstrap-timepicker/css/bootstrap-timepicker.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/bootstrap-select/bootstrap-select.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/bootstrap-combobox/css/bootstrap-combobox.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/select2/dist/css/select2.min.css" rel="stylesheet" />
    <link href="../PRINT/assets/plugins/gritter/css/jquery.gritter.css" rel="stylesheet" />
	<!-- ================== END PAGE LEVEL STYLE ================== -->
	
	<!-- ================== BEGIN PAGE LEVEL STYLE ================== -->
    <link href="../PRINT/assets/plugins/DataTables/media/css/dataTables.bootstrap.min.css" rel="stylesheet" />
	<link href="../PRINT/assets/plugins/DataTables/extensions/Buttons/css/buttons.bootstrap.min.css" rel="stylesheet" />
    <link href="../PRINT/assets/plugins/DataTables/extensions/AutoFill/css/autoFill.bootstrap.min.css" rel="stylesheet" />
    <link href="../PRINT/assets/plugins/DataTables/extensions/Responsive/css/responsive.bootstrap.min.css" rel="stylesheet" />
	 <link href="../PRINT/assets/plugins/nvd3/build/nv.d3.css" rel="stylesheet" />
	<!-- ================== END PAGE LEVEL STYLE ================== -->
	
	<!-- ================== BEGIN BASE JS ================== -->
	<script src="../PRINT/assets/plugins/pace/pace.min.js"></script>
	<!-- ================== END BASE JS ================== -->
	
	<!-- ================== BEGIN BASE JS ================== -->
	<script src="../PRINT/assets/plugins/jquery/jquery-3.2.1.min.js"></script>
	<script src="../PRINT/assets/plugins/jquery-ui/jquery-ui.min.js"></script>
	<script src="../PRINT/assets/plugins/bootstrap/4.1.0/js/bootstrap.bundle.min.js"></script>
	<!--[if lt IE 9]>
		<script src="../PRINT/assets/crossbrowserjs/html5shiv.js"></script>
		<script src="../PRINT/assets/crossbrowserjs/respond.min.js"></script>
		<script src="../PRINT/assets/crossbrowserjs/excanvas.min.js"></script>
	<![endif]-->
	<script src="../PRINT/assets/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<script src="../PRINT/assets/plugins/js-cookie/js.cookie.js"></script>
	<script src="../PRINT/assets/js/theme/default.min.js"></script>
	<script src="../PRINT/assets/js/apps.min.js"></script>
	<!-- ================== END BASE JS ================== -->
	
	<!-- ================== BEGIN PAGE LEVEL JS ================== 
	<script src="../PRINT/assets/plugins/gritter/js/jquery.gritter.js"></script>
	<script src="../PRINT/assets/plugins/bootstrap-sweetalert/sweetalert.min.js"></script>
	<script src="../PRINT/assets/plugins/flot/jquery.flot.min.js"></script>
	<script src="../PRINT/assets/plugins/flot/jquery.flot.time.min.js"></script>
	<script src="../PRINT/assets/plugins/flot/jquery.flot.resize.min.js"></script>
	<script src="../PRINT/assets/plugins/flot/jquery.flot.pie.min.js"></script>
	<script src="../PRINT/assets/plugins/sparkline/jquery.sparkline.js"></script>
	<script src="../PRINT/assets/plugins/jquery-jvectormap/jquery-jvectormap.min.js"></script>
	<script src="../PRINT/assets/plugins/jquery-jvectormap/jquery-jvectormap-world-mill-en.js"></script>
	<script src="../PRINT/assets/plugins/bootstrap-datepicker/js/bootstrap-datepicker.js"></script>
	<script src="../PRINT/assets/plugins/bootstrap-timepicker/js/bootstrap-timepicker.min.js"></script>
	<script src="../PRINT/assets/plugins/bootstrap-select/bootstrap-select.min.js"></script>
	<script src="../PRINT/assets/plugins/bootstrap-combobox/js/bootstrap-combobox.js"></script>
	<script src="../PRINT/assets/plugins/select2/dist/js/select2.min.js"></script>
	<script src="../PRINT/assets/js/demo/dashboard.min.js"></script>
	<!-- ================== END PAGE LEVEL JS ================== -->	
	
	<script src="../PRINT/assets/plugins/DataTables/media/js/jquery.dataTables.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/media/js/dataTables.bootstrap.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/dataTables.buttons.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/buttons.bootstrap.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/buttons.flash.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/jszip.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/pdfmake.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/vfs_fonts.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/buttons.html5.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Buttons/js/buttons.print.min.js"></script>
	<script src="../PRINT/assets/plugins/DataTables/extensions/Responsive/js/dataTables.responsive.min.js"></script>
</head>
<body>
	<!-- begin widget-chart -->
	<div class="widget widget-rounded m-b-5">
		<!-- begin widget-header -->
		<div class="widget-header">
			<h4 class="widget-header-title">
			<span class="label label-inverse pull-right">I N T I</span>
			<strong>General Ledger PT. NIKP</strong></h4>
		</div>
		<!-- end widget-header -->
		
		<%
			    if v_txtFcba <> "ALL" then
				   whereCons 	= " and ca.fcba = '" & v_txtFcba & "'"
				   whereConsti 	= " and ti.fcba = '" & v_txtFcba & "'"
				end if
				
				if v_txtYear <> "" then
						v_txtYear = Request("txtYear")
					else
						v_txtYear = Year(Date)
				end if
				
				if v_txtMonthFrom <> "" then
						v_txtMonthFrom = Request("txtMonthFrom")
					else
						v_txtMonthFrom = Month(Date)
				end if
				
				if v_txtMonthTo <> "" then
						v_txtMonthTo = Request("txtMonthTo")
					else
						v_txtMonthTo = Month(Date)
				end if
				
				
		%>
		<!-- begin vertical-box -->
		<div class="horizontal-box with-grid with-border-top">
			<div class="alert alert-secondary fade show">
				<div class="note note-secondary m-b-0">
					<form name="frmBTL">
									<div class="btn-group">
										<select class="form-control company" data-size="10" id='company' data-plugin='select2' name='company' data-live-search="true" data-style="btn-primary">
											<option value="ALL">Select Company </option>
											<%
											strCompany	=	" SELECT COMPANYCODE FROM COMPANY WHERE COMPANYCODE = 'PT.NIKP'"
											
											Set rsCompany = Server.CreateObject("ADODB.Recordset")
											rsCompany.Open strCompany, connstr, adOpenStatic
											
											If rsCompany.Bof Or rsCompany.eof Then
											%>
											<option value=""></option>
											<%
											End If
											
											if not rsCompany.eof then
											Do While Not rsCompany.eof
											%>
											<option value="<%=rsCompany("COMPANYCODE")%>" <%if v_txtCompany = rsCompany("COMPANYCODE") then%> Selected <%End If%>><%=rsCompany("COMPANYCODE")%></option>
											<%
												rsCompany.movenext
											loop
											end if
											rsCompany.Close
											set rsCompany = Nothing
											%>
										</select>
									</div>
									<div class="btn-group">
										<select class="form-control txtFcba" data-size="10" id='txtFcba' data-plugin='select2' name='txtFcba' data-live-search="true" data-style="btn-primary">
											<option value="ALL">--ALL--</option>
											<%
											Dim rsFcba, strFcba
											strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
														" WHERE ISACTIVE = 'TRUE' and USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
														" ORDER BY ORDERDATA"
											
											Set rsFcba = Server.CreateObject("ADODB.Recordset")
											rsFcba.Open strFcba, connstr, adOpenStatic
											
											If rsFcba.Bof Or rsFcba.eof Then
											%>
											<option value=""></option>
											<%
											End If
											
											if not rsFcba.eof then
											Do While Not rsFcba.eof
											%>
											<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
											<%
												rsFcba.movenext
											loop
											end if
											rsFcba.Close
											set rsFcba = Nothing
											%>
										</select>
									</div>
									<div class="btn-group">
										<% 
											mon = Month(Session("v_Costbook"))
											Yr = Year(Session("v_Costbook"))
											MinYr = 2015
											MaxYr = Yr  + 2											
										%>
										<select class="form-control txtYear" data-size="10" id='txtYear' data-plugin='select2' name='txtYear' data-live-search="true" data-style="btn-primary">
										<%
											For i = MinYr to MaxYr
										%>
											<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
										<%
											Next
										%>
										</select>
									</div>
									From
									<div class="btn-group">										
										<select class="form-control txtMonthFrom" data-size="10" id='txtMonthFrom' data-plugin='select2' name='txtMonthFrom' data-live-search="true" data-style="btn-primary">
											<%
												For i = 1 to 15
											%>
												<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonthFrom) then%> Selected <%End If%>><%=BulanAP(i)%></option>
											<%
												Next
											%>
										</select>
									</div>
									To
									<div class="btn-group">										
										<select class="form-control txtMonthTo" data-size="10" id='txtMonthTo' data-plugin='select2' name='txtMonthTo' data-live-search="true" data-style="btn-primary">
											<%
												For i = 1 to 15
											%>
												<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonthTo) then%> Selected <%End If%>><%=BulanAP(i)%></option>
											<%
												Next
											%>
										</select>
									</div>
									Job Code
									<div class="btn-group">
										<select class="form-control txtJobCode" data-size="10" id='txtJobCode' data-plugin='select2' name='txtJobCode' data-live-search="true" data-style="btn-primary">
											<option value="ALL">--ALL--</option>
											<%
												Dim rsJobcode, strJobcode
												strJobcode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '"& v_txtFcba &"'"
												
												Set rsJobcode = Server.CreateObject("ADODB.Recordset")
												rsJobcode.Open strJobcode, connstr, adOpenStatic
												
												If rsJobcode.Bof Or rsJobcode.eof Then
												%>
												<option value=""></option>
												<%
												End If
												
												if not rsJobcode.eof then
												Do While Not rsJobcode.eof
												%>
													<option value="<%=rsJobcode("FCCODE")%>" <%if v_txtJobCode = rsJobcode("FCCODE") then%> Selected <%End If%>><%=rsJobcode("FCCODE")%></option>
												<%
													rsJobcode.movenext
												loop
												end if
												rsJobcode.Close
												set rsJobcode = Nothing
											%>
										</select>
									</div>
									<div class="btn-group">
										<button type="submit" id="btnLoad" class="btn btn-success btn-sm">Load</button>&nbsp;
									</div>
									<div class="btn-group">
										<button type="button" id="btnExit" onclick="exitReport()" class="btn btn-warning btn-sm">Exit</button>&nbsp;
									</div>
					</form>
				</div>
					<div class="tab-content">
						<table id="table" class="table table-hover display nowrap" cellspacing="0" width="100%">
								<thead>
									<tr>
										<th>Job Code</th>
										<th>Job Name</th>
										<th>Business Unit</th>
										<th>Month</th>
										<th>Source Type</th>
										<th>Source Code</th>
										<th>Target Code</th>
										<th>Reference No</th>
										<th>Description</th>
										<th>Debet</th>
										<th>Credit</th>
										<th>Saldo</th>
									</tr>
								</thead>
								<tbody>
									<%
										strAnalisaBukuBesar		=	"  SELECT ca.fcba,																					"&_
																	"         ca.tdate,                                                                                 "&_
																	"         ca.fccode,                                                                                "&_
																	"         ca.fcname,                                                                                "&_
																	"         2 orderby,                                                                                "&_
																	"         ca.sourcetype,                                                                            "&_
																	"         ca.sourcecode,                                                                            "&_
																	"         ca.targettype,                                                                            "&_
																	"         ca.targetcode,                                                                            "&_
																	"         ca.ownership,                                                                             "&_
																	"         ca.koperasi,                                                                              "&_
																	"         ca.referenceno,                                                                           "&_
																	"         ca.remarks,                                                                               "&_
																	"         NVL (SUM (nDb_AmtInti), 0) nDb_AmtInti,                                                   "&_
																	"         NVL (SUM (nCr_AmtInti), 0) nCr_AmtInti,                                                   "&_
																	"         NULL saldo                                                                                "&_
																	"    FROM (  SELECT fcba, tdate, fccode, fcname,                                                    "&_
																	"                   sourcetype, sourcecode, targettype, targetcode,                                 "&_
																	"                   ownership, koperasi, referenceno, remarks,                                      "&_
																	"                   NVL (SUM (nDb_AmtInti), 0) nDb_AmtInti,                                         "&_
																	"                   NVL (SUM (nCr_AmtInti), 0) nCr_AmtInti                                          "&_
																	"              FROM (SELECT ca.fcba, ca.tdate, jb.fccode, jb.fcname,                                "&_
																	"                           ca.sourcetype, ca.sourcecode, ca.targettype, ca.targetcode,             "&_
																	"                           'INTI' ownership, 'INTI' koperasi, ca.referenceno,                      "&_
																	"                           (SELECT remarks                                                         "&_
																	"                              FROM View_Voucher_to_Cost_Activity                                   "&_
																	"                             WHERE     SourceType = ca.sourcetype                                  "&_
																	"                                   AND fcba = ca.fcba                                              "&_
																	"                                   AND ReferenceNo = ca.referenceno                                "&_
																	"                                   AND SourceCode = ca.sourcecode                                  "&_
																	"                                   AND jobcode = ca.jobcode                                        "&_
																	"                                   AND TargetType = ca.targettype                                  "&_
																	"                                   AND TargetCode = ca.targetcode                                  "&_
																	"                                   AND ROWNUM = 1)                                                 "&_
																	"                              remarks,                                                             "&_
																	"                           CASE WHEN ca.amount > 0 THEN ca.amount END nDb_AmtInti,                 "&_
																	"                           CASE WHEN ca.amount < 0 THEN ca.amount END nCr_AmtInti                  "&_
																	"                      FROM cost_activity ca                                                        "&_
																	"                           INNER JOIN job jb                                                       "&_
																	"                              ON (ca.fcba = jb.fcba AND ca.jobcode = jb.fccode)                    "&_
																	"                     WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) = "& v_txtYear &"            "&_
																	"                           AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) = "& v_txtMonthTo &"           "&_
																	"                           AND jb.job_own_tb = 'INTI'                                              "&_
																	"                           " & whereCons &")                                                       "& _
																	"          GROUP BY fcba, tdate, fccode, fcname, sourcetype, sourcecode, targettype, targetcode,    "&_
																	"                    ownership,koperasi, referenceno, remarks                                       "&_
																	"          UNION ALL                                                                                "&_
																	"            SELECT fcba, tdate, fccode, fcname, sourcetype, sourcecode,                            "&_
																	"                   targettype, targetcode, ownership, koperasi,                                    "&_
																	"                   referenceno, remarks,                                                           "&_
																	"                   SUM (nDb_AmtInti) nDb_AmtInti,                                                  "&_
																	"                   SUM (nCr_AmtInti) nDb_AmtInti                                                   "&_
																	"              FROM (SELECT ca.fcba, ca.tdate, jb.fccode, jb.fcname,                                "&_
																	"                           ca.sourcetype, ca.sourcecode, ca.targettype,                            "&_
																	"                           field_api.get_plantingyear (ca.targetcode, ca.fcba) targetcode,         "&_
																	"                           'INTI' ownership, 'INTI' koperasi, ca.referenceno,                      "&_
																	"                           (SELECT remarks                                                         "&_
																	"                              FROM View_Voucher_to_Cost_Activity                                   "&_
																	"                             WHERE     SourceType = ca.sourcetype                                  "&_
																	"                                   AND fcba = ca.fcba                                              "&_
																	"                                   AND ReferenceNo = ca.referenceno                                "&_
																	"                                   AND SourceCode = ca.sourcecode                                  "&_
																	"                                   AND jobcode = ca.jobcode                                        "&_
																	"                                   AND TargetType = ca.targettype                                  "&_
																	"                                   AND TargetCode = ca.targetcode                                  "&_
																	"                                   AND ROWNUM = 1)                                                 "&_
																	"                              remarks,                                                             "&_
																	"                           CASE WHEN ca.amount > 0 THEN ca.amount END nDb_AmtInti,                 "&_
																	"                           CASE WHEN ca.amount < 0 THEN ca.amount END nCr_AmtInti                  "&_
																	"                      FROM cost_activity ca                                                        "&_
																	"                           INNER JOIN job jb                                                       "&_
																	"                              ON (ca.fcba = jb.fcba AND ca.jobcode = jb.fccode)                    "&_
																	"                           INNER JOIN field ff                                                     "&_
																	"                              ON (ca.fcba = ff.fcba AND ca.targetcode = ff.fccode)                 "&_
																	"                     WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) = "& v_txtYear &"            "&_
																	"                           AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) = "& v_txtMonthTo &"           "&_
																	"                           AND ff.ownership = 'INTI'                                               "&_
																	"                           AND ca.targettype = 'FF'                                                "&_
																	"                           " & whereCons &")                                                       "& _
																	"          GROUP BY fcba, tdate, fccode, fcname, sourcetype, sourcecode, targettype, targetcode,    "&_
																	"                   ownership, koperasi, referenceno, remarks                                       "&_
																	"          UNION ALL                                                                                "&_
																	"            SELECT fcba, tdate, fccode, fcname,                                                    "&_
																	"                   sourcetype, sourcecode, targettype, targetcode,                                 "&_
																	"                   ownership, koperasi, referenceno, remarks,                                      "&_
																	"                   SUM (nDb_AmtInti) nDb_AmtInti,                                                  "&_
																	"                   SUM (nCr_AmtInti) nDb_AmtInti                                                   "&_
																	"              FROM (SELECT ca.fcba, ca.tdate, jb.fccode, jb.fcname,                                "&_
																	"                           ca.sourcetype, ca.sourcecode, ca.targettype, ca.targetcode,             "&_
																	"                           'INTI' ownership, 'INTI' koperasi, ca.referenceno,                      "&_
																	"                           (SELECT remarks                                                         "&_
																	"                              FROM View_Voucher_to_Cost_Activity                                   "&_
																	"                             WHERE     SourceType = ca.sourcetype                                  "&_
																	"                                   AND fcba = ca.fcba                                              "&_
																	"                                   AND ReferenceNo = ca.referenceno                                "&_
																	"                                   AND SourceCode = ca.sourcecode                                  "&_
																	"                                   AND jobcode = ca.jobcode                                        "&_
																	"                                   AND TargetType = ca.targettype                                  "&_
																	"                                   AND TargetCode = ca.targetcode                                  "&_
																	"                                   AND ROWNUM = 1)                                                 "&_
																	"                              remarks,                                                             "&_
																	"                           CASE WHEN ca.amount > 0 THEN ca.amount END nDb_AmtInti,                 "&_
																	"                           CASE WHEN ca.amount < 0 THEN ca.amount END nCr_AmtInti                  "&_
																	"                      FROM iplasprod.cost_activity ca                                              "&_
																	"                           INNER JOIN iplasprod.job jb                                             "&_
																	"                              ON (ca.fcba = jb.fcba AND ca.jobcode = jb.fccode)                    "&_
																	"                     WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) = "& v_txtYear &"            "&_
																	"                           AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) = "& v_txtMonthTo &"           "&_
																	"                           AND ca.targettype = 'GC'                                                "&_
																	"                           AND ca.targetcode NOT LIKE 'CO%'                                        "&_
																	"                           " & whereCons &")                                                       "& _
																	"          GROUP BY fcba, tdate, fccode, fcname, sourcetype, sourcecode, targettype, targetcode,    "&_
																	"                   ownership, koperasi, referenceno,remarks) ca                                    "&_
																	"GROUP BY ca.fcba, ca.tdate, ca.fccode, ca.fcname, ca.sourcetype, ca.sourcecode,                    "&_
																	"         ca.targettype, ca.targetcode, ca.ownership, ca.koperasi,                                  "&_
																	"         ca.referenceno,ca.remarks                                                                 "&_
																	"UNION ALL                                                                                          "&_
																	"  SELECT ti.fcba fcba,                                                                             "&_
																	"         TO_DATE (ti.period_month, 'mm') tdate,                                                    "&_
																	"         ti.fccode fccode,                                                                         "&_
																	"         ti.fcname fcname,                                                                         "&_
																	"         1 orderby, NULL sourcetype, NULL sourcecode, NULL targettype, NULL targetcode,            "&_
																	"         NULL ownership, NULL koperasi, NULL referenceno, 'OPENING BALANCE' remarks,               "&_
																	"         NULL nDb_AmtInti,NULL nCr_AmtInti,                                                        "&_
																	"         SUM (ti.opening_balance) saldo                                                            "&_
																	"    FROM trialbalance_inti ti                                                                      "&_
																	"   WHERE     ti.period_month = "& v_txtMonthTo &" - ("& v_txtMonthTo &" - "& v_txtMonthFrom &")    "&_
																	"         "& whereConsti &"                                                                   		"&_
																	"         AND ti.period_year = "& v_txtYear &"                                                      "&_
																	"GROUP BY ti.fcba, TO_DATE (ti.period_month, 'mm'), ti.fccode, ti.fcname                            "&_
																	"UNION ALL                                                                                          "&_
																	"  SELECT ti.fcba fcba,                                                                             "&_
																	"         TO_DATE (ti.period_month, 'mm') tdate,                                                    "&_
																	"         ti.fccode fccode,                                                                         "&_
																	"         ti.fcname fcname,                                                                         "&_
																	"         3 orderby, NULL sourcetype, NULL sourcecode, NULL targettype, NULL targetcode,            "&_
																	"         NULL ownership, NULL koperasi, NULL referenceno, 'ENDING BALANCE' remarks,                "&_
																	"         NULL nDb_AmtInti, NULL nCr_AmtInti,                                                       "&_
																	"         SUM (ti.ending_balance) saldo                                                             "&_
																	"    FROM trialbalance_inti ti                                                                      "&_
																	"   WHERE     ti.period_month = "& v_txtMonthTo &" - ("& v_txtMonthTo &" - "& v_txtMonthFrom &")    "&_
																	"         "& whereConsti &"                                                                   		"&_
																	"         AND ti.period_year = "& v_txtYear &"                                                      "&_
																	"GROUP BY ti.fcba, TO_DATE (ti.period_month, 'mm'), ti.fccode, ti.fcname                            "&_
																	"ORDER BY fcba, fccode, fcname, orderby                                                             "
									
									'Response.Write strAnalisaBukuBesar

									Set rsAnalisaBukuBesar = Server.CreateObject("ADODB.Recordset")
									rsAnalisaBukuBesar.Open strAnalisaBukuBesar, connstr, adOpenStatic
									if Not rsAnalisaBukuBesar.eof then
										rsAnalisaBukuBesar.MoveFirst

										While Not rsAnalisaBukuBesar.EOF
												rowno	=	rowno	+	1
												Currsourcetype = rsAnalisaBukuBesar("fccode")
												Currsourcetypedesc = rsAnalisaBukuBesar("fcname")
												'If Cdbl(rsAnalisaBukuBesar("amount")) = 0 Then
												'	GrandTotalPositive = GrandTotalPositive + 0
												'	GrandTotalNegative = GrandTotalNegative + 0
												'Else
												'	If Cdbl(rsAnalisaBukuBesar("amount")) > 0 Then
												'		GrandTotalPositive = GrandTotalPositive + Cdbl(rsAnalisaBukuBesar("amount"))
												'	ElseIf Cdbl(rsAnalisaBukuBesar("amount")) < 0 Then
												'		GrandTotalNegative = GrandTotalNegative + Cdbl(rsAnalisaBukuBesar("amount"))
												'	End If
												'End If
												
												%>													
												<tr>
												<td align="left"><%=rsAnalisaBukuBesar("fccode")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("fcname")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("fcba")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("tdate")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("sourcetype")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("sourcecode")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("targetcode")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("referenceno")%></td>
												<td align="left"><%=rsAnalisaBukuBesar("remarks")%></td>
												<td align="left"><%=FormatNumber(to_nol(rsAnalisaBukuBesar("nDb_AmtInti")), 2)%></td>
												<td align="left"><%=FormatNumber(to_nol(rsAnalisaBukuBesar("nCr_AmtInti")), 2)%></td>
												<td align="left"><%=FormatNumber(to_nol(rsAnalisaBukuBesar("saldo")), 2)%></td>
												</tr>
												<%
										
												Prevsourcetype = Currsourcetype
												Prevsourcetypedesc = Currsourcetypedesc
												NewSourceType = rsAnalisaBukuBesar("fcname")
												rsAnalisaBukuBesar.MoveNext
										Wend
									end if
									n_SaldoSub = n_SaldoSub
									n_SaldoGrand = n_SaldoGrand + GrandTotalPositive + GrandTotalNegative
									%>

								</tbody>
						</table>
					</div>
			</div>
		</div>
		<!-- end vertical-box -->
	</div>
<!-- end widget-chart -->
	
<!-- #modal-dialog -->
	<div class="modal" id="modal-dialog">
		<div class="modal-dialog" style="max-width:623px">
			<div class="modal-content">
				<div class="modal-body">
					<table id="tabledetail" class="table" width="100%">
						<thead width="100%">
							<tr>
								<th></th>
								<th> <span class="label label-success">Debet</span></th>
								<th> <span class="label label-warning">Credit</span></th>
							</tr>
						</thead>
						<tbody width="100%">
							<tr><td> Job Code </td>
								<td><label id="lblJOBCODE_DB"></label></td>
								<td><label id="lblJOBCODE_CR"></label></td>
							</tr>
							<tr>
								<td> Job Name </td>
								<td><label id="lblFCNAME_DB"></label></td>
								<td><label id="lblFCNAME_CR"></label></td>
							</tr>
							<tr>
								<td> Source Type </td>
								<td><label id="lblSOURCETYPE_DB"></label></td>
								<td><label id="lblSOURCETYPE_CR"></label></td>
							</tr>
							<tr>
								<td> Source Code </td>
								<td><label id="lblSOURCECODE_DB"></label></td>
								<td><label id="lblSOURCECODE_CR"></label></td>
							</tr>
							<tr>
								<td> Target Code </td>
								<td><label id="lblTARGETCODE_DB"></label></td>
								<td><label id="lblTARGETCODE_CR"></label></td>
							</tr>
							<tr>
								<td> Reference no </td>
								<td><label id="lblREFERENCENO_DB"></label></td>
								<td><label id="lblREFERENCENO_CR"></label></td>
							</tr>
							<tr>
								<td> Amount </td>
								<td><label id="lblAMOUNT_DB"></label></td>
								<td><label id="lblAMOUNT_CR"></label></td>
							</tr>
						</tbody>
					</table>
				</div>
			</div>
		</div>
	</div>
							
<script>

$(document).ready(function() {
    App.init();
	
	   table = $('#table').DataTable({
        "dom": "Bfrtip",
        "buttons": [{
            extend: "copy",
            className: "btn-xs btn-grey"
        }, {
            extend: "csv",
            className: "btn-xs btn-indigo"
        }, {
            extend: "excel",
            className: "btn-xs btn-green"
        }, {
            extend: "pdf",
            className: "btn-xs btn-danger"
        }, {
            extend: "print",
            className: "btn-xs btn-success"
        }],
        "processing": true, //Feature control the processing indicator.
        "serverSide": false, //Feature control DataTables' server-side processing mode.
        "order": [], //Initial no order.
        "pageLength": 25,
        "scrollY": "400px",
        "scrollX": true,
        "scrollCollapse": true,
        "columnDefs": [{
            "targets": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11], //last column
            "orderable": true, //set not orderable
        }, ],
    });

	$('#table tbody').on('click', 'tr', function () {
        var data = table.row( this ).data();
		
		//Ajax Load data from ajax
			$.ajax({
				url : "http://iplas.imcplantations.com:8080/IPLASAPI//costactivity/readbyreferenceno.php?fcba="+data[2]+"&referenceno="+data[7],
				type: "GET",
				dataType: "JSON",
				success: function(data)
				{
					$.each(data, function(i, data) {
						$('#lblJOBCODE_DB').text(data.JOBCODE_DB);
						$('#lblJOBCODE_CR').text(data.JOBCODE_CR);
						$('#lblFCNAME_DB').text(data.FCNAME_DB);
						$('#lblFCNAME_CR').text(data.FCNAME_CR);
						$('#lblSOURCETYPE_DB').text(data.SOURCETYPE_DB);
						$('#lblSOURCETYPE_CR').text(data.SOURCETYPE_CR);
						$('#lblSOURCECODE_DB').text(data.SOURCECODE_DB);
						$('#lblSOURCECODE_CR').text(data.SOURCECODE_CR);
						$('#lblTARGETCODE_DB').text(data.TARGETCODE_DB);
						$('#lblTARGETCODE_CR').text(data.TARGETCODE_CR);
						$('#lblREFERENCENO_DB').text(data.REFERENCENOCR);
						$('#lblREFERENCENO_CR').text(data.REFERENCENOCR);
						$('#lblAMOUNT_DB').text(data.AMOUNT_DB.toString().replace(/(\d)(?=(\d\d\d)+(?!\d))/g, "$1,"));
						$('#lblAMOUNT_CR').text(data.AMOUNT_CR.toString().replace(/(\d)(?=(\d\d\d)+(?!\d))/g, "$1,"));
						$('#modal-dialog').modal('show');
					});		
							
				},
				error: function (jqXHR, textStatus, errorThrown)
				{
					alert('Error get data from ajax');
				}
			});
    });

});


function exitReport() {
    window.location.href = "https://iplas.imcplantations.com/Report/Dashboard.asp";
}

</script>
</body>
</html>