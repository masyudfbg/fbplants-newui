<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/db_param.asp"-->
<!--#include virtual="/md5.asp"-->
<%
' Bersihkan password lama dari cookie "ingat saya": cookie ditulis ulang hanya berisi user & BA
If Request.Cookies("myCookies")("txtPassword") & "" <> "" Then
	Dim ckUser, ckBA
	ckUser = Request.Cookies("myCookies")("txtUser")
	ckBA = Request.Cookies("myCookies")("txtBA")
	Response.Cookies("myCookies") = ""   ' kosongkan dulu: ASP ikut menyalin sub-key lama dari cookie request
	Response.Cookies("myCookies")("txtUser") = ckUser
	Response.Cookies("myCookies")("txtBA") = ckBA
	Response.Cookies("myCookies").Expires = Now() + 365
End If
%>
<!doctype html>
<html lang="en">
	<head>
		<meta name="robots" content="noindex, nofollow">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		<meta name="theme-color" content="#1B7863">
		<link rel="icon" href="/IMAGES/icons.png" type="image/png"/>
		<title>Sign in &middot; FB-Plants</title>
		<link rel="preconnect" href="https://fonts.googleapis.com">
		<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
		<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
		<link href="css/login/font-awesome.min.css" rel="stylesheet" type="text/css">
		<style>
			:root {
				--brand: #1B7863;
				--brand-dark: #12584a;
				--teal: #1F9E8E;
				--accent: #F7BD28;
				--ink: #17262a;
				--muted: #5d6d69;
				--line: #d5dfdc;
				--field: #ffffff;
				--bg: #f3f6f5;
				--danger: #b42318;
				--danger-bg: #fef3f2;
				--warn: #8a5a00;
				--warn-bg: #fff8e6;
				--radius: 10px;
			}
			* { box-sizing: border-box; }
			html, body { height: 100%; }
			body {
				margin: 0;
				background: var(--bg);
				color: var(--ink);
				font-family: 'Plus Jakarta Sans', 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
				font-size: 15px;
				line-height: 1.5;
				-webkit-font-smoothing: antialiased;
			}

			.page { display: flex; min-height: 100vh; }

			/* ---------- Panel brand (kiri) ---------- */
			.brand-panel {
				position: relative;
				flex: 0 0 44%;
				max-width: 640px;
				padding: 48px 56px;
				display: flex;
				flex-direction: column;
				justify-content: space-between;
				color: #fff;
				background: linear-gradient(155deg, var(--brand) 0%, var(--brand-dark) 70%);
				overflow: hidden;
			}
			.brand-panel .shape { position: absolute; pointer-events: none; }
			.brand-panel .shape-1 { right: -120px; top: -90px; width: 420px; opacity: .10; }
			.brand-panel .shape-2 { left: -80px; bottom: 120px; width: 260px; opacity: .08; }
			.brand-top { position: relative; display: flex; align-items: center; gap: 12px; font-weight: 600; letter-spacing: .2px; }
			.brand-top img { width: 44px; height: 44px; border-radius: 10px; background: #f7f7f7; padding: 4px; }
			.brand-copy { position: relative; max-width: 440px; }
			.brand-copy .eyebrow {
				display: inline-block;
				font-size: 12px; font-weight: 600; letter-spacing: 1.2px; text-transform: uppercase;
				color: var(--accent);
				margin-bottom: 14px;
			}
			.brand-copy h1 { margin: 0 0 14px; font-size: 38px; line-height: 1.15; font-weight: 700; }
			.brand-copy p { margin: 0; font-size: 16px; color: rgba(255,255,255,.82); }
			.group { position: relative; }
			.group-label { font-size: 12px; letter-spacing: 1px; text-transform: uppercase; color: rgba(255,255,255,.7); margin-bottom: 12px; }
			.group-logos { display: flex; flex-wrap: wrap; gap: 10px; }
			.group-logos span {
				display: flex; align-items: center; justify-content: center;
				height: 44px; padding: 6px 10px;
				background: #fff; border-radius: 8px;
			}
			.group-logos img { max-height: 30px; max-width: 70px; display: block; }

			/* ---------- Panel form (kanan) ---------- */
			.form-panel {
				flex: 1;
				display: flex;
				flex-direction: column;
				align-items: center;
				justify-content: center;
				padding: 40px 24px;
			}
			.card {
				width: 100%;
				max-width: 420px;
				background: #fff;
				border: 1px solid var(--line);
				border-radius: 16px;
				padding: 36px 32px 32px;
				box-shadow: 0 12px 40px rgba(23, 38, 42, .06);
			}
			.card-logo { display: none; text-align: center; margin-bottom: 12px; }
			.card-logo img { width: 64px; height: 64px; border-radius: 14px; }
			.card h2 { margin: 0; font-size: 26px; font-weight: 700; }
			.card .subtitle { margin: 4px 0 24px; color: var(--muted); }

			.alert {
				display: flex; gap: 10px; align-items: flex-start;
				padding: 12px 14px; margin-bottom: 20px;
				border-radius: var(--radius);
				background: var(--danger-bg);
				border: 1px solid #fecdca;
				color: var(--danger);
				font-size: 14px;
			}
			.alert .fa { margin-top: 3px; }

			.field { margin-bottom: 18px; }
			.field label { display: block; font-size: 14px; font-weight: 600; margin-bottom: 6px; }
			.control { position: relative; }
			.control > .fa {
				position: absolute; left: 14px; top: 50%; transform: translateY(-50%);
				color: #8a9a96; font-size: 16px; pointer-events: none; width: 16px; text-align: center;
			}
			.control input, .control select {
				width: 100%;
				height: 46px;
				padding: 0 14px 0 42px;
				font: inherit;
				font-size: 16px;
				color: var(--ink);
				background: var(--field);
				border: 1px solid var(--line);
				border-radius: var(--radius);
				outline: none;
				transition: border-color .15s, box-shadow .15s;
				-webkit-appearance: none; appearance: none;
			}
			.control select {
				padding-right: 40px;
				cursor: pointer;
				background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1.5l5 5 5-5' fill='none' stroke='%235d6d69' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
				background-repeat: no-repeat;
				background-position: right 14px center;
			}
			.control select:disabled { background-color: #f5f7f6; color: #8a9a96; cursor: not-allowed; }
			.control input:focus, .control select:focus {
				border-color: var(--brand);
				box-shadow: 0 0 0 3px rgba(27, 120, 99, .18);
			}
			.control input.has-toggle { padding-right: 48px; }
			.toggle-pass {
				position: absolute; right: 4px; top: 50%; transform: translateY(-50%);
				width: 40px; height: 38px;
				border: 0; background: transparent; border-radius: 8px;
				color: #6b7b77; cursor: pointer; font-size: 17px;
			}
			.toggle-pass:hover { color: var(--ink); background: #f1f4f3; }
			.toggle-pass:focus-visible { outline: 2px solid var(--brand); outline-offset: 1px; }

			.field.invalid input, .field.invalid select { border-color: var(--danger); }
			.field.invalid input:focus, .field.invalid select:focus { box-shadow: 0 0 0 3px rgba(180, 35, 24, .15); }
			.field-error { display: none; margin-top: 6px; font-size: 13px; color: var(--danger); }
			.field.invalid .field-error { display: block; }
			.hint { margin-top: 6px; font-size: 13px; color: var(--muted); }
			.hint.warn {
				display: none;
				color: var(--warn); background: var(--warn-bg);
				padding: 6px 10px; border-radius: 8px;
			}
			.hint.warn.show { display: block; }
			.hint .retry { color: var(--brand); font-weight: 600; background: none; border: 0; padding: 0; font: inherit; cursor: pointer; text-decoration: underline; }


			.btn {
				width: 100%;
				height: 48px;
				margin-top: 6px;
				display: flex; align-items: center; justify-content: center; gap: 10px;
				font: inherit; font-size: 16px; font-weight: 600;
				color: #fff;
				background: var(--brand);
				border: 0; border-radius: var(--radius);
				cursor: pointer;
				transition: background .15s, transform .05s;
			}
			.btn:hover { background: var(--brand-dark); }
			.btn:active { transform: translateY(1px); }
			.btn:focus-visible { outline: 3px solid rgba(27, 120, 99, .35); outline-offset: 2px; }
			.btn[disabled] { opacity: .75; cursor: progress; }
			.btn .fa-spinner { display: none; }
			.btn.loading .fa-spinner { display: inline-block; }

			/* ---------- Dropdown dengan pencarian ---------- */
			.combo-native {
				position: absolute !important; width: 1px !important; height: 1px !important;
				padding: 0 !important; margin: -1px; overflow: hidden; clip: rect(0 0 0 0); border: 0 !important;
			}
			.combo-trigger {
				width: 100%; height: 46px;
				padding: 0 40px 0 42px;
				display: flex; align-items: center;
				font: inherit; font-size: 16px; text-align: left;
				color: var(--ink);
				background: var(--field) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1.5l5 5 5-5' fill='none' stroke='%235d6d69' stroke-width='1.6' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E") no-repeat right 14px center;
				border: 1px solid var(--line);
				border-radius: var(--radius);
				outline: none; cursor: pointer;
				transition: border-color .15s, box-shadow .15s;
			}
			.combo-trigger span { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
			.combo-trigger.placeholder { color: #8a9a96; }
			.combo-trigger:focus, .combo.open .combo-trigger { border-color: var(--brand); box-shadow: 0 0 0 3px rgba(27, 120, 99, .18); }
			.combo-trigger:disabled { background-color: #f5f7f6; color: #8a9a96; cursor: not-allowed; }
			.field.invalid .combo-trigger { border-color: var(--danger); }
			.combo-panel {
				display: none;
				position: absolute; left: 0; right: 0; top: calc(100% + 6px); z-index: 30;
				background: #fff;
				border: 1px solid var(--line);
				border-radius: var(--radius);
				box-shadow: 0 16px 40px rgba(23, 38, 42, .14);
				overflow: hidden;
			}
			.combo.open .combo-panel { display: block; }
			.combo-search { position: relative; padding: 8px; border-bottom: 1px solid var(--line); }
			.combo-search .fa { position: absolute; left: 20px; top: 50%; transform: translateY(-50%); color: #8a9a96; }
			.combo-search input {
				width: 100%; height: 40px;
				padding: 0 12px 0 36px;
				font: inherit; font-size: 16px; color: var(--ink);
				background: #f5f7f6;
				border: 1px solid transparent; border-radius: 8px; outline: none;
			}
			.combo-search input:focus { background: #fff; border-color: var(--brand); }
			.combo-list { list-style: none; margin: 0; padding: 4px; max-height: 260px; overflow-y: auto; }
			.combo-list li {
				padding: 9px 12px; border-radius: 6px;
				font-size: 14px; cursor: pointer;
				display: flex; justify-content: space-between; align-items: center; gap: 8px;
			}
			.combo-list li.active { background: #eaf4f1; }
			.combo-list li[aria-selected="true"] { font-weight: 600; color: var(--brand); }
			.combo-list li[aria-selected="true"]::after { content: "\f00c"; font-family: FontAwesome; font-weight: normal; }
			.combo-list mark { background: #fdeeb8; color: inherit; border-radius: 2px; padding: 0 1px; }
			.combo-empty { padding: 14px 12px; font-size: 14px; color: var(--muted); text-align: center; }

			.footnote { margin-top: 24px; text-align: center; font-size: 13px; color: var(--muted); }
			.mobile-group { display: none; }

			/* ---------- Responsif ---------- */
			@media (max-width: 980px) {
				.brand-panel { display: none; }
				.card-logo { display: block; }
				.card h2, .card .subtitle { text-align: center; }
				.mobile-group { display: flex; flex-wrap: wrap; justify-content: center; gap: 8px; margin-top: 20px; }
				.mobile-group img { height: 26px; max-width: 64px; border-radius: 4px; background: #fff; padding: 3px; }
			}
			@media (max-width: 480px) {
				.form-panel { padding: 24px 16px; justify-content: flex-start; }
				.card { padding: 28px 20px 24px; border-radius: 14px; }
			}
		</style>
	</head>


  <%



  '--ATAS ISI--
  Dim txtUser,txtBA,txtPassword,q,sesCompany, tempStringUsers, tempStringFcba, rsFcba,firstAddress_,txtBU,rsCompany,txtCompany,loginMsg
  txtUser=Request("txtUser")
  txtPassword=Request("txtPassword")
  txtBU=Request("businessunit")
  txtCompany=Request("company")

	Session("fcBa")= ""
	Session("Crop")=""
	Session("BA") = ""
	Session("COMPANY")=""
	Session("FCMANAGER") = ""
	Session("KTU") = ""
	Session("Estate_Mill") = ""
	Session("v_Bank")=""
	Session("v_Purchase")=""
	Session("v_Stores")=""
	Session("v_Employee")=""
	Session("v_Workshop")=""
	Session("v_Vehicle")=""
	Session("v_Contract")=""
	Session("v_Asset")=""
	Session("v_Nursery")=""
	Session("v_Project")=""
	Session("v_Building")=""
	Session("v_General_Charges")=""
	Session("v_Reversed")=""
	Session("v_Costbook")=""
	Session("v_Harvesting")=""
	Session("v_ERPReport")=""
	Session("v_WeighBridge")=""
	Session("v_Mill_Production")=""
	Session("v_Delivery")=""
	Session("v_Sales")=""
    Session("USERNAME") = ""
	Session("CodeName") = ""

	Session("cls_Bank")= ""
	Session("cls_Purchase")= ""
	Session("cls_Stores")= ""
	Session("cls_Employee")= ""
	Session("cls_Workshop")= ""
	Session("cls_Vehicle")= ""
	Session("cls_Contract")= ""
	Session("cls_Asset")=""
	Session("cls_Nursery")= ""
	Session("cls_Project")= ""
	Session("cls_Building")= ""
	Session("cls_General_Charges")= ""
	Session("cls_Reversed")= ""	
	Session("cls_Costbook")= ""
	Session("cls_ERPReport")= ""	
	Session("cls_Harvesting")= ""
	Session("cls_WeighBridge")= ""
	Session("cls_Mill_Production")= ""
	Session("cls_Delivery")= ""
	Session("cls_Sales")= ""
	
	Session("NextMonth") =""
	
	
	
	' Add By Akbar  28 Desember 2015
	
	Session("CurP_Bank")=""
	Session("CurP_Purchase")=""
	Session("CurP_Stores")=""
	Session("CurP_Employee")=""
	Session("CurP_Workshop")=""
	Session("CurP_Vehicle")=""
	Session("CurP_Contract")=""
	Session("CurP_Asset")=""
	Session("CurP_Nursery")=""
	Session("CurP_Project")=""
	Session("CurP_Building")=""
	Session("CurP_General_Charges")=""
	Session("CurP_Reversed")=""
	Session("CurP_Costbook")=""
	Session("CurP_Harvesting")=""
	Session("CurP_ERPReport")=""
	Session("CurP_WeighBridge")=""
	Session("CurP_Mill_Production")=""
	Session("CurP_Delivery")=""
	Session("CurP_Sales")=""
  

	Session("ClsP_Bank")= ""
	Session("ClsP_Purchase")= ""
	Session("ClsP_Stores")= ""
	Session("ClsP_Employee")= ""
	Session("ClsP_Workshop")= ""
	Session("ClsP_Vehicle")= ""
	Session("ClsP_Contract")= ""
	Session("ClsP_Asset")= ""
	Session("ClsP_Nursery")= ""
	Session("ClsP_Project")= ""
	Session("ClsP_Building")= ""
	Session("ClsP_General_Charges")= ""
	Session("ClsP_Reversed")= ""	
	Session("ClsP_Costbook")= ""
	Session("ClsP_ERPReport")= ""	
	Session("ClsP_Harvesting")= ""
	Session("ClsP_WeighBridge")= ""
	Session("ClsP_Mill_Production")= ""
	Session("ClsP_Delivery")= ""
	Session("ClsP_Sales")= ""
	
	 
 IF Request("cmdok")="cmdok" and txtuser<>"" and txtpassword<>"" then

  ' The Query has been modify as per NEW Concept, By NAF, 1 May 2013  
  ' Perbaikan keamanan: username/password dikirim sebagai parameter (dulu disambung ke SQL,
  ' sehingga username seperti  admin' --  bisa login tanpa password).
  tempStringUsers = "Select fccode, fullname,firstAddress from users_tab where fccode = ? and " & _
                    "fcpassword = ? and isActive = 'TRUE' and sysdate between Valid_from and Valid_Until "
  
  Set rs = DbQuery(tempStringUsers, Array(txtUser, md5(txtPassword)))
  if not rs.eof then
    Session("USERNAME") = RS("FCCODE")
	Session("CodeName") = RS("fccode")
	session("FULLNAME") = RS("fullname")
	'add by akbar 24 feb 2014
	
	firstAddress_ = RS("firstAddress")
	
	tempStringFcba = "select fcba from user_fcba_tab where fcusercode = ? and fcba = ? "
	Set rsFcba = DbQuery(tempStringFcba, Array(RS("FCCODE"), txtBU))
	
	if not rsFcba.eof then
		txtBA = rsFcba("fcba")
    else
	    txtBA = ""
		loginMsg = "You don't have access to business unit <b>" & Server.HTMLEncode(txtBU) & "</b>. Please contact your administrator."
		' Perbaikan keamanan: dulu sesi login (USERNAME) tetap terisi walau BA ditolak
		Session("USERNAME") = "" : Session("CodeName") = "" : Session("FULLNAME") = ""
	end if
	rsFcba.Close
	set rsFcba = Nothing
	
	' Pick the Default of FCBA and Pick the Default of Company, by NAF, 1 May 2013
	if txtBA <> "" then
	DbExec "UPDATE user_fcba_tab SET isdefault = 'FALSE' where fcusercode = ? ", Array(RS("FCCODE"))
	
	DbExec "UPDATE user_fcba_tab SET isdefault = 'TRUE' where fcusercode = ? and fcba = ? ", Array(RS("FCCODE"), txtBA)
	end if
	
	
	'Set rsFcba = connStr.Execute(tempStringFcba)
	
	'if not rsFcba.eof then
'		txtBA = rsFcba("fcba")
'    else
'	    txtBA = ""
		'Response.write "You must have at least one Default Business Unit..., Please check it with your Administrator"
	'end if
	
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE,companyno  from BusinessUnit a INNER JOIN COMPANY b ON a.fccompanycode = b.companycode WHERE fcCode = ?"
    set rsCat = DbQuery(q, Array(txtBA))
	if not rsCat.eof then
	
	   Session("fcBa")= rsCat("fcCode")
	   Session("BA") = rsCat("fcName")
	   Session("COMPANYCODE")=rsCat("fccompanycode")
	   Session("COMPANY")=rsCat("FCCOMPANYNAME")
	   Session("FCMANAGER") = rsCat("FCMANAGER")
	   Session("KTU") = rsCat("FCKTU")
	   Session("Crop")=rsCat("FCCROP")
	   Session("Estate_Mill")=rsCat("fcType")
	   Session("COMPANYNO")=rsCat("companyno")
	   
	   
	   Session("v_Bank")=cmdFindPeriod("Bank")
	   Session("v_Purchase")=cmdFindPeriod("Purchase")
	   Session("v_Stores")=cmdFindPeriod("Stores")
	   Session("v_Employee")=cmdFindPeriod("Employee")
	   Session("v_Workshop")=cmdFindPeriod("Workshop")
	   Session("v_Vehicle")=cmdFindPeriod("Vehicle")
	   Session("v_Contract")=cmdFindPeriod("Contract")
	   Session("v_Asset")=cmdFindPeriod("Asset")
	   Session("v_Nursery")=cmdFindPeriod("Nursery")
	   Session("v_Project")=cmdFindPeriod("Project")
	   Session("v_Building")=cmdFindPeriod("Building")
	   Session("v_General_Charges")=cmdFindPeriod("General_Charges")
	   Session("v_Reversed")=cmdFindPeriod("Reversed")
	   Session("v_Costbook")=cmdFindPeriod("Costbook")
	   Session("v_Harvesting")=cmdFindPeriod("Harvesting")
	   Session("v_WeighBridge")=cmdFindPeriod("WeighBridge") 
	   Session("v_Mill_Production")=cmdFindPeriod("Mill_Production") 
	   Session("v_Delivery")=cmdFindPeriod("Delivery") 
	   Session("v_ERPReport")=cmdFindPeriod("ERPReport")

	   Session("cls_Bank")=cmdPeriodClose("Bank")
	   Session("cls_Purchase")=cmdPeriodClose("Purchase")
	   Session("cls_Stores")=cmdPeriodClose("Stores")
	   Session("cls_Employee")=cmdPeriodClose("Employee")
	   Session("cls_Workshop")=cmdPeriodClose("Workshop")
	   Session("cls_Vehicle")=cmdPeriodClose("Vehicle")
	   Session("cls_Contract")=cmdPeriodClose("Contract")
	   Session("cls_Asset")=cmdPeriodClose("Asset")
	   Session("cls_Nursery")=cmdPeriodClose("Nursery")
	   Session("cls_Project")=cmdPeriodClose("Project")
	   Session("cls_Building")=cmdPeriodClose("Building")
	   Session("cls_General_Charges")=cmdPeriodClose("General_Charges")
	   Session("cls_Reversed")=cmdPeriodClose("Reversed")
	   Session("cls_Costbook")=cmdPeriodClose("Costbook")
	   Session("cls_Harvesting")=cmdPeriodClose("Harvesting")
	   Session("cls_WeighBridge")=cmdPeriodClose("WeighBridge") 
	   Session("cls_Mill_Production")=cmdPeriodClose("Mill_Production") 
	   Session("cls_Delivery")=cmdPeriodClose("Delivery") 
	   Session("cls_ERPReport")=cmdPeriodClose("ERPReport")
	   
       Session("NextMonth") = cmdPeriodClose("Financial")
	   
	 ' Add by Ali Akbar  28 Desember 2015  -----
	 
	   Session("CurP_Bank")=cmdFindAccountPeriod("Bank")
	   Session("CurP_Purchase")=cmdFindAccountPeriod("Purchase")
	   Session("CurP_Stores")=cmdFindAccountPeriod("Stores")
	   Session("CurP_Employee")=cmdFindAccountPeriod("Employee")
	   Session("CurP_Workshop")=cmdFindAccountPeriod("Workshop")
	   Session("CurP_Vehicle")=cmdFindAccountPeriod("Vehicle")
	   Session("CurP_Contract")=cmdFindAccountPeriod("Contract")
	   Session("CurP_Asset")=cmdFindAccountPeriod("Asset")
	   Session("CurP_Nursery")=cmdFindAccountPeriod("Nursery")
	   Session("CurP_Project")=cmdFindAccountPeriod("Project")
	   Session("CurP_Building")=cmdFindAccountPeriod("Building")
	   Session("CurP_General_Charges")=cmdFindAccountPeriod("General_Charges")
	   Session("CurP_Reversed")=cmdFindAccountPeriod("Reversed")
	   Session("CurP_Costbook")=cmdFindAccountPeriod("Costbook")
	   Session("CurP_Harvesting")=cmdFindAccountPeriod("Harvesting")
	   Session("CurP_WeighBridge")=cmdFindAccountPeriod("WeighBridge") 
	   Session("CurP_Mill_Production")=cmdFindAccountPeriod("Mill_Production") 
	   Session("CurP_Delivery")=cmdFindAccountPeriod("Delivery") 
	   Session("CurP_ERPReport")=cmdFindAccountPeriod("ERPReport")

	   Session("ClsP_Bank")=cmdAccountPeriodClose("Bank")
	   Session("ClsP_Purchase")=cmdAccountPeriodClose("Purchase")
	   Session("ClsP_Stores")=cmdAccountPeriodClose("Stores")
	   Session("ClsP_Employee")=cmdAccountPeriodClose("Employee")
	   Session("ClsP_Workshop")=cmdAccountPeriodClose("Workshop")
	   Session("ClsP_Vehicle")=cmdAccountPeriodClose("Vehicle")
	   Session("ClsP_Contract")=cmdAccountPeriodClose("Contract")
	   Session("ClsP_Asset")=cmdAccountPeriodClose("Asset")
	   Session("ClsP_Nursery")=cmdAccountPeriodClose("Nursery")
	   Session("ClsP_Project")=cmdAccountPeriodClose("Project")
	   Session("ClsP_Building")=cmdAccountPeriodClose("Building")
	   Session("ClsP_General_Charges")=cmdAccountPeriodClose("General_Charges")
	   Session("ClsP_Reversed")=cmdAccountPeriodClose("Reversed")
	   Session("ClsP_Costbook")=cmdAccountPeriodClose("Costbook")
	   Session("ClsP_Harvesting")=cmdAccountPeriodClose("Harvesting")
	   Session("ClsP_WeighBridge")=cmdAccountPeriodClose("WeighBridge") 
	   Session("ClsP_Mill_Production")=cmdAccountPeriodClose("Mill_Production") 
	   Session("ClsP_Delivery")=cmdAccountPeriodClose("Delivery") 
	   Session("ClsP_ERPReport")=cmdAccountPeriodClose("ERPReport")
	   
	   
	   
	   if request("persistent") = "y" Then
		  response.Cookies("myCookies") = ""   ' kosongkan dulu: ASP ikut menyalin sub-key lama (termasuk password) dari cookie request
		  response.Cookies("myCookies")("txtUser") = trim(request("txtUser"))
		  response.Cookies("myCookies")("txtBA") = trim(request("txtBA"))
		  ' Keamanan: password TIDAK disimpan di cookie (dulu password asli tersimpan 365 hari)
		  response.Cookies("myCookies").Expires = Now() + 365
	   End if
	   q="Insert into HISTORICAL_USER_TAB " & _
	        "( USERNAME, IP_ADDRESS, BU, DATE_ACT,DATE_ACT_FORM, TIME_ACT, LASTUPDATE, LASTTIME ) " & _
	       	"values (?, ?, ?, ?, sysdate, to_char(sysdate,'HH24:MI'), sysdate, to_char(sysdate,'HH24:MI'))"
	  DbExec q, Array(Session("CodeName"), Request.ServerVariables("REMOTE_ADDR"), Session("fcBA"), FormatDateTime(Date, 1))
	    If len(trim(firstAddress_))	 > 1 then
	       Response.Redirect(firstAddress_)
	    Else
		  Response.Redirect("login.asp")	  
   	    End If
	end if
  Else
	loginMsg = "Incorrect username or password. Please check and try again."
  End if
  rs.close()
  rs=null
  End if

  ' Isi ulang username dari input terakhir (atau cookie lama); password tidak pernah dicetak ulang
  if txtUser = "" then txtUser = Request.Cookies("myCookies")("txtUser")
  %>
<body>
	<div class="page">
		<aside class="brand-panel" aria-hidden="true">
			<svg class="shape shape-1" viewBox="0 0 100 100"><path d="M50 6 L94 52 L80 66 L50 35 L20 66 L6 52 Z" fill="#fff"/></svg>
			<svg class="shape shape-2" viewBox="0 0 100 100"><rect x="25" y="25" width="50" height="50" transform="rotate(45 50 50)" fill="#F7BD28"/></svg>

			<div class="brand-top">
				<img src="css/login/images/logofbg.jpeg" alt="">
				<span>First Borneo Group</span>
			</div>

			<div class="brand-copy">
				<span class="eyebrow">Plantation System</span>
				<h1>FB-Plants</h1>
				<p>One place to manage estate, mill, purchasing, and finance operations across every business unit.</p>
			</div>

			<div class="group">
				<div class="group-label">Group companies</div>
				<div class="group-logos">
					<span><img src="css/login/images/logoarj.jpeg" alt=""></span>
					<span><img src="css/login/images/logobrs.jpg" alt=""></span>
					<span><img src="css/login/images/logobia.jpeg" alt=""></span>
					<span><img src="css/login/images/logokaa.jpg" alt=""></span>
					<span><img src="css/login/images/logocpp.png" alt=""></span>
					<span><img src="css/login/images/logogas.jpeg" alt=""></span>
					<span><img src="css/login/images/logoksr.jpg" alt=""></span>
				</div>
			</div>
		</aside>

		<main class="form-panel">
			<div class="card">
				<div class="card-logo"><img src="css/login/images/logofbg.jpeg" alt="First Borneo Group"></div>
				<h2>Welcome back</h2>
				<p class="subtitle">Sign in to continue to FB-Plants.</p>

				<% if loginMsg <> "" then %>
				<div class="alert" role="alert">
					<i class="fa fa-exclamation-circle" aria-hidden="true"></i>
					<div><%=loginMsg%></div>
				</div>
				<% end if %>

				<form action="logout.asp" method="post" name="frmLogin" id="frmLogin" novalidate
					data-company="<%=Server.HTMLEncode(txtCompany & "")%>"
					data-bu="<%=Server.HTMLEncode(txtBU & "")%>">

					<div class="field">
						<label for="usn">Username</label>
						<div class="control">
							<i class="fa fa-user" aria-hidden="true"></i>
							<input type="text" name="txtUser" id="usn" maxlength="30" value="<%=Server.HTMLEncode(txtUser & "")%>"
								autocomplete="username" autocapitalize="off" autocorrect="off" spellcheck="false"
								placeholder="Enter your username" aria-describedby="usn-err">
						</div>
						<div class="field-error" id="usn-err">Please enter your username.</div>
					</div>

					<div class="field">
						<label for="pwd">Password</label>
						<div class="control">
							<i class="fa fa-lock" aria-hidden="true"></i>
							<input type="password" name="txtPassword" id="pwd" maxlength="30" class="has-toggle"
								autocomplete="current-password" placeholder="Enter your password" aria-describedby="pwd-err pwd-caps pwd-hint">
							<button type="button" class="toggle-pass" id="togglePass" aria-label="Show password" aria-pressed="false">
								<i class="fa fa-eye" aria-hidden="true"></i>
							</button>
						</div>
						<div class="field-error" id="pwd-err">Please enter your password.</div>
						<div class="hint warn" id="pwd-caps" aria-live="polite"><i class="fa fa-arrow-up" aria-hidden="true"></i> Caps Lock is on</div>
						<div class="hint" id="pwd-hint">Username and password are case-sensitive.</div>
					</div>

					<div class="field">
						<label for="company">Company</label>
						<div class="control">
							<i class="fa fa-building-o" aria-hidden="true"></i>
							<select name="company" id="company" aria-describedby="company-err">
								<option value="">Select a company</option>
								<%
								q = " SELECT companycode, companyname " & _
									" FROM company " & _
									" order by companyname"
								Set rsCompany = Server.CreateObject("ADODB.Recordset")
								rsCompany.CursorType = 3
								rsCompany.CursorLocation = 3
								rsCompany.Open q, connstr
								Do While Not rsCompany.eof
								%>
								<option value="<%=Server.HTMLEncode(rsCompany("companycode") & "")%>"<%if cstr("" & txtCompany) = cstr("" & rsCompany("companycode")) then %> selected<% end if %>><%=Server.HTMLEncode(rsCompany("companycode") & "")%> &ndash; <%=Server.HTMLEncode(rsCompany("companyname") & "")%></option>
								<%
									rsCompany.movenext
								Loop
								rsCompany.Close
								set rsCompany = Nothing
								q = ""
								%>
							</select>
						</div>
						<div class="field-error" id="company-err">Please select your company.</div>
					</div>

					<div class="field">
						<label for="businessunit">Business unit</label>
						<div class="control">
							<i class="fa fa-briefcase" aria-hidden="true"></i>
							<select name="businessunit" id="businessunit" disabled aria-describedby="bu-err bu-hint">
								<option value="">Select a company first</option>
							</select>
						</div>
						<div class="field-error" id="bu-err">Please select a business unit.</div>
						<div class="hint" id="bu-hint" aria-live="polite"></div>
					</div>

					<button type="submit" class="btn" id="btn_signin">
						<i class="fa fa-spinner fa-spin" aria-hidden="true"></i>
						<span>Sign in</span>
					</button>
					<input type="hidden" value="cmdok" name="cmdok">
				</form>
			</div>

			<div class="mobile-group" aria-hidden="true">
				<img src="css/login/images/logoarj.jpeg" alt="">
				<img src="css/login/images/logobrs.jpg" alt="">
				<img src="css/login/images/logobia.jpeg" alt="">
				<img src="css/login/images/logokaa.jpg" alt="">
				<img src="css/login/images/logocpp.png" alt="">
				<img src="css/login/images/logogas.jpeg" alt="">
				<img src="css/login/images/logoksr.jpg" alt="">
			</div>

			<p class="footnote">&copy; <%=Year(Date)%> First Borneo Group &middot; Trouble signing in? Contact your IT administrator.</p>
		</main>
	</div>

<script type="text/javascript">
(function () {
	var form = document.getElementById('frmLogin');
	var usn = document.getElementById('usn');
	var pwd = document.getElementById('pwd');
	var company = document.getElementById('company');
	var bu = document.getElementById('businessunit');
	var buHint = document.getElementById('bu-hint');
	var btn = document.getElementById('btn_signin');
	var toggle = document.getElementById('togglePass');
	var caps = document.getElementById('pwd-caps');

	// Simpan pilihan company/BU terakhir di browser (opsional, aman bila storage diblokir)
	function save(key, val) { try { localStorage.setItem(key, val); } catch (e) {} }
	function load(key) { try { return localStorage.getItem(key); } catch (e) { return null; } }

	// Ubah <select> menjadi dropdown dengan kolom pencarian.
	// <select> asli tetap dipakai untuk submit form & event 'change'.
	function makeSearchable(select, searchPlaceholder) {
		var control = select.parentNode;
		var listId = select.id + '-list';
		var wrap = document.createElement('div');
		wrap.className = 'combo';
		control.appendChild(wrap);

		var trigger = document.createElement('button');
		trigger.type = 'button';
		trigger.className = 'combo-trigger';
		trigger.setAttribute('role', 'combobox');
		trigger.setAttribute('aria-haspopup', 'listbox');
		trigger.setAttribute('aria-expanded', 'false');
		trigger.setAttribute('aria-controls', listId);
		trigger.innerHTML = '<span></span>';
		wrap.appendChild(trigger);

		var panel = document.createElement('div');
		panel.className = 'combo-panel';
		panel.innerHTML =
			'<div class="combo-search"><i class="fa fa-search" aria-hidden="true"></i>' +
			'<input type="text" autocomplete="off" spellcheck="false" aria-autocomplete="list"></div>' +
			'<ul class="combo-list" role="listbox"></ul>';
		wrap.appendChild(panel);

		var search = panel.querySelector('input');
		var list = panel.querySelector('ul');
		search.placeholder = searchPlaceholder;
		search.setAttribute('aria-controls', listId);
		list.id = listId;

		// Label tetap terhubung ke kontrol yang terlihat
		var label = document.querySelector('label[for="' + select.id + '"]');
		if (label) {
			label.id = label.id || select.id + '-label';
			trigger.setAttribute('aria-labelledby', label.id);
			list.setAttribute('aria-labelledby', label.id);
			label.addEventListener('click', function (e) { e.preventDefault(); if (!trigger.disabled) trigger.focus(); });
		}
		select.classList.add('combo-native');
		select.tabIndex = -1;
		select.setAttribute('aria-hidden', 'true');

		var items = [];
		var active = -1;

		function escapeHtml(s) {
			return s.replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
		}

		// Tandai semua bagian teks yang cocok dengan kata kunci
		function highlight(text, tokens) {
			var lower = text.toLowerCase();
			var marked = [];
			tokens.forEach(function (t) {
				var from = 0, at;
				while (t && (at = lower.indexOf(t, from)) !== -1) {
					for (var k = at; k < at + t.length; k++) marked[k] = true;
					from = at + t.length;
				}
			});
			var out = '', open = false;
			for (var i = 0; i < text.length; i++) {
				if (marked[i] && !open) { out += '<mark>'; open = true; }
				if (!marked[i] && open) { out += '</mark>'; open = false; }
				out += escapeHtml(text.charAt(i));
			}
			return out + (open ? '</mark>' : '');
		}

		function render() {
			var tokens = search.value.toLowerCase().split(/\s+/).filter(Boolean);
			list.innerHTML = '';
			items = [];
			for (var i = 0; i < select.options.length; i++) {
				var opt = select.options[i];
				if (!opt.value) continue;
				var text = opt.text;
				var lower = text.toLowerCase();
				// Semua kata harus ada (urutan bebas), mis. "borneo first"
				if (!tokens.every(function (t) { return lower.indexOf(t) !== -1; })) continue;
				var li = document.createElement('li');
				li.id = select.id + '-opt-' + i;
				li.setAttribute('role', 'option');
				li.setAttribute('aria-selected', opt.value === select.value ? 'true' : 'false');
				li.innerHTML = '<span>' + highlight(text, tokens) + '</span>';
				li.dataset.value = opt.value;
				list.appendChild(li);
				items.push(li);
			}
			if (!items.length) {
				var empty = document.createElement('li');
				empty.className = 'combo-empty';
				empty.setAttribute('role', 'presentation');
				empty.textContent = 'No results for "' + search.value.trim() + '"';
				list.appendChild(empty);
			}
			var selectedIndex = -1;
			items.forEach(function (li, idx) { if (li.dataset.value === select.value) selectedIndex = idx; });
			setActive(tokens.length ? 0 : Math.max(selectedIndex, 0));
		}

		function setActive(idx) {
			if (active > -1 && items[active]) items[active].classList.remove('active');
			active = items.length ? Math.min(Math.max(idx, 0), items.length - 1) : -1;
			if (active > -1) {
				items[active].classList.add('active');
				search.setAttribute('aria-activedescendant', items[active].id);
				var li = items[active];
				if (li.offsetTop < list.scrollTop) list.scrollTop = li.offsetTop - 4;
				else if (li.offsetTop + li.offsetHeight > list.scrollTop + list.clientHeight) list.scrollTop = li.offsetTop + li.offsetHeight - list.clientHeight + 4;
			} else {
				search.removeAttribute('aria-activedescendant');
			}
		}

		function open(initialText) {
			if (trigger.disabled || wrap.classList.contains('open')) return;
			wrap.classList.add('open');
			trigger.setAttribute('aria-expanded', 'true');
			search.value = initialText || '';
			render();
			search.focus();
		}

		function close(focusTrigger) {
			if (!wrap.classList.contains('open')) return;
			wrap.classList.remove('open');
			trigger.setAttribute('aria-expanded', 'false');
			if (focusTrigger) trigger.focus();
		}

		function choose(value) {
			var changed = select.value !== value;
			select.value = value;
			refresh();
			close(true);
			if (changed) select.dispatchEvent(new Event('change', { bubbles: true }));
		}

		// Samakan tampilan dengan isi & status <select> asli
		function refresh() {
			var opt = select.options[select.selectedIndex];
			var hasValue = !!(opt && opt.value);
			trigger.firstChild.textContent = opt ? opt.text : '';
			trigger.classList.toggle('placeholder', !hasValue);
			trigger.disabled = select.disabled;
			if (select.disabled) close(false);
		}

		trigger.addEventListener('click', function () {
			wrap.classList.contains('open') ? close(true) : open('');
		});
		trigger.addEventListener('keydown', function (e) {
			if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter' || e.key === ' ') {
				e.preventDefault();
				open('');
			} else if (e.key && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
				// Langsung mengetik di dropdown = mulai mencari
				e.preventDefault();
				open(e.key);
			}
		});
		search.addEventListener('input', render);
		search.addEventListener('keydown', function (e) {
			if (e.key === 'ArrowDown') { e.preventDefault(); setActive(active + 1); }
			else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(active - 1); }
			else if (e.key === 'Enter') { e.preventDefault(); if (active > -1) choose(items[active].dataset.value); }
			else if (e.key === 'Escape') { e.preventDefault(); close(true); }
			else if (e.key === 'Tab') { close(false); }
		});
		list.addEventListener('mousedown', function (e) { e.preventDefault(); });
		list.addEventListener('click', function (e) {
			var li = e.target.closest('li[role="option"]');
			if (li) choose(li.dataset.value);
		});
		list.addEventListener('mousemove', function (e) {
			var li = e.target.closest('li[role="option"]');
			if (li && items.indexOf(li) !== active) setActive(items.indexOf(li));
		});
		document.addEventListener('mousedown', function (e) {
			if (!wrap.contains(e.target)) close(false);
		});

		select.addEventListener('change', refresh);
		if (window.MutationObserver) {
			new MutationObserver(refresh).observe(select, { childList: true, attributes: true, attributeFilter: ['disabled'] });
		}
		refresh();
		return { trigger: trigger, refresh: refresh };
	}

	var companyCombo = makeSearchable(company, 'Search company...');
	var buCombo = makeSearchable(bu, 'Search business unit...');

	function setInvalid(input, invalid) {
		input.closest('.field').classList.toggle('invalid', invalid);
		input.setAttribute('aria-invalid', invalid ? 'true' : 'false');
	}

	function setOptions(list, placeholder) {
		bu.innerHTML = '';
		var first = document.createElement('option');
		first.value = '';
		first.textContent = placeholder;
		bu.appendChild(first);
		for (var i = 0; i < list.length; i++) {
			var opt = document.createElement('option');
			opt.value = list[i].id;
			opt.textContent = list[i].text;
			bu.appendChild(opt);
		}
	}

	function loadBusinessUnits(selected) {
		buHint.innerHTML = '';
		setInvalid(bu, false);
		if (!company.value) {
			setOptions([], 'Select a company first');
			bu.disabled = true;
			return;
		}
		setOptions([], 'Loading business units...');
		bu.disabled = true;

		var xhr = new XMLHttpRequest();
		xhr.open('POST', 'getlistfcbaajax.asp?fccompanycode=' + encodeURIComponent(company.value));
		xhr.onload = function () {
			var list;
			try { list = JSON.parse(xhr.responseText).results || []; } catch (e) { list = null; }
			if (xhr.status !== 200 || !list) { return showLoadError(); }

			setOptions(list, list.length ? 'Select a business unit' : 'No business units available');
			bu.disabled = list.length === 0;
			if (selected) {
				for (var i = 0; i < list.length; i++) {
					if (list[i].id === selected) { bu.value = selected; buCombo.refresh(); break; }
				}
			}
		};
		xhr.onerror = showLoadError;
		xhr.send();
	}

	function showLoadError() {
		setOptions([], 'Could not load business units');
		bu.disabled = true;
		buHint.innerHTML = 'Something went wrong while loading business units. <button type="button" class="retry">Try again</button>';
		buHint.querySelector('.retry').onclick = function () { loadBusinessUnits(null); };
	}

	company.addEventListener('change', function () {
		setInvalid(company, false);
		save('fbp_company', company.value);
		loadBusinessUnits(company.value === load('fbp_company_bu_owner') ? load('fbp_bu') : null);
	});

	bu.addEventListener('change', function () {
		setInvalid(bu, false);
		save('fbp_bu', bu.value);
		save('fbp_company_bu_owner', company.value);
	});

	[usn, pwd].forEach(function (input) {
		input.addEventListener('input', function () { if (input.value) setInvalid(input, false); });
	});

	// Tampilkan / sembunyikan password
	toggle.addEventListener('click', function () {
		var show = pwd.type === 'password';
		pwd.type = show ? 'text' : 'password';
		toggle.setAttribute('aria-label', show ? 'Hide password' : 'Show password');
		toggle.setAttribute('aria-pressed', show ? 'true' : 'false');
		toggle.firstElementChild.className = 'fa ' + (show ? 'fa-eye-slash' : 'fa-eye');
		pwd.focus();
	});

	// Peringatan Caps Lock (password case-sensitive)
	function checkCaps(e) {
		if (e.getModifierState) caps.classList.toggle('show', e.getModifierState('CapsLock'));
	}
	pwd.addEventListener('keydown', checkCaps);
	pwd.addEventListener('keyup', checkCaps);
	pwd.addEventListener('blur', function () { caps.classList.remove('show'); });

	form.addEventListener('submit', function (e) {
		var fields = [usn, pwd, company, bu];
		var firstInvalid = null;
		for (var i = 0; i < fields.length; i++) {
			var empty = !fields[i].value;
			setInvalid(fields[i], empty);
			if (empty && !firstInvalid) firstInvalid = fields[i];
		}
		if (firstInvalid) {
			e.preventDefault();
			var target = firstInvalid.disabled ? company : firstInvalid;
			(target === company ? companyCombo.trigger : target === bu ? buCombo.trigger : target).focus();
			return;
		}
		btn.classList.add('loading');
		btn.disabled = true;
		btn.lastElementChild.textContent = 'Signing in...';
	});

	// Kembalikan tombol bila halaman dibuka lagi lewat tombol Back
	window.addEventListener('pageshow', function () {
		btn.classList.remove('loading');
		btn.disabled = false;
		btn.lastElementChild.textContent = 'Sign in';
	});

	// Inisialisasi: pakai pilihan dari submit terakhir, atau yang tersimpan di browser
	var initialCompany = form.getAttribute('data-company') || '';
	var initialBu = form.getAttribute('data-bu') || '';
	if (!initialCompany) {
		var savedCompany = load('fbp_company');
		if (savedCompany && company.querySelector('option[value="' + savedCompany.replace(/"/g, '') + '"]')) {
			company.value = savedCompany;
			companyCombo.refresh();
			if (load('fbp_company_bu_owner') === savedCompany) initialBu = load('fbp_bu') || '';
		}
	}
	if (company.value) loadBusinessUnits(initialBu);

	(usn.value ? pwd : usn).focus();
})();
</script>
</body>
</html>
