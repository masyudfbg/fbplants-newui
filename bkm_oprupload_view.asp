<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim btnApprove,bkmno1,sql,totalpremi,bkmdate
dim ScriptName,bkmno2,fcba,hujan
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")



function Akses_User
Dim rsUsrPosting, strUsrPosting,Akses_Usr
    Akses_Usr  =  0
 	strUsrPosting =	" SELECT 1 result " & _
					" FROM mbkmhdr hdr " & _
					" INNER JOIN authorizetemplate at ON (at.document_type = 'AA' AND at.document_classification = hdr.document_classification)" & _
					" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "')" & _
					" WHERE hdr.bkmno='" & trim(request("bkmno")) & "' " & _
					" AND hdr.rowstate = 'Planned'" & _
					" AND hdr.postingflg = 0" & _
					" AND at.fccode = '501'" & _
					" AND hdr.fcba='" & trim(session("fcba")) & "' " & _
					" AND authorizelevelcode = 900" & _
				    " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' AND sectionname='" & SECTIONNAME & "') " 
 				
 	Set rsUsrPosting = Server.CreateObject("ADODB.Recordset")
 	rsUsrPosting.Open strUsrPosting, connstr, 3
	
 	If Not rsUsrPosting.EOF Then
 		Akses_Usr = cdbl(rsUsrPosting("result"))
 	Else
 		Akses_Usr = 0
 	End If
 	rsUsrPosting.Close
 	set rsUsrPosting = Nothing
 
	Akses_User = Akses_Usr
end function 

function Akses_Posting
Dim rsAvailable, strAvailable,Access_Post
    Access_Post  =  0
 	strAvailable =	" select authorizestructure_api.get_access_post('501','" & businessunit & "','" & Session("CodeName") & "') result" & _
 					" from dual"
		
 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
 	rsAvailable.Open strAvailable, connstr, 3
	
 	If Not rsAvailable.EOF Then
 		Access_Post = cdbl(rsAvailable("result"))
 	Else
 		Access_Post = 0
 	End If
 	rsAvailable.Close
 	set rsAvailable = Nothing
 
	Akses_Posting = Access_Post
end function

dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
bkmno		= trim(request("bkmno"))
businessunit= trim(session("fcba"))

'response.write bkmno

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID,supervisi1,SUPERVISI2,SUPERVISI3,SUPERVISI4,SUPERVISINAME1,SUPERVISINAME2,SUPERVISINAME3,SUPERVISINAME4,Clasification,Clasificationname,ROWSTATE,POSTINGFLG,HUJANCK
sqlheader = "select h.BKMNO,TO_CHAR(h.BKMDATE,'mm/dd/yyyy')BKMDATE,h.GANGCODE,h.SUPERVISI1,(SELECT s1.FCNAME FROM EMPLOYEE  s1 WHERE h.SUPERVISI1=s1.FCCODE and s1.fcba=h.fcba ) AS SUPERVISINAME1, " & _
" h.SUPERVISI2,(SELECT s2.FCNAME FROM EMPLOYEE s2 WHERE h.SUPERVISI2=s2.FCCODE and s2.fcba=h.fcba) AS SUPERVISINAME2,h.SUPERVISI3, " & _
" (SELECT s3.FCNAME FROM EMPLOYEE s3 WHERE h.SUPERVISI3=s3.FCCODE and s3.fcba=h.fcba) AS SUPERVISINAME3,h.SUPERVISI4, " & _
" (SELECT s4.FCNAME FROM EMPLOYEE s4 WHERE h.SUPERVISI4=s4.FCCODE and s4.fcba=h.fcba) AS SUPERVISINAME4,h.SECTIONCODE,h.DOCUMENT_CLASSIFICATION,c.fcname, h.ROWSTATE,h.POSTINGFLG, h.hujan " & _
" from mbkmhdr h,DOCUMENT_CLASSIFICATION c where h.DOCUMENT_CLASSIFICATION=c.FCCODE and h.BKMNO='" & bkmno & "'"
' response.write sqlheader
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bkmnohdr = rsHeader("BKMNO")
	DATECREATED = rsHeader("BKMDATE")
	SECTIONNAME = rsHeader("SECTIONCODE")
	GANGCODE = rsHeader("GANGCODE")
	supervisi1 = rsHeader("SUPERVISI1")
	SUPERVISI2 = rsHeader("SUPERVISI2")
	SUPERVISI3 = rsHeader("SUPERVISI3")
	SUPERVISI4 = rsHeader("SUPERVISI4")
	SUPERVISINAME1 = rsHeader("SUPERVISINAME1")
	SUPERVISINAME2 = rsHeader("SUPERVISINAME2")
	SUPERVISINAME3 = rsHeader("SUPERVISINAME3")
	SUPERVISINAME4 = rsHeader("SUPERVISINAME4")
	Clasification = rsHeader("DOCUMENT_CLASSIFICATION")
	Clasificationname 	= rsHeader("fcname")
	ROWSTATE   = rsHeader("ROWSTATE")
	POSTINGFLG = rsHeader("POSTINGFLG")
	HUJANCK = rsHeader("HUJAN")
	
	
end if

IF POSTINGFLG = "0" THEN 
   POSTINGFLG = "TIDAK"
ELSE
   POSTINGFLG	= "YA"
END IF

'count detailupah
Dim SQL_Rows, NoOfRows,RSCount,supervisinew1

SQL_Rows = "SELECT Count (*) from BKMUPAHDTL " & _ 
" where BKMNO='" & bkmno & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close

'count detailbahan
Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem,nombkm

SQL_Rowsitem = "SELECT Count (*) from BKMBAHANDTL where BKMNO='" & bkmno & "'"
set RSCountitem = connStr.Execute(SQL_Rowsitem)
if not RSCountitem.eof then 
	NoOfRowsitem = RSCountitem(0).value
end if
RSCountitem.close


'proses dibawah ini digunakan untuk approved bkm dan create jurnal
btnApprove 	= request("btnsave")
bkmno1 		= request("bkmno1")
bkmdate 	= request("bkmdate")
hujan 		= Request("hujan")

dim strQueryStatus,modulFile,executeQry,modulFile1,executeQry1
DIM sql1,rs1,sqlexec
nombkm = "M"&bkmno1
if hujan="on" then 
	hujan = 1
else 
	hujan = 0
end if
IF btnApprove = "UPLOAD" then
		supervisinew1  = Request("supervisi1") 

		sqlexec ="UPDATE MBKMHDR SET supervisi1='" & supervisinew1 & "',hujan='" & hujan & "' where bkmno='" & bkmno1 & "'"
		connStr.execute(sqlexec)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & sqlexec
		ExecuteAndLog modulFile,executeQry

	 'Insert header
		sqlexec     = " MERGE INTO BKMHDR a " & _ 
						" USING ( SELECT 'M'||trim(BKMNO) BKMNO,BKMDATE,SECTIONCODE,GANGCODE,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,CREATEDATE,SYSDATE LASTUPDATE,FCENTRY,'" &  Session("CodeName") & "' FCEDIT,'" & Request.ServerVariables("REMOTE_ADDR") & "' FCIP,FCBA,'1' POSTINGFLG,'Approved' ROWSTATE,DOCUMENT_CLASSIFICATION,0 TOTUPAH,0 TOTBAHAN,0 TOTVRA,0 TOTWIP,VEHICLECODE,'" & hujan & "' HUJAN,CONTRACTORCODE " & _
						"         FROM MBKMHDR WHERE BKMNO = '" & bkmno1 & "' " & _
						"         AND FCBA='" & businessunit  & "'" & _
						"       )b ON (a.BKMNO=b.BKMNO and a.fcba=b.fcba) " & _ 
						" WHEN NOT MATCHED THEN " &_ 
						" Insert (BKMNO,BKMDATE,SECTIONCODE,GANGCODE,SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,POSTINGFLG, ROWSTATE,DOCUMENT_CLASSIFICATION,TOTUPAH,TOTBAHAN,TOTVRA,TOTWIP,VEHICLECODE,HUJAN,CONTRACTORCODE) " & _
						" VALUES (b.BKMNO,b.BKMDATE,b.SECTIONCODE,b.GANGCODE,b.SUPERVISI1,b.SUPERVISI2,b.SUPERVISI3,b.SUPERVISI4,b.CREATEDATE,b.LASTUPDATE,b.FCENTRY,b.FCEDIT,b.FCIP,b.FCBA,b.POSTINGFLG,b.ROWSTATE,b.DOCUMENT_CLASSIFICATION,b.TOTUPAH,b.TOTBAHAN,b.TOTVRA,b.TOTWIP,b.VEHICLECODE,b.HUJAN,b.CONTRACTORCODE)"
		' Response.Write sqlexec
		connStr.execute(sqlexec)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & sqlexec
		ExecuteAndLog modulFile,executeQry
		
		'Insert upahdtl
		sqlexec =   " MERGE INTO BKMUPAHDTL a " & _
					" USING ( " & _
					"   SELECT " & _
					"   'M'||trim(M.BKMNO) BKMNO,M.EMPLOYEECODE,M.ATTENDANCE,M.ACTCODE,M.LOCATIONCODE," & _
					"   M.MANDAYS,M.RATE,M.UNIT,M.OUTPUT,M.ABW," & _
					"   M.BUNCH_WEIGHT,M.KG_BRONDOLAN," & _
					"   M.OVERTIME_HOURS,M.TYPE_OVERTIME," & _
					"   M.CHARGEUNIT,M.CHARGEJOB,M.CHARGECODE," & _
					"   M.ATTENDANCE_FIN,F.TIMEIN IN_FIN,F.TIMEOUT OUT_FIN,M.MANDAYS_FIN," & _
					"   M.ALLOWANCE_CODE,M.ALLOWANCE_AMT," & _
					"   row_number() OVER (PARTITION BY M.bkmno ORDER BY M.employeecode,M.actcode,M.locationcode) LINENOKEY," & _
					"   M.REMARKS," & _
					"   M.CREATEDATE,SYSDATE LASTUPDATE,M.FCENTRY,M.FCEDIT,'" & Request.ServerVariables("REMOTE_ADDR") & "' FCIP," & _
					"   M.FCBA," & _
					"   M.OTHCONVERTION,M.BUNCHUN,M.BUNCHBA" & _
					"   FROM mbkmupahdtl M" & _
					"   INNER JOIN MBKMHDR H ON H.BKMNO=M.BKMNO AND H.FCBA=M.FCBA" & _
					"   LEFT JOIN (" & _
					"       SELECT EMPNO, FCBA, TDATE," & _
					"              MAX(TIMEIN)  AS TIMEIN," & _
					"              MAX(TIMEOUT) AS TIMEOUT" & _
					"       FROM USERFINGER.ATTENDANCE_TAB" & _
					"       GROUP BY EMPNO, FCBA, TDATE" & _
					"   ) F ON (F.EMPNO=M.EMPLOYEECODE AND F.FCBA=M.FCBA AND F.TDATE=H.BKMDATE)" & _
					"   WHERE M.BKMNO = '" & bkmno1 & "'" & _
					"     AND M.FCBA  = '" & businessunit & "'" & _
					" ) b ON (" & _
					"     a.BKMNO        = b.BKMNO" & _
					" AND a.EMPLOYEECODE = b.EMPLOYEECODE" & _
					" AND a.ACTCODE      = b.ACTCODE" & _
					" AND a.LOCATIONCODE = b.LOCATIONCODE" & _
					" AND a.CHARGEUNIT   = b.CHARGEUNIT" & _
					" AND a.FCBA         = b.FCBA" & _
					" AND a.LINENOKEY    = b.LINENOKEY" & _
					" )" & _
					" WHEN MATCHED THEN" & _
					"   UPDATE SET" & _
					"     a.ATTENDANCE     = b.ATTENDANCE," & _
					"     a.MANDAYS        = b.MANDAYS," & _
					"     a.RATE           = b.RATE," & _
					"     a.UNIT           = b.UNIT," & _
					"     a.OUTPUT         = b.OUTPUT," & _
					"     a.ABW            = b.ABW," & _
					"     a.BUNCH_WEIGHT   = b.BUNCH_WEIGHT," & _
					"     a.KG_BRONDOLAN   = b.KG_BRONDOLAN," & _
					"     a.OVERTIME_HOURS = b.OVERTIME_HOURS," & _
					"     a.TYPE_OVERTIME  = b.TYPE_OVERTIME," & _
					"     a.CHARGEJOB      = b.CHARGEJOB," & _
					"     a.CHARGECODE     = b.CHARGECODE," & _
					"     a.ATTENDANCE_FIN = b.ATTENDANCE_FIN," & _
					"     a.IN_FIN         = b.IN_FIN," & _
					"     a.OUT_FIN        = b.OUT_FIN," & _
					"     a.MANDAYS_FIN    = b.MANDAYS_FIN," & _
					"     a.ALLOWANCE_CODE = b.ALLOWANCE_CODE," & _
					"     a.ALLOWANCE_AMT  = b.ALLOWANCE_AMT," & _
					"     a.REMARKS        = b.REMARKS," & _
					"     a.LASTUPDATE     = b.LASTUPDATE," & _
					"     a.FCEDIT         = b.FCEDIT," & _
					"     a.FCIP           = b.FCIP," & _
					"     a.OTHCONVERTION  = b.OTHCONVERTION," & _
					"     a.BUNCHUN        = b.BUNCHUN," & _
					"     a.BUNCHBA        = b.BUNCHBA" & _
					" WHEN NOT MATCHED THEN" & _
					"   INSERT (BKMNO,EMPLOYEECODE,ATTENDANCE,ACTCODE,LOCATIONCODE," & _
					"     MANDAYS,RATE,UNIT,OUTPUT,ABW," & _
					"     BUNCH_WEIGHT,KG_BRONDOLAN," & _
					"     OVERTIME_HOURS,TYPE_OVERTIME," & _
					"     CHARGEUNIT,CHARGEJOB,CHARGECODE," & _
					"     ATTENDANCE_FIN,IN_FIN,OUT_FIN,MANDAYS_FIN," & _
					"     ALLOWANCE_CODE,ALLOWANCE_AMT," & _
					"     LINENOKEY,REMARKS," & _
					"     CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP," & _
					"     FCBA,OTHCONVERTION,BUNCHUN,BUNCHBA)" & _
					"   VALUES (b.BKMNO,b.EMPLOYEECODE,b.ATTENDANCE,b.ACTCODE,b.LOCATIONCODE," & _
					"     b.MANDAYS,b.RATE,b.UNIT,b.OUTPUT,b.ABW," & _
					"     b.BUNCH_WEIGHT,b.KG_BRONDOLAN," & _
					"     b.OVERTIME_HOURS,b.TYPE_OVERTIME," & _
					"     b.CHARGEUNIT,b.CHARGEJOB,b.CHARGECODE," & _
					"     b.ATTENDANCE_FIN,b.IN_FIN,b.OUT_FIN,b.MANDAYS_FIN," & _
					"     b.ALLOWANCE_CODE,b.ALLOWANCE_AMT," & _
					"     b.LINENOKEY,b.REMARKS," & _
					"     b.CREATEDATE,b.LASTUPDATE,b.FCENTRY,b.FCEDIT,b.FCIP," & _
					"     b.FCBA,b.OTHCONVERTION,b.BUNCHUN,b.BUNCHBA)"
					''	Response.Write sqlexec
						connStr.execute(sqlexec)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & sqlexec
						ExecuteAndLog modulFile,executeQry
		'Proses BKM
		sqlexec = "begin bkmproses_upah_api.insert_bkm('M" & bkmno1 & "',TO_DATE('" & bkmdate & "', 'mm/dd/yyyy'),'" & businessunit & "','" & Session("CodeName") & "'); end;"
		connStr.execute(sqlexec)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & sqlexec
		ExecuteAndLog modulFile,executeQry
		
		'SET uploadflg=1
		sqlexec = "UPDATE mbkmhdr SET postingflg=1 WHERE bkmno = '" & bkmno1 & "' AND fcba='" & businessunit & "'"
		connStr.execute(sqlexec)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & sqlexec
		ExecuteAndLog modulFile,executeQry

		sqlexec = "begin bkmjournal_api.Update_bkm(TO_DATE('" & bkmdate & "', 'mm/dd/yyyy'),'" & businessunit & "','" & Session("CodeName") & "',1); end;"
		connStr.execute(sqlexec)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & sqlexec
		ExecuteAndLog modulFile,executeQry

		response.redirect("bkm_oprupload_view.asp?bkmno="&bkmno1&"&tglbkm="&bkmdate&"&fcba="&businessunit)


end if
IF btnApprove = "UN-UPLOAD" then

	sql1 = "UPDATE MBKMHDR SET POSTINGFLG=0 WHERE BKMNO='" & bkmno1 & "'"
	connStr.execute(sql1)
	SQL = "begin bkmproses_upah_api.delete_bkm('" & nombkm & "',TO_DATE('" & bkmdate & "', 'mm/dd/yyyy'),'" & businessunit & "','" & Session("CodeName") & "'); end;"
	connStr.execute(SQL)
	WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL
	response.redirect("bkm_oprupload_view.asp?bkmno="&bkmno1)

end if

IF btnApprove = "UPLOAD BAHAN" then

	bkmno2 = "M" & bkmno1
	'Response.Write bkmno2
	'Response.END

	sql = "DELETE FROM BKMBAHANDTL WHERE BKMNO='" & bkmno2 & "'"
	connStr.execute(sql)
	WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sql
	
	sql1 = "INSERT INTO BKMBAHANDTL(BKMNO, ACTCODE, LOCATIONTYPE,LOCATIONCODE, ITEMCODE,UOM,QTY," & _
	" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, " & _
	" FCIP, FCBA, LINENOKEY) " & _
	" Select 'M'||trim(BKMNO) BKMNO,ACTCODE,'-' LOCATIONTYPE,LOCATIONCODE,ITEMCODE,UOM,QTY, " & _
	" CREATEDATE,LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,LINENOKEY " & _
	" From   MBKMBAHANDTL  " & _
	" WHERE BKMNO='" & bkmno1 & "'"
	' response.write sql1
	connStr.execute(sql1)
	WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sql1

	
	'Response.Write bkmno2

	 SQL = "begin bkmproses_bahan_api.insert_bkm('" & bkmno2 & "',TO_DATE('" & bkmdate & "', 'mm/dd/yyyy'),'" & Session("fcba") & "','" & Session("CodeName") & "'); end;"
	 connStr.execute(SQL)
	 WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & SQL
	
	 sql1 = "UPDATE MBKMBAHANDTL SET ISUPLOAD=1 WHERE BKMNO='" & bkmno1 & "'"
	 connStr.execute(sql1)
	 WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & sql1

	 response.redirect("bkm_oprupload_view.asp?bkmno="&bkmno1)

END IF

IF  btnApprove = "SIMPAN DENDA" THEN

	dim sqldenda,nQtdFieldsdenda,nCountdenda,linenokeydenda,empcodedenda,linedenda
	bkmno2 = "M" & bkmno1
			nQtdFieldsdenda 	= request.form("jumldatadenda")
			nCountdenda 		= 1
			linenokeydenda 		= 1	

			do while cint(nCountdenda) <= cint(nQtdFieldsdenda)

				  
				empcodedenda = request.form("empcode" & nCountdenda)
				linedenda = request.form("linedenda" & nCountdenda)
				
				if empcodedenda<>"" then 

					sqldenda = " insert into BKMDENDA ( BKMNO, EMPLOYEECODE, LINENOKEY, REMARKS, CREATEDATE, " & _
					" LASTUPDATE, FCENTRY, FCEDIT, FCIP, FCBA, " & _
					" TBSA, TBSBA, TBSTP, TBSBK, TBSO, " & _
					" TBSAB, TBSE, ANCAKBT, ANCAKPS, ANCAKSP, " & _
					" ANCAKBRD, ANCAKOP, ANCAKBM, ANCAKBDT) " & _
					" values ('" & bkmno2 & "','" & empcodedenda & "','" & linenokeydenda & "','Denda Panen', " & _ 
					" TO_DATE('" & bkmdate & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & businessunit & "', " & _
					" '" & request.form("ktbsa" & nCountdenda) & "','" & request.form("ktbsba" & nCountdenda) & "', " & _ 
					" '" & request.form("ktbstp" & nCountdenda) & "','" & request.form("ktbsbk" & nCountdenda) & "', " & _ 
					" '" & request.form("ktbso" & nCountdenda) & "','" & request.form("ktbsab" & nCountdenda) & "', " & _ 
					" '" & request.form("ktbse" & nCountdenda) & "','" & request.form("kancakbt" & nCountdenda) & "', " & _ 
					" '" & request.form("kancakps" & nCountdenda) & "','" & request.form("kancaksp" & nCountdenda) & "', "  & _ 
					" '" & request.form("kancakbrd" & nCountdenda) & "','" & request.form("kancakop" & nCountdenda) & "', "  & _ 
					" '" & request.form("kancakbm" & nCountdenda) & "','" & request.form("kancakbdt" & nCountdenda) & "')"
					'Response.Write sqldenda
					connStr.execute(sqldenda)

				end if

 			nCountdenda 	= nCountdenda + 1
			linenokeydenda 	= linenokeydenda + 1

			loop
			response.redirect("bkm_oprupload_view.asp?bkmno="&bkmno1&"&tglbkm="&bkmdate&"&fcba="&businessunit)

END IF

'count detaildenda
Dim SQL_Rowsdenda, NoOfRowsdenda,RSCountdenda

SQL_Rowsdenda = "SELECT Count (*) from BKMDENDA where BKMNO='" & bkmno2 & "'"
set RSCountdenda = connStr.Execute(SQL_Rowsdenda)
if not RSCountdenda.eof then 
	NoOfRowsdenda = RSCountdenda(0).value
end if
RSCountdenda.close


%>

<html>
	<head>
		<title>FBPlants - View BKM Operasional</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261015a"></script>
<% end if %>
</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: -15px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 48px;
	}
	.lebarlokasi{
		width: 60px;
	}
	.lebartc{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		width: 110px;
	}
	.lebarallow{
		width: 90px;
	}
	.lebarupah{
		width: 80px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
	}
	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	.lebarovtime{
		width: 70px;
	}
	.lebardenda{
		width: 60px;
	}
	.warnaamount{
		background-color: #eee;
	}
	
</style>



<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-bkm"<% end if %>>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM OPERASIONAL VIEW</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_oprupload_view.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="<%=NoOfRowsitem%>" />
	<input type="hidden" name="jumldatadenda" id="jumldatadenda" value="<%=NoOfRowsdenda%>" />
	<input type="hidden" name="proses" id="proses" value="2" />
	<input type="hidden" name="bkmno1" id="bkmno1" value="<%=bkmno%>" />
	<input type="hidden" name="bkmdate" id="bkmdate" value="<%=DATECREATED%>" />
	<table class="data posisitengah" border="0" width="90%">			
		<tr>
			<td align="left" >BKM No:</td>
			<td align="left">
				<%=bkmno%>
			</td>


			<td align="left" >Section Name:</td>
			<td align="left">
				<%=SECTIONNAME%>
			</td>
			<td align="left" >Mandor I</td>
			<td align="left">
				<select name="supervisi1" id="supervisi1" class="lebarsection" >
					<option value="<%=supervisi1%>"><%=SUPERVISINAME1%></option>
				</select>
			</td>
			<td align="left" >Kerani Panen</td>
			<td align="left">
				<select name="supervisi3" id="supervisi3" class="lebarsection" >
					<option value="<%=SUPERVISI3%>"><%=SUPERVISINAME3%></option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Tanggal:</td>
			<td align="left">
				<%=DATECREATED%>
			</td>
			
			<td align="left" >Gang:</td>
			<td align="left">
				<%=GANGCODE%>
			</td>
			<td align="left" >Mandor</td>
			<td align="left">
				<select name="supervisi2" id="supervisi2" class="lebarsection" >
					<option value="<%=SUPERVISI2%>"><%=SUPERVISI2%></option>
				</select>
			</td>
			<td align="left" >Kerani Transport</td>
			<td align="left">
				<select name="supervisi4" id="supervisi4" class="lebarsection" >
					<option value="<%=SUPERVISI4%>"><%=SUPERVISINAME4%></option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Posting:</td>
			<td align="left"><%=POSTINGFLG%>-<%=ROWSTATE%></td>
		
			<td align="left" >Hujan</td>
			<td align="left">
				<% if HUJANCK <> "0" then %>
					<input type="checkbox" name="hujan" id="hujan" class="" checked>
				<% else %>
					<input type="checkbox" name="hujan" id="hujan" class="">
				<% end if %>
			</td>
		
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id='tableall'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM BAHAN</li>
						<li class="TabbedPanelsTab">BKM DENDA</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm upah-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="600" class='data' id='table2'>
								<tr id="trheader2">
									<th>Employee</th>
									<th>TimeIn</th>
									<th>TimeOut</th>
									<th>Lokasi</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>Uom</th>
									<th>TC</th>
									<th>HK</th>
									<th>Unit</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
									<th>OVTH</th>
									<th>TOVTime</th>
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,ttlhk,ttlunit,ttloutput,ttlrate,ttlrateall,ttlovth
								sqldtlupah = "SELECT u.BKMNO,u.fcba, u.chargeunit, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
								" u.ALLOWANCE_CODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
								" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
								" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
								" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)UNIT,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(u.OTHCONVERTION,0)OTHCONVERTION ,a.uom_unit" & _
								" FROM MBKMUPAHDTL u " & _ 
								" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
								" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE) " & _
								" LEFT JOIN ALLOWDED al ON (al.FCCODE = u.ALLOWANCE_CODE and al.fcba=u.fcba) " & _ 
								" LEFT JOIN overtime_convertion o ON (o.type=u.TYPE_OVERTIME and o.HOURS=u.OVERTIME_HOURS and o.fcba=u.fcba)" & _
								" where u.BKMNO='" & bkmno & "' order by e.fcname, a.fcname,u.LOCATIONCODE,u.ALLOWANCE_CODE"
								set rsdtlupah = connStr.Execute(sqldtlupah)
								'response.write sqldtlupah
								

																
								no = 0
								ttlhk = 0
								ttlunit = 0
								ttloutput = 0
								ttlrate = 0
								ttlrateall = 0
								ttlovth = 0
								totalamount = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("fcname")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("IN_FINDEPAN")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OUT_FINDEPAN")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("chargeunit")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("ATTENDANCE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("activityname")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("uom_unit")%>
									</td>
									
									<td class="ratatengah">
										<%=rsdtlupah("LOCATIONCODE")%>
									</td>
									<td>
										<%=rsdtlupah("MANDAYS")%>
									</td>
									
									<td>
										<%=rsdtlupah("UNIT")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("OUTPUT")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("rate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("ALLOWDEDNAME")%>
									</td>
																		
									<td class="ratakiri">
										<%=FormatNumber(rsdtlupah("ALLOWANCE_AMT"),0)%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OVERTIME_HOURS")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("TYPE_OVERTIME")%>&nbsp;<%=rsdtlupah("type_ovname")%>
									</td>
									
									
									<td></td>
									
								</tr>
								<%
									'dim ttlhk
									no = no +1
									ttlhk = cdbl(ttlhk) + cdbl(rsdtlupah("MANDAYS"))
									ttlunit = cdbl(ttlunit) + cdbl(rsdtlupah("unit"))
									ttloutput = cdbl(ttloutput) + cdbl(rsdtlupah("output"))
									ttlrate = cdbl(ttlrate) + cdbl(rsdtlupah("rate"))
									ttlrateall = cdbl(ttlrateall) + cdbl(rsdtlupah("ALLOWANCE_AMT"))
									ttlovth = cdbl(ttlovth) + cdbl(rsdtlupah("OVERTIME_HOURS"))
									
									
									rsdtlupah.MoveNext
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="8" class="tebal">Total</td>
										<td class="tebal"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlhk,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlunit,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttloutput,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarrate" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlrate,2)%>"/></td>
										<td></td>
										<td class="tebal ratakiri"><input class="ratakanan lebarrate" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlrateall,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlovth,0)%>"/></td>
										<td></td>
										
									</tr>
									<tr>
										<td colspan="16" align="center" >
										
											<input type=Submit name="btnsave" id="btnsave" value="UPLOAD">&nbsp;&nbsp;
											<input type=Submit name="btnsave" id="btnsave" value="UN-UPLOAD">&nbsp;&nbsp;
											<input type=button onclick="window.close();" name="btnclose" value="CLOSE">
										</td>
									</tr>
							</table>
						</div>

						<!-- view bkm bahan-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="1" width="850" class='data' id='table1'>
									<tr id="trheader1">
										<th>Activity</th>
										<th>TC</th>
										<th>Item Barang</th>
										<th>Satuan</th>
										<th>Qty</th>
										<th>Status Upload</th>
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										dim sqlbarang,rsbarang 
										sqlbarang = "select b.BKMNO,b.ACTCODE,b.ACTCODE || '-' ||  a.fcname as activityname,b.ITEMCODE,s.FCCODE || '-' || s.FCNAME ITEMBARANG, " & _
										" s.UNITOFMEASUREMENT,b.LOCATIONCODE,NVL(b.qty,0)qty,b.isupload " & _
										" from MBKMBAHANDTL b " & _
										" inner join STOCK s on (b.ITEMCODE=s.FCCODE AND b.fcba=s.fcba) " & _
										" inner join jobmaster a on (b.ACTCODE=a.FCCODE) where b.BKMNO='" & bkmno & "' order by b.ITEMCODE asc"
										 set rsbarang = connStr.Execute(sqlbarang)
										 'response.write sqlbarang
									%>

									<% 
										dim nobr,isupload
										nobr = 0
										While Not rsbarang.EOF 
											isupload = rsbarang("isupload")
											if isupload = "1" then
												isupload ="YA"
											else 
												isupload ="TIDAK"
											END IF
									%>
										<tr id='trb<%=nobr%>'>
											
											<td class="ratakiri">
												<%=rsbarang("activityname")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("LOCATIONCODE")%>
											</td>
											<td class="ratakiri">
												<%=rsbarang("ITEMCODE")%>&nbsp;<%=rsbarang("ITEMBARANG")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("UNITOFMEASUREMENT")%>
											</td>
											<td class="ratatengah">
												<%=rsbarang("qty")%>
											</td>
											<td class="ratatengah">
												<%=isupload%>
											</td>
											<td></td>
										</tr>

									<%
										nobr = nobr + 1
										rsbarang.MoveNext
										Wend
									
									%>
										<tr>
											<td colspan="6" align="center">
												<input type=Submit name="btnsave" id="btnsave" value="UPLOAD BAHAN">&nbsp;&nbsp;
											</td>
										</tr>
									
								</table>
							</div>
							
						</div>

						<!-- view bkm denda-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="detaildenda">
								<table align=center border="0" width="950" class='data' id='table3'>
									<tr >
										
											<th rowspan="2">Employee</th>
											<th colspan="7">Kualitas TBS</th>
											<th colspan="7">Kualitas Ancak</th>
											<th rowspan="2" colspan="2">&nbsp;</th>
									</tr>
									<tr id="trheader3">
										
										<th>(A)</th>
										<th>(BA)</th>
										<th>(TP)</th>
										<th>(BK)</th>
										<th>(O)</th>
										<th>(AB)</th>
										<th>(E)</th>

										<th>(BT)</th>
										<th>(PS)</th>
										<th>(SP)</th>
										<th>(BRD)</th>
										<th>(OP)</th>
										<th>(BM)</th>
										<th>(BDT)</th>
									</tr>

									<%
									
										dim rsdenda,nodenda 
										sqldenda = "Select B.EMPLOYEECODE,E.FCNAME, B.LINENOKEY, B.REMARKS, " & _
										" B.TBSA, B.TBSBA, B.TBSTP, B.TBSBK, B.TBSO, " & _
										" B.TBSAB, B.TBSE, B.ANCAKBT, B.ANCAKPS, B.ANCAKSP, " & _
										" B.ANCAKBRD, B.ANCAKOP, B.ANCAKBM, B.ANCAKBDT " & _
										" From   BKMDENDA B " & _
										" LEFT JOIN EMPLOYEE E ON (B.EMPLOYEECODE=E.FCCODE AND B.FCBA=E.FCBA) " & _
										" WHERE B.BKMNO='" & bkmno & "' order by e.fcname asc"
										 set rsdenda = connStr.Execute(sqldenda)
										 'response.write sqlbarang
									%>

									<% 
										 
										nodenda = 0
										While Not rsdenda.EOF 
									%>
										<tr id='trd<%=nodenda%>'>
											<td class="ratakiri">
												<select class="lebaractivity " name="empcode<%=nodenda%>" id="empcode<%=nodenda%>" onchange="getLineDenda('<%=nodenda%>')">
													<option value=<%=rsdenda("EMPLOYEECODE")%>><%=rsdenda("FCNAME")%></option>
												</select><input type="hidden" name="linedenda<%=nodenda%>" id="linedenda<%=nodenda%>" value="<%=rsdenda("LINENOKEY")%>"></input>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsa<%=nodenda%>" id="ktbsa<%=nodenda%>" value="<%=rsdenda("TBSA")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsba<%=nodenda%>" id="ktbsba<%=nodenda%>" value="<%=rsdenda("TBSBA")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbstp<%=nodenda%>" id="ktbstp<%=nodenda%>" value="<%=rsdenda("TBSTP")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsbk<%=nodenda%>" id="ktbsbk<%=nodenda%>" value="<%=rsdenda("TBSBK")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbso<%=nodenda%>" id="ktbso<%=nodenda%>" value="<%=rsdenda("TBSO")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbsab<%=nodenda%>" id="ktbsab<%=nodenda%>" value="<%=rsdenda("TBSAB")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="ktbse<%=nodenda%>" id="ktbse<%=nodenda%>" value="<%=rsdenda("TBSE")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbt<%=nodenda%>" id="kancakbt<%=nodenda%>" value="<%=rsdenda("ANCAKBT")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakps<%=nodenda%>" id="kancakps<%=nodenda%>" value="<%=rsdenda("ANCAKPS")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancaksp<%=nodenda%>" id="kancaksp<%=nodenda%>" value="<%=rsdenda("ANCAKSP")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbrd<%=nodenda%>" id="kancakbrd<%=nodenda%>" value="<%=rsdenda("ANCAKBRD")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakop<%=nodenda%>" id="kancakop<%=nodenda%>" value="<%=rsdenda("ANCAKOP")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbm<%=nodenda%>" id="kancakbm<%=nodenda%>" value="<%=rsdenda("ANCAKBM")%>"/>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebardenda" type="text" name="kancakbdt<%=nodenda%>" id="kancakbdt<%=nodenda%>" value="<%=rsdenda("ANCAKBDT")%>"/>
											</td>
											

											<td>
												<img src='css/login/images/add.jpg' onclick="addrowDenda('<%=nodenda%>')" />
											</td>
											<td>
												<% if nodenda >= 0 then %>
													<img src='css/login/images/del.png' onclick="deleterowDenda('<%=nodenda%>')" />
												<% else 
													response.write "&nbsp;"
												end if %>
											</td>
											<td></td>
										</tr>

											
										
									<%
										nodenda = nodenda + 1
										rsdenda.MoveNext
										Wend
									
									%>

										<tr>
											<td colspan="16" align="center">
												<input type=Submit name="btnsave" id="btnsave" value="SIMPAN DENDA">&nbsp;&nbsp;
											</td>
										</tr>
									
								</table>
							</div> 
							<!-- ============================================================================================================================= -->
						</div>

					</div>
				</div>
			</td>
		</tr>
		
	</table>
	<br>
	

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	
	//$("#tglbkm").datepicker();
	showSupervisi1();
	showSupervisi2();
	showSupervisi3();
	showSupervisi4();

	var jumldatadendax = $('#jumldatadenda').val();

	if (jumldatadendax=="0"){
		addrowDenda(1);
		
	}
	
});


//untuk input denda panen
var jumldenda = $('#jumldatadenda').val();
if (jumldenda=="0"){
		var StartIterationdenda = 1
}else{
	var StartIterationdenda =  parseInt(jumldenda) + 1
}
function addrowDenda(trid)
{

	var lsdenda = this.StartIterationdenda++;
	
	if  (lsdenda < 2) {
		var row = $("#trheader3").after("<tr id='trd" + lsdenda + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowDenda(" + lsdenda + ");' />";
	} else {
		var row = $('#trd' + trid).after("<tr id='trd" + lsdenda + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowDenda(" + lsdenda + ");' />";
	}

	var empcode = "<select name='empcode" + lsdenda + "' id='empcode" + lsdenda + "' onchange='getLineDenda(" + lsdenda + ")' class='lebaractivity ratatengah'></select><input type='hidden' name='linedenda" + lsdenda + "' id='linedenda" + lsdenda + "'></input>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + empcode + "</td>");

	var ktbsa = "<input type='text' name='ktbsa" + lsdenda + "' id='ktbsa" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsa + "</td>");

	var ktbsba = "<input type='text' name='ktbsba" + lsdenda + "' id='ktbsba" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsba + "</td>");

	var ktbstp = "<input type='text' name='ktbstp" + lsdenda + "' id='ktbstp" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbstp + "</td>");

	var ktbsbk = "<input type='text' name='ktbsbk" + lsdenda + "' id='ktbsbk" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsbk + "</td>");

	var ktbso = "<input type='text' name='ktbso" + lsdenda + "' id='ktbso" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbso + "</td>");

	var ktbsab = "<input type='text' name='ktbsab" + lsdenda + "' id='ktbsab" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsab + "</td>");

	var ktbse = "<input type='text' name='ktbse" + lsdenda + "' id='ktbse" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbse + "</td>");

	var kancakbt = "<input type='text' name='kancakbt" + lsdenda + "' id='kancakbt" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbt + "</td>");

	var kancakps = "<input type='text' name='kancakps" + lsdenda + "' id='kancakps" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakps + "</td>");

	var kancaksp = "<input type='text' name='kancaksp" + lsdenda + "' id='kancaksp" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancaksp + "</td>");

	var kancakbrd = "<input type='text' name='kancakbrd" + lsdenda + "' id='kancakbrd" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbrd + "</td>");

	var kancakop = "<input type='text' name='kancakop" + lsdenda + "' id='kancakop" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakop + "</td>");

	var kancakbm = "<input type='text' name='kancakbm" + lsdenda + "' id='kancakbm" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbm + "</td>");

	var kancakbdt = "<input type='text' name='kancakbdt" + lsdenda + "' id='kancakbdt" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebardenda ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbdt + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowDenda(" + lsdenda + ");' />";
	$('#trd' + lsdenda).append("<td>" + tambahbutton + "</td>");

	$('#trd' + lsdenda).append("<td>" + kurangbutton + "</td>");

	$('#jumldatadenda').val(lsdenda);

	showEmployeedenda("empcode" + lsdenda);
	
	
}

//hapus row bahan
function deleterowDenda(i) {
	Spry.Widget.Utils.destroyWidgets('empcode' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsa' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsba' + i);
	Spry.Widget.Utils.destroyWidgets('ktbstp' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsbk' + i);
	Spry.Widget.Utils.destroyWidgets('ktbso' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsab' + i);
	Spry.Widget.Utils.destroyWidgets('ktbse' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbt' + i);
	Spry.Widget.Utils.destroyWidgets('kancakps' + i);
	Spry.Widget.Utils.destroyWidgets('kancaksp' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbrd' + i);
	Spry.Widget.Utils.destroyWidgets('kancakop' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbm' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbdt' + i);
	 $('#trd' + i).remove();
	
}

function showEmployeedenda(selid) {
	var tglbkm = $('#bkmdate').val();
	var bkmno = $('#bkmno1').val();
	//$('#fccode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp?emp=" + 6 + "&tglbkm=" + tglbkm + "&bkmno=" +bkmno,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//view data oth convertion
function getLineDenda(i)
{	
	var bkmno = $("#bkmno").val();
	var empcode = $("#empcode" + i).val();
	if (empcode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 5 + "&empcode=" +  empcode + "&bkmno=" + bkmno, function(data) {
			$('#linedenda' + i).empty();
			var json  = JSON.parse(data);
			var linedenda = json.results[0].linedenda
			$('#linedenda' + i).val(linedenda);
						
		});
		
	}

	
}

//get supervisi 1
function showSupervisi1() {
	
	$('#supervisi1').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 2
function showSupervisi2() {
	
	$('#supervisi2').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 3
function showSupervisi3() {
	
	$('#supervisi3').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 1
function showSupervisi4() {
	
	$('#supervisi4').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

</script>
