<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
maxCol = 0

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")

txttdate = Request("txttdate")

' If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
' If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))

If txtMonth = "" OR txtYear = ""  Then
    txtMonth = Month(Now())
    txtYear = Year(Now())
End If
If "" & txttdate = "" Then txttdate = textDate(date)

if (txtMonth < 10) then txtMonth = "0" & txtMonth else txtMonth = Month(Session("v_Costbook"))

txttdate = txtMonth & "" & txtYear

add_col_bud_prod_qty = ""
d_span = 26
	
v_txtFcba = Request("txtFcba")
v_txtGangCode = Request("txtGangCode")
v_txtDivision = Request("txtDivision")
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

				If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
					v_txtFcba = varDefaultBa
				End If	
  
				if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
			       whereCons = whereCons & " And fcba = '" & v_txtfcba & "'"
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

				' if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			 '      if v_txtDivision <>  "--ALL--"  then
				' 	whereCons = whereCons & " AND SECTIONNAME = '" & v_txtDivision & "'"
			 '      end if
    '               selectCons = selectCons &  " , SECTIONNAME "
				'   orderCons  = orderCons & ",SECTIONNAME "
			 '   else
			 '   nSpanTotal = nSpanTotal - 1  
			 '   end if

			 '     if  v_txtGangCode <> "--ALL--"  and v_txtGangCode <> "" then
				'        whereCons = whereCons & " And fccode = '" & v_txtGangCode & "' "
				' 	end if
				'     ' selectCons = selectCons & ",fcba"
				'     ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
				' 	q = "Select fccode from gang where fcCode='" & v_txtGangCode & "'"
				' 	set rsCat = connStr.Execute(q)
				' 	if not rsCat.eof then
				' 	   fcCode = rsCat("fcCode")
				' else
				'     nSpanTotal = nSpanTotal - 1 
				' end if

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_field_eof = true

' status_payroll = "FALSE"
' strKJ = " SELECT NVL(payroll_executed, 'FALSE') detail" & _
' 		" FROM empalldedstatus" & _
' 		" WHERE fcba = '" & v_txtFcba & "'" & _
' 		"       AND year = " & txtYear & "" & _
' 		"       AND lpad(month,2,0) = '" & txtMonth & "'"

' Set rsKJ = Server.CreateObject("ADODB.Recordset")
' rsKJ.Open strKJ, connstr, adOpenStatic

' IF NOT rsKJ.EOF THEN
' 	status_payroll = rsKJ("detail")
' ELSE
' 	status_payroll = "FALSE"
' END IF
' rsKJ.Close
' Set rsKJ = Nothing
' ' response.write strKJ
' IF status_payroll = "TRUE" THEN
strPusingan	= 	"         SELECT fcba, sectionname," & _
				"         SUM (Jumlah_TK_PNN_AWAL) Jumlah_TK_PNN_AWAL," & _
				"         SUM (Jumlah_TK_PNN_MASUK) Jumlah_TK_PNN_MASUK," & _
				"         SUM (Jumlah_TK_PNN_KELUAR) Jumlah_TK_PNN_KELUAR," & _
				"         SUM (Jumlah_TK_PML_AWAL) Jumlah_TK_PML_AWAL," & _
				"         SUM (Jumlah_TK_PML_MASUK) Jumlah_TK_PML_MASUK," & _
				"         SUM (Jumlah_TK_PML_KELUAR) Jumlah_TK_PML_KELUAR," & _
				"         SUM (Jumlah_TK_MD_AWAL) Jumlah_TK_MD_AWAL," & _
				"         SUM (Jumlah_TK_MD_MASUK) Jumlah_TK_MD_MASUK," & _
				"         SUM (Jumlah_TK_MD_KELUAR) Jumlah_TK_MD_KELUAR," & _
				"         SUM (Jumlah_TK_TR_AWAL) Jumlah_TK_TR_AWAL," & _
				"         SUM (Jumlah_TK_TR_MASUK) Jumlah_TK_TR_MASUK," & _
				"         SUM (Jumlah_TK_TR_KELUAR) Jumlah_TK_TR_KELUAR," & _
				"         SUM (Jumlah_TK_KU_AWAL) Jumlah_TK_KU_AWAL," & _
				"         SUM (Jumlah_TK_KU_MASUK) Jumlah_TK_KU_MASUK," & _
				"         SUM (Jumlah_TK_KU_KELUAR) Jumlah_TK_KU_KELUAR" & _
				"    FROM (  SELECT fcba," & _
				"                   sectionname," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('PN')" & _
				"					AND dateterminate is null" & _
				"                   AND dateofjoin < TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('MD')" & _
				"					AND dateterminate is null" & _
				"                   AND dateofjoin < TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
				"					AND dateterminate is null" & _
				"                   AND dateofjoin < TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
				"					AND dateterminate is null" & _
				"                   AND dateofjoin < TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) NOT IN ('PN'," & _
				"                                                       'KT'," & _
				"                                                       'KM'," & _
				"                                                       'UM'," & _
				"                                                       'OS'," & _
				"                                                       'MT'," & _
				"                                                       'TB'," & _
				"                                                       'MD')" & _
				"					AND dateterminate is null" & _
				"                   AND dateofjoin < TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('PN')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('MD')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) NOT IN ('PN'," & _
				"                                                       'KT'," & _
				"                                                       'KM'," & _
				"                                                       'UM'," & _
				"                                                       'OS'," & _
				"                                                       'MT'," & _
				"                                                       'TB'," & _
				"                                                       'MD')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('PN')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('MD')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   0 Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname" & _
				"          UNION ALL" & _
				"            SELECT fcba," & _
				"                   sectionname," & _
				"                   0 Jumlah_TK_PNN_AWAL," & _
				"                   0 Jumlah_TK_PNN_MASUK," & _
				"                   0 Jumlah_TK_PNN_KELUAR," & _
				"                   0 Jumlah_TK_MD_AWAL," & _
				"                   0 Jumlah_TK_MD_MASUK," & _
				"                   0 Jumlah_TK_MD_KELUAR," & _
				"                   0 Jumlah_TK_TR_AWAL," & _
				"                   0 Jumlah_TK_TR_MASUK," & _
				"                   0 Jumlah_TK_TR_KELUAR," & _
				"                   0 Jumlah_TK_KU_AWAL," & _
				"                   0 Jumlah_TK_KU_MASUK," & _
				"                   0 Jumlah_TK_KU_KELUAR," & _
				"                   0 Jumlah_TK_PML_AWAL," & _
				"                   0 Jumlah_TK_PML_MASUK," & _
				"                   COUNT (presentgrade) AS Jumlah_TK_PML_KELUAR" & _
				"              FROM employee" & _
				"             WHERE     SUBSTR (gangcode, 1, 2) NOT IN ('PN'," & _
				"                                                       'KT'," & _
				"                                                       'KM'," & _
				"                                                       'UM'," & _
				"                                                       'OS'," & _
				"                                                       'MT'," & _
				"                                                       'TB'," & _
				"                                                       'MD')" & _
				"                   AND dateofjoin = TO_DATE ('"& txttdate &"', 'MMYYYY')" & _
				"					"& whereCons &""& _
				"          GROUP BY fcba, sectionname)" & _
				" GROUP BY fcba, sectionname ORDER BY sectionname" 

	response.write strPusingan
	Set qORS = Server.CreateObject("ADODB.Recordset")
	qORS.Open strPusingan, connstr, adOpenStatic

	If Not qORS.EOF Then
		arr_field = qORS.getrows()
		arr_field_eof = false
	End If
	qORS.Close
	set qORS = Nothing
	' "  SELECT fcba," & _
	' "         sectionname," & _
	' "         SUM (" & _
	' "              Jumlah_TK_PNN_MASUK" & _
	' "            + Jumlah_TK_PML_MASUK" & _
	' "            + Jumlah_TK_MD_MASUK" & _
	' "            + Jumlah_TK_TR_MASUK" & _
	' "            + Jumlah_TK_KU_MASUK)" & _
	' "            AWAL," & _
	' "         SUM (Jumlah_TK_PNN_MASUK) Jumlah_TK_PNN_MASUK," & _
	' "         SUM (Jumlah_TK_PML_MASUK) Jumlah_TK_PML_MASUK," & _
	' "         SUM (Jumlah_TK_MD_MASUK) Jumlah_TK_MD_MASUK," & _
	' "         SUM (Jumlah_TK_TR_MASUK) Jumlah_TK_TR_MASUK," & _
	' "         SUM (Jumlah_TK_KU_MASUK) Jumlah_TK_KU_MASUK," & _
	' "         SUM (Jumlah_TK_PNN_KELUAR) Jumlah_TK_PNN_KELUAR," & _
	' "         SUM (Jumlah_TK_PML_KELUAR) Jumlah_TK_PML_KELUAR," & _
	' "         SUM (Jumlah_TK_MD_KELUAR) Jumlah_TK_MD_KELUAR," & _
	' "         SUM (Jumlah_TK_TR_KELUAR) Jumlah_TK_TR_KELUAR," & _
	' "         SUM (Jumlah_TK_KU_KELUAR) Jumlah_TK_KU_KELUAR," & _
	' "         SUM (" & _
	' "              Jumlah_TK_PNN_KELUAR" & _
	' "            + Jumlah_TK_PML_KELUAR" & _
	' "            + Jumlah_TK_MD_KELUAR" & _
	' "            + Jumlah_TK_TR_KELUAR" & _
	' "            + Jumlah_TK_KU_KELUAR)" & _
	' "            AKHIR" & _
	' "    FROM (  SELECT fcba," & _
	' "                   sectionname," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('PN')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('MD')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_PML_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) NOT IN ('PN'," & _
	' "                                                       'KT'," & _
	' "                                                       'KM'," & _
	' "                                                       'UM'," & _
	' "                                                       'OS'," & _
	' "                                                       'MT'," & _
	' "                                                       'TB'," & _
	' "                                                       'MD')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateofjoin, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('PN')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('MD')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   0 Jumlah_TK_PML_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname" & _
	' "          UNION ALL" & _
	' "            SELECT fcba," & _
	' "                   sectionname," & _
	' "                   0 Jumlah_TK_PNN_KELUAR," & _
	' "                   0 Jumlah_TK_PNN_MASUK," & _
	' "                   0 Jumlah_TK_MD_KELUAR," & _
	' "                   0 Jumlah_TK_MD_MASUK," & _
	' "                   0 Jumlah_TK_TR_KELUAR," & _
	' "                   0 Jumlah_TK_TR_MASUK," & _
	' "                   0 Jumlah_TK_KU_KELUAR," & _
	' "                   0 Jumlah_TK_KU_MASUK," & _
	' "                   COUNT (presentgrade) AS Jumlah_TK_PML_KELUAR," & _
	' "                   0 Jumlah_TK_PML_MASUK" & _
	' "              FROM employee_hist_tab" & _
	' "             WHERE     SUBSTR (gangcode, 1, 2) NOT IN ('PN'," & _
	' "                                                       'KT'," & _
	' "                                                       'KM'," & _
	' "                                                       'UM'," & _
	' "                                                       'OS'," & _
	' "                                                       'MT'," & _
	' "                                                       'TB'," & _
	' "                                                       'MD')" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'MM')) < '"& txtMonth &"'" & _
	' "                   AND TO_NUMBER (TO_CHAR (dateterminate, 'YYYY')) = '"& txtYear &"'" & _
	' "					"& whereCons &""& _
	' "          GROUP BY fcba, sectionname)" & _
	' " GROUP BY fcba, sectionname" 
' else
' 	strPusingan	= 	" select sectionname, fcba" & _
' 					" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_PNN,0)) as JMLKHT_TK_PNN" & _
' 					" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_MD,0)) as JMLKHT_TK_MD" & _
' 					" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_PML,0)) as JMLKHT_TK_PML" & _
' 					" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_TR,0)) as JMLKHT_TK_TR" & _
' 					" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_KU,0)) as JMLKHT_TK_KU" & _
' 					" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_PNN,0)) as JMLKHL_TK_PNN" & _
' 					" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_MD,0)) as JMLKHL_TK_MD" & _
' 					" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_PML,0)) as JMLKHL_TK_PML" & _
' 					" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_TR,0)) as JMLKHL_TK_TR" & _
' 					" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_KU,0)) as JMLKHL_TK_KU" & _
' 					" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_PNN,0)) as JMLKBT_TK_PNN" & _
' 					" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_MD,0)) as JMLKBT_TK_MD" & _
' 					" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_PML,0)) as JMLKBT_TK_PML" & _
' 					" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_TR,0)) as JMLKBT_TK_TR" & _
' 					" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_KU,0)) as JMLKBT_TK_KU" & _
' 					" from (" & _
' 					" SELECT fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade," & _
' 					"       COUNT (presentgrade) AS Jumlah_TK_PNN," & _
' 					"       0 Jumlah_TK_MD," & _
' 					"       0 Jumlah_TK_TR," & _
' 					"       0 Jumlah_TK_KU," & _
' 					"       0 Jumlah_TK_PML" & _
' 					"  FROM employee" & _
' 					" WHERE  dateterminate IS NULL" & _
' 					"       AND SUBSTR (gangcode, 1, 2) IN ('PN')" & _
' 					"     "& whereCons & "" & _
' 					" GROUP BY fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade" & _
' 					" UNION ALL" & _
' 					" SELECT fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade," & _
' 					"      0 Jumlah_TK_PNN," & _
' 					"       COUNT (presentgrade) AS Jumlah_TK_MD," & _
' 					"       0 Jumlah_TK_TR," & _
' 					"       0 Jumlah_TK_KU," & _
' 					"       0 Jumlah_TK_PML" & _
' 					"  FROM employee" & _
' 					" WHERE dateterminate IS NULL" & _
' 					"       AND SUBSTR (gangcode, 1, 2) IN ('MD')" & _
' 					"     "& whereCons & "" & _
' 					" GROUP BY fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade" & _
' 					" UNION ALL" & _
' 					" SELECT fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade," & _
' 					"       0 Jumlah_TK_PNN," & _
' 					"       0 Jumlah_TK_MD," & _
' 					"       COUNT (presentgrade) AS Jumlah_TK_TR," & _
' 					"       0 Jumlah_TK_KU," & _
' 					"       0 Jumlah_TK_PML" & _
' 					"  FROM employee" & _
' 					" WHERE dateterminate IS NULL" & _
' 					"       AND SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
' 					"     "& whereCons & "" & _
' 					" GROUP BY fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade" & _
' 					" UNION ALL" & _
' 					" SELECT fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade," & _
' 					"       0 Jumlah_TK_PNN," & _
' 					"       0 Jumlah_TK_MD," & _
' 					"       0 Jumlah_TK_TR," & _
' 					"       COUNT (presentgrade) AS Jumlah_TK_KU," & _
' 					"       0 Jumlah_TK_PML" & _
' 					"  FROM employee" & _
' 					" WHERE dateterminate IS NULL" & _
' 					"       AND SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
' 					"     "& whereCons & "" & _
' 					" GROUP BY fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade" & _
' 					" UNION ALL" & _
' 					" SELECT fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade," & _
' 					"       0 Jumlah_TK_PNN," & _
' 					"       0 Jumlah_TK_MD," & _
' 					"       0 Jumlah_TK_TR," & _
' 					"       0 Jumlah_TK_KU," & _
' 					"       COUNT (presentgrade) AS Jumlah_TK_PML" & _
' 					"  FROM employee" & _
' 					" WHERE dateterminate IS NULL" & _
' 					"       AND SUBSTR (gangcode, 1, 2) NOT IN ('PN','KT','KM','UM','OS','MT','TB','MD')" & _
' 					"     "& whereCons & "" & _
' 					" GROUP BY fcba," & _
' 					"       sectionname," & _
' 					"       presentgrade )emp" & _
' 					"       GROUP BY sectionname, fcba" & _
' 					"       order by sectionname, fcba" 


' 	response.write strPusingan
' 	Set qORS = Server.CreateObject("ADODB.Recordset")
' 	qORS.Open strPusingan, connstr, adOpenStatic

' 	If Not qORS.EOF Then
' 		arr_field = qORS.getrows()
' 		arr_field_eof = false
' 	End If
' 	qORS.Close
' 	set qORS = Nothing

' end if

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"
' response.write strBaInfo
Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>Laporan Turn Over Tenaga Kerja - iPlaS Reporting</title>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
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
							
							<th>TAHUN</th>
							<th>BULAN</th>
							<th>KEBUN</th>
							<th>&nbsp;</th>
						</tr>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM

							Yr = txtYear
							MinYr = 2015
							MaxYr = Yr  + 2
						%>
								<td width="4">
									<select name="txtYear">
						<%
							For i = MinYr to MaxYr
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
						<%
							Next
						%>
									</select>
								</td>
								
								<td width="2">
									<select name="txtMonth">
						<% 
							For i = 1 to 12
								
								' if i <= 9 then 
								' 	iBulan = 0&i
								' else
								' 	iBulan = i
								' end if
								i	= iBulan
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
						<%
							Next
						%>
									</select>
								</td>
								<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
							            " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba = rsFcba("fcba") then
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
						<td align="center"><input type="submit" value="View" /></td>
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			
			<div id="subheader" class="centered">
				<h2>LAPORAN TURN OVER TENAGA KERJA</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	 <tr>
					<th rowspan="2">Kebun</th>
					<th rowspan="2">Division</th>
					<th colspan="4">Pemeliharaan</th>
					<th colspan="4">Panen</th>
					<th colspan="4">Supervisi</th>
					<th colspan="4">Traksi</th>
					<th colspan="4">Kantor/Umum</th>
					<th colspan="4">Total</th>
				</tr>

				<tr>
					<th>AWAL</th>
					<th>MASUK</th>
					<th>KELUAR</th>
					<th>AKHIR</th>

					<th>AWAL</th>
					<th>MASUK</th>
					<th>KELUAR</th>
					<th>AKHIR</th>

					<th>AWAL</th>
					<th>MASUK</th>
					<th>KELUAR</th>
					<th>AKHIR</th>

					<th>AWAL</th>
					<th>MASUK</th>
					<th>KELUAR</th>
					<th>AKHIR</th>

					<th>AWAL</th>
					<th>MASUK</th>
					<th>KELUAR</th>
					<th>AKHIR</th>

					<th>AWAL</th>
					<th>MASUK</th>
					<th>KELUAR</th>
					<th>AKHIR</th>
				</tr>

    </thead>
	<tbody>
		<%
		Set rsCat = connStr.Execute(strPusingan)
		if Not rsCat.eof then
			While Not rsCat.EOF
			TOTALKELUARPML = cdbl(to_nol(rsCat("Jumlah_TK_PML_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_PML_MASUK"))) - cdbl(to_nol(rsCat("Jumlah_TK_PML_KELUAR")))
			TOTALKELUARPNN = cdbl(to_nol(rsCat("Jumlah_TK_PNN_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_PNN_MASUK"))) - cdbl(to_nol(rsCat("Jumlah_TK_PNN_KELUAR")))
			TOTALKELUARMD = cdbl(to_nol(rsCat("Jumlah_TK_MD_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_MD_MASUK"))) - cdbl(to_nol(rsCat("Jumlah_TK_MD_KELUAR")))
			TOTALKELUARTR = cdbl(to_nol(rsCat("Jumlah_TK_TR_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_TR_MASUK"))) - cdbl(to_nol(rsCat("Jumlah_TK_TR_KELUAR")))
			TOTALKELUARKU = cdbl(to_nol(rsCat("Jumlah_TK_KU_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_KU_MASUK"))) - cdbl(to_nol(rsCat("Jumlah_TK_KU_KELUAR")))

			TOTALAWALTOTAL =  cdbl(to_nol(rsCat("Jumlah_TK_PML_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_PNN_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_MD_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_TR_AWAL"))) + cdbl(to_nol(rsCat("Jumlah_TK_KU_AWAL")))
			TOTALMASUKTOTAL =  cdbl(to_nol(rsCat("Jumlah_TK_PML_MASUK"))) + cdbl(to_nol(rsCat("Jumlah_TK_PNN_MASUK"))) + cdbl(to_nol(rsCat("Jumlah_TK_MD_MASUK"))) + cdbl(to_nol(rsCat("Jumlah_TK_TR_MASUK"))) + cdbl(to_nol(rsCat("Jumlah_TK_KU_MASUK")))
			TOTALKELUARTOTAL =  cdbl(to_nol(rsCat("Jumlah_TK_PML_KELUAR"))) + cdbl(to_nol(rsCat("Jumlah_TK_PNN_KELUAR"))) + cdbl(to_nol(rsCat("Jumlah_TK_MD_KELUAR"))) + cdbl(to_nol(rsCat("Jumlah_TK_TR_KELUAR"))) + cdbl(to_nol(rsCat("Jumlah_TK_KU_KELUAR")))
			TOTALAKHIRTOTAL = TOTALAWALTOTAL + TOTALMASUKTOTAL - TOTALKELUARTOTAL
			
			TOTALAWALBOTPML = TOTALAWALBOTPML + cdbl(to_nol(rsCat("Jumlah_TK_PML_AWAL")))
			TOTALMASUKBOTPML = TOTALMASUKBOTPML + cdbl(to_nol(rsCat("Jumlah_TK_PML_MASUK")))
			TOTALKELUARBOTPML = TOTALKELUARBOTPML + cdbl(to_nol(rsCat("Jumlah_TK_PML_KELUAR")))
			TOTALAKHIRBOTPML = TOTALAKHIRBOTPML + TOTALKELUARPML

			TOTALAWALBOTPNN = TOTALAWALBOTPNN + cdbl(to_nol(rsCat("Jumlah_TK_PNN_AWAL")))
			TOTALMASUKBOTPNN = TOTALMASUKBOTPNN + cdbl(to_nol(rsCat("Jumlah_TK_PNN_MASUK")))
			TOTALKELUARBOTPNN = TOTALKELUARBOTPNN + cdbl(to_nol(rsCat("Jumlah_TK_PNN_KELUAR")))
			TOTALAKHIRBOTPNN = TOTALAKHIRBOTPNN + TOTALKELUARPNN

			TOTALAWALBOTMD = TOTALAWALBOTMD + cdbl(to_nol(rsCat("Jumlah_TK_MD_AWAL")))
			TOTALMASUKBOTMD = TOTALMASUKBOTMD + cdbl(to_nol(rsCat("Jumlah_TK_MD_MASUK")))
			TOTALKELUARBOTMD = TOTALKELUARBOTMD + cdbl(to_nol(rsCat("Jumlah_TK_MD_KELUAR")))
			TOTALAKHIRBOTMD = TOTALAKHIRBOTMD + TOTALKELUARMD

			TOTALAWALBOTTR = TOTALAWALBOTTR + cdbl(to_nol(rsCat("Jumlah_TK_TR_AWAL")))
			TOTALMASUKBOTTR = TOTALMASUKBOTTR + cdbl(to_nol(rsCat("Jumlah_TK_TR_MASUK")))
			TOTALKELUARBOTTR= TOTALKELUARBOTTR + cdbl(to_nol(rsCat("Jumlah_TK_TR_KELUAR")))
			TOTALAKHIRBOTTR = TOTALAKHIRBOTTR + TOTALKELUARTR

			TOTALAWALBOTKU = TOTALAWALBOTKU + cdbl(to_nol(rsCat("Jumlah_TK_KU_AWAL")))
			TOTALMASUKBOTKU = TOTALMASUKBOTKU + cdbl(to_nol(rsCat("Jumlah_TK_KU_MASUK")))
			TOTALKELUARBOTKU = TOTALKELUARBOTKU + cdbl(to_nol(rsCat("Jumlah_TK_KU_KELUAR")))
			TOTALAKHIRBOTKU = TOTALAKHIRBOTKU + TOTALKELUARKU
			
			TOTALAWALTOTALBOT = TOTALAWALTOTALBOT + TOTALAWALTOTAL 
			TOTALMASUKTOTALBOT = TOTALMASUKTOTALBOT + TOTALMASUKTOTAL 
			TOTALKELUARTOTALBOT = TOTALKELUARTOTALBOT + TOTALKELUARTOTAL 
			TOTALAKHIRTOTALBOT = TOTALAKHIRTOTALBOT + TOTALAKHIRTOTAL 
			%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sectionname")%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_PML_AWAL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_PML_MASUK")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_PML_KELUAR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKELUARPML%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_PNN_AWAL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_PNN_MASUK")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_PNN_KELUAR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKELUARPNN%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_MD_AWAL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_MD_MASUK")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_MD_KELUAR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKELUARMD%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_TR_AWAL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_TR_MASUK")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_TR_KELUAR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKELUARTR%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_KU_AWAL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_KU_MASUK")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jumlah_TK_KU_KELUAR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKELUARKU%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALAWALTOTAL%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALMASUKTOTAL%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKELUARTOTAL%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALAKHIRTOTAL%></td>			
		</tr>

	<%
		rsCat.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='27'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
		<tr>
			<td colspan="2" align="center" style="background: #F4F4A0;"> TOTAL</td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAWALBOTPML%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALMASUKBOTPML%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKELUARBOTPML%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAKHIRBOTPML%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAWALBOTPNN%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALMASUKBOTPNN%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKELUARBOTPNN%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAKHIRBOTPNN%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAWALBOTMD%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALMASUKBOTMD%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKELUARBOTMD%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAKHIRBOTMD%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAWALBOTTR%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALMASUKBOTTR%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKELUARBOTTR%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAKHIRBOTTR%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAWALBOTKU%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALMASUKBOTKU%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKELUARBOTKU%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAKHIRBOTKU%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAWALTOTALBOT%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALMASUKTOTALBOT%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKELUARTOTALBOT%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALAKHIRTOTALBOT%></td>
		</tr>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</tbody>
</table>


<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
	</body>
</html>