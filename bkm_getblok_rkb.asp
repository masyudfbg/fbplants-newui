
<%
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
%>
<%
Response.LCID = 1046 
%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!--#includes file="include/JSON.asp" --> 
<!--#includes file="include/json.objectClass.asp"-->
<!--#includes file="include/aspJSON.asp" -->

<%


dim txtcari,txtFcba,txtsupplier,txtSearch,txtactivty,txtColom,sqlcoa,rscoa,txtjob_category,txtlokasi,tcby,bl
dim tahun,bulan
txtactivty 	= Request("activity")
txtcari 	= Request("keyword")
txtUnit 	= Session("fcba")
txtFcba 	= Request("lokasi")
tcby 		= Request("tcby")
bl 		    = Request("blok")
rkbno       = Request("rkbno")
tahun       = Request("tahun")
bulan       = Request("bulan")
afd       	= Request("section")

select case bl 

case 1  
	txtColom = "SELECT FCCODE id,FCCODE text " & _ 
	" FROM FIELD WHERE fcba = '" & txtUnit & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE FETCH FIRST 100 ROWS ONLY"					
case 9 
	txtColom = "SELECT locationcode id,locationcode text " & _ 
	" FROM RKBUPAHDTL WHERE fcba = '" & txtUnit & "' AND rkbno='" & rkbno &"' AND actcode = '" & txtactivty &"' AND locationcode like UPPER('%" & txtcari & "%') ORDER BY locationcode FETCH FIRST 100 ROWS ONLY"					

case else


		sqlcoa ="select job_category,job_forfieldstatus from jobmaster where fccode='" & txtactivty & "' /*and fcba='" & txtUnit & "'*/"
		Set rscoa = Server.CreateObject("ADODB.Recordset")
		rscoa.Open sqlcoa, connstr, adOpenStatic 



		if not rscoa.eof then 
			txtjob_category = rscoa("job_category")
			txtjob_forfieldstatus =  rscoa("job_forfieldstatus")
			'IF txtjob_forfieldstatus = "LC - TB" OR txtjob_forfieldstatus = "TBM" THEN 
			'	txtjob_forfieldstatus = "TBM"
			'ELSE 
			'	txtjob_forfieldstatus =  "TM"
			'END IF
		'	IF txtjob_forfieldstatus = "LC - TB" OR txtjob_forfieldstatus = "TBM" OR txtjob_forfieldstatus = "LC" THEN 
		'		txtjob_forfieldstatus = "TBM"
		'	ELSE 
		'		txtjob_forfieldstatus =  "TM"
		'	END IF
	'	Response.Write txtjob_category
		end if

		if tcby = 1 then 
				txtFcba = txtUnit
			else 
				txtFcba= Request("lokasi")
		end if


		if 	LEN(TRIM(txtjob_category)) > 0 then
			
				  SELECT CASE TRIM(txtjob_category) 
						CASE "FF"  
										txtColom = " SELECT FCCODE id,FCCODE || ' - ' || dx_block_status text FROM FIELD_HIST_TAB WHERE dx_block_status='" & txtjob_forfieldstatus & "' and year='" & tahun & "' AND month='" & bulan & "' AND fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "									
						CASE "CB"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM gcmaster WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND Item='CB' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "					
						CASE "GC"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM gcmaster WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND Item='GC' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "					
						CASE "BA"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM bank WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "		
						CASE "MA"  
										txtColom = " SELECT FCCODE id,FCCODE text  FROM Machine WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%')  ORDER BY FCCODE "	
						CASE "VH"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM Vehicle WHERE vehiclegroupcode <> 'WS' AND fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "	
						CASE "ST"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM store WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "	
						CASE "MN"  
										txtColom = " SELECT NurseryCode id,DESCRIPTION text FROM nursery WHERE fcba = '" & txtFcba & "' AND Status='Y' AND NurseryCode like UPPER('%" & txtcari & "%') ORDER BY NurseryCode "	
						CASE "PN"  
										txtColom = " SELECT NurseryCode id,DESCRIPTION text FROM nursery WHERE fcba = '" & txtFcba & "' AND Status='Y' AND NurseryCode like UPPER('%" & txtcari & "%') ORDER BY NurseryCode "	
						CASE "BM"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM building WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "			
						CASE "AP"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM Supplier WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') UNION ALL SELECT FCCODE,FCNAME FROM Transporter WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "			
						CASE "AR"  
										txtColom = " SELECT CustomerCode id,FCCODE text FROM customer WHERE fcba = '" & txtFcba & "' AND isactive='Y' ORDER BY CustomerCode "			
						CASE "CA"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM Contractor WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "			
						CASE "FA"  
										txtColom = " SELECT FCCODE id,FCNAME text FROM Fixedasset WHERE fcba = '" & txtFcba & "' AND isactive='1' AND FCCODE like UPPER('%" & txtcari & "%')  UNION ALL  SELECT FCCODE id,FCNAME text FROM BUILDING WHERE fcba = '" & txtFcba & "' AND activation='Y'  AND FCCODE like UPPER('%" & txtcari & "%')  "	
						CASE "WS"  
										txtColom = " SELECT FCCODE id,FCCODE text FROM Workshop WHERE fcba = '" & txtFcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "									
						CASE ELSE				
										
						txtColom =  " SELECT '' FCCODE, '-' FCNAME " & _
									" FROM DUAL"  

				END SELECT	
		end if
end select
' response.write txtColom

Set List = Server.CreateObject("ADODB.Recordset")

List.Open txtColom, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(txtColom)

Dim List_jsonString
List_jsonString = caljson.toJSON("results",rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

Response.write List_jsonString
%>