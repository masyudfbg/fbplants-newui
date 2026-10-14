
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


dim txtcari,fcba,txtsection ,txtlokasi,txtactivty
txtcari = Request("keyword")
txtsection = Request("sectionname")
txtlokasi = Request("lokasi")
fcba = session("fcba")
gancode = Request("gancode")
activitywip = Request("actcode")
txttglbkm   = Request("tglbkm")

txtMonth  = left(txttglbkm,2)
txtYear   = right(txttglbkm,4)


Set List = Server.CreateObject("ADODB.Recordset")

if gancode = 1 then 
    sqlview =  "select fccode as id, fccode || '-' || fcname as text from gang " & _
    "  where remarks='" & txtsection & "' and fcba='" & fcba & "' and fcname like UPPER('%" & txtcari & "%') order by fcname asc"

end if

if gancode = 2 then 
    sqlview =  "select fccode as id, fccode || '-' || fcname as text from gang " & _
    "  where remarks='" & txtsection & "' and fcba='" & fcba & "' and fcname like UPPER('%" & txtcari & "%') order by fcname asc"
end if

'tampilkan blok di contract agreement'
if gancode = 3 then 

    'sqlview =  "SELECT FCCODE id,FCCODE text " & _ 
	'" FROM FIELD WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE FETCH FIRST 100 ROWS ONLY"
   '' response.write sqlview

   dim sqlcoa,rscoa,txtjob_category,txtColom,txtjob_forfieldstatus
  '  txtactivty = Request("actcode")
    sqlcoa ="select job_category,job_forfieldstatus from jobmaster where fccode='" & activitywip & "'"
    Set rscoa = Server.CreateObject("ADODB.Recordset")
    rscoa.Open sqlcoa, connstr, adOpenStatic 

    if not rscoa.eof then 
        txtjob_category = rscoa("job_category")
		txtjob_forfieldstatus =  rscoa("job_forfieldstatus")
    end if


    if 	LEN(TRIM(txtjob_category)) > 0 then
	
	      SELECT CASE TRIM(txtjob_category) 
				CASE "FF"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM FIELD_HIST_TAB WHERE dx_block_status='" & txtjob_forfieldstatus & "' AND fcba = '" & txtlokasi & "' AND YEAR='" & txtYear & "' AND MONTH='" & txtMonth & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%')  ORDER BY FCCODE "					
				CASE "CB"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM gcmaster WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND Item='CB' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "					
				CASE "GC"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM gcmaster WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND Item='GC' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "					
				CASE "BA"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM bank WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "		
        		CASE "MA"  
								sqlview = " SELECT FCCODE id,FCCODE text  FROM Machine WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') UNION ALL SELECT FCCODE id,FCCODE text  FROM station WHERE fcba = '" & fcba & "' AND FCCODE like UPPER('%" & txtcari & "%') "	
				
				
				CASE "VH"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM Vehicle WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "	
				CASE "ST"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM store WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "	
				CASE "MN"  
								sqlview = " SELECT NurseryCode id,DESCRIPTION text FROM nursery WHERE fcba = '" & txtlokasi & "' AND Status='Y' AND NurseryCode like UPPER('%" & txtcari & "%') ORDER BY NurseryCode "	
				CASE "PN"  
								sqlview = " SELECT NurseryCode id,DESCRIPTION text FROM nursery WHERE fcba = '" & txtlokasi & "' AND Status='Y' AND NurseryCode like UPPER('%" & txtcari & "%') ORDER BY NurseryCode "	
	      		CASE "BM"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM building WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "			
				CASE "AP"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM Supplier WHERE fcba = '" & fcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') UNION ALL SELECT FCCODE,FCNAME FROM Transporter WHERE fcba = '" & fcba & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "			
				CASE "AR"  
								sqlview = " SELECT CustomerCode id,FCCODE text FROM customer WHERE fcba = '" & txtlokasi & "' AND isactive='Y' ORDER BY CustomerCode "			
				CASE "CA"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM Contractor WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "			
				CASE "FA"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM Fixedasset WHERE fcba = '" & txtlokasi & "' AND isactive='1' AND FCCODE like UPPER('%" & txtcari & "%') UNION ALL SELECT FCCODE,FCCODE || ' - ' || FCNAME AS FCNAME FROM BUILDING WHERE fcba = '" & fcba & "' AND FCCODE like UPPER('%" & txtcari & "%') "	
        		CASE "WS"  
								sqlview = " SELECT FCCODE id,FCCODE text FROM Workshop WHERE fcba = '" & txtlokasi & "' AND ACTIVATION='Y' AND FCCODE like UPPER('%" & txtcari & "%') ORDER BY FCCODE "									
				CASE ELSE				
	 							
				sqlview =  " SELECT '' FCCODE, '-' FCNAME " & _
	      				    " FROM DUAL"  

        END SELECT	
        'Response.write sqlview
    end if

end if

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("results",rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

Response.write List_jsonString
%>