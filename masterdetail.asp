<%
'########################################################################
'########################################################################
'### CDataGrid Calss (C) 2003 by Wafi Mohtaseb wafi_m@hotmail.com     ###
'### Version 1.0 												      ###
'### to get the full version component please contact me at :	      ###
'### wafi_m@hotmail.com or 00962795071467						      ###
'########################################################################
'########################################################################
ActiveCol = Request("ActiveCol")
PageNumber = Request("PageNumber")
IF PageNumber = "" Then PageNumber = 1

'-----------------------------------------------------------------------------------
'#######################################################################
'#########Class Definetion##############################################
Class CMasterDetailDataGrid
   '--------------------------------------------
   '############################################
   '########Private Var's
   Private vdbConn 
   Private vMasterRst 
   Private vDetailRst
   Private vConStr 
   Private vMasterSQL
   Private vDetailSQL
   Private vAddLink 
   Private vRelatedFields
   Private vConnectField
   Private vLinks
   Private vLinksURLs 
   Private vLinksParams
   Private vAllowPaging 
   Private vRecordsInPage 
   Private vShowNavigation
   Private vPageNumber 
   Private vPageName
   Private vAutoAddColumns
   Private vAddColumns 
   Private vTableBackGroundColor
   Private vTableWidth 
   Private vTableBorder
   Private vTableCellPadding 
   Private vTableCellSpacing
   Private vTableBorderColor 
   Private vAllowSorting
   Private vSortedColumns 
   Private vActiveColumn
   Private vImagesPath
   Private vDetailLinks
   Private vDetailLinksParams
   Private vDetailLinksURLs
   Private vAddDetailLinks
   
   Private vEvenRowBgColor
   Private vOddRowBgColor
   Private vNormalTextStyleClass
   Private vHeaderStyleClass
   '--------------------------------------------
   '--------------------------------------------
   '############################################
   '########'Consturctor,Destructor
   '########'Initialize Class Var's
   Private Sub Class_Initialize()
     Set vdbConn = Server.CreateObject("ADODB.Connection")
     Set vMasterRst = Server.CreateObject("ADODB.RecordSet") 
     Set vDetailRst = Server.CreateObject("ADODB.RecordSet")
     'Default Settings
     vAddLink = False
     vAddDetailLinks = False
     vAllowPaging = False
     vShowNavigation = False
     vAutoAddColumns = True
     vAllowSorting = False
     vAddColumns = ""
     vConStr = ""
     vRelatedFields = ""
     vMasterSQL = ""
     vDetailSQL = ""
     vActiveColumn = ""
     vSortedColumns = ""
     vConnectField = ""
     vRecordsInPage = 20
     vPageNumber = 0
     '----------------------
     vTableBorderColor = "#FFD318"
     vTableBackGroundColor = "#333333"'"#f1c200"
     vTableWidth = "100%"
     vTableBorder = 0
     vTableCellPadding = 0
     vTableCellSpacing = 0
     vImagesPath = "http://"+Request.ServerVariables("REMOTE_HOST")+"/img/"
     
     vEvenRowBgColor = "#f1c200"
     vOddRowBgColor = "#FFD318"
     vNormalTextStyleClass = "StudentText"
     vHeaderStyleClass = "headerlight"
   End Sub
   '--------------------------------------------------
   Private Sub Class_Terminate()
     
     Set vDetailRst = Nothing
     vMasterRst.Close
     Set vMasterRst = Nothing
     vdbConn.Close
     Set vdbConn = Nothing 
   End Sub
   '############################################
   '--------------------------------------------
   '--------------------------------------------
   '############################################
   '########'Properties
   'Get Connection String
   Public Property Let ConnectionString(sConStr)
      vConStr = sConStr
   End Property
   '-----------------------------------------
   'Get Master SQL Statment
   Public Property Let MasterSQLStatmnet(sMasterSQL)
      vMasterSQL = sMasterSQL
   End Property
   '-----------------------------------------
   'Get Master SQL Statment
   Public Property Let DetailSQLStatmnet(sDetailSQL)
      vDetailSQL = sDetailSQL
   End Property
   '-----------------------------------------
   'Check if there links
   Public Property Let AddLink(sAddLink)
      vAddLink = sAddLink
   End Property
   '-----------------------------------------
   'Get User Links Values
   Public Property Let Links(sLinks)
      vLinks = sLinks
   End Property
   '--------------------------------------------
   'Get Links URL
   Public Property Let LinksURLs(sLinksURLs)
     vLinksURLs = sLinksURLs
   End Property
   '--------------------------------------------
   'Get Links Parameters
   Public Property Let LinksParams(sLinksParams)
     vLinksParams = sLinksParams
   End Property
   '--------------------------------------------
   'Check If User Want Paging
   Public Property Let AllowPaging(sAllowPaging)
     vAllowPaging = sAllowPaging
   End Property
   '--------------------------------------------
   'Get Records To View per Page
   Public Property Let RecordsInPage(sRecordsInPage)
     vRecordsInPage = sRecordsInPage
   End Property
   '--------------------------------------------
   'Show Next & Preivous
   Public Property Let ShowNavigation(sShowNavigation)
     vShowNavigation = sShowNavigation
   End Property
   '--------------------------------------------
   'Get Current Page Number
   Public Property Let PageNumber(sPageNumber)
     vPageNumber = sPageNumber
   End Property
   '--------------------------------------------
   'Add Columns Automatic
   Public Property Let AutoAddColumns(sAutoAddColumns)
     vAutoAddColumns = sAutoAddColumns
   End Property
   '-------------------------------------------------
   'Add Columns That User want to Add 
   Public Property Let AddColumns(sAddColumns)
     vAddColumns = sAddColumns
   End Property
   '-------------------------------------------------
   'Get Table Width
   Public Property Let TableWidth(sTableWidth)
     vTableWidth = sTableWidth
   End Property
   '-------------------------------------------------
   'Get Table Border
   Public Property Let TableBorder(sTableBorder)
     vTableBorder = sTableBorder
   End Property
   '-------------------------------------------------
   'Get Table BackGround Color
   Public Property Let TableBackGroundColor(sTableBackGroundColor)
     vTableBackGroundColor = sTableBackGroundColor
   End Property
   '-------------------------------------------------
   'Get Cell CellPadding
   Public Property Let TableCellPadding(sTableCellPadding)
     vTableCellPadding = sTableCellPadding
   End Property
   '-------------------------------------------------
   'Get CellSpacing
   Public Property Let TableCellSpacing(sTableCellSpacing)
     vTableCellSpacing = sTableCellSpacing
   End Property
   '-------------------------------------------------
   'Get Table Border Color
   Public Property Let TableBorderColor(sTableBorderColor)
     vTableBorderColor = sTableBorderColor
   End Property
    '-------------------------------------------------
   'Get Even Row bgColor
   Public Property Let EvenRowBgColor(sEvenRowBgColor)
     vEvenRowBgColor = sEvenRowBgColor
   End Property
   '-------------------------------------------------
   'Get Odd	 Row bgColor
   Public Property Let OddRowBgColor(sOddRowBgColor)
     vOddRowBgColor = sOddRowBgColor
   End Property
   '-------------------------------------------------
   'Normal Text Style
   Public Property Let NormalTextStyleClass(sNormalTextStyleClass)
	 vNormalTextStyleClass = sNormalTextStyleClass
   End Property
   '-------------------------------------------------
   'Header Text Style
   Public Property Let HeaderStyleClass(sHeaderStyleClass)
	 vHeaderStyleClass = sHeaderStyleClass
   End Property
   '-------------------------------------------------
   '-------------------------------------------------
   'Allow Column Sorting
   Public Property Let AllowSorting(sAllowSorting)
     vAllowSorting = sAllowSorting
   End Property
   '-------------------------------------------------
   'Allow Column Sorting
   Public Property Let SortedColumns(sSortedColumns)
     vSortedColumns = sSortedColumns
   End Property
   '-------------------------------------------------
   'Allow Column Sorting
   Public Property Let ActiveColumn(sActiveColumn)
     vActiveColumn = sActiveColumn
   End Property
   '-------------------------------------------------
   'Allow Column Sorting
   Public Property Let ConnectField(sConnectField)
     vConnectField = sConnectField
   End Property
   '-------------------------------------------------
   'Allow Column Sorting
   Public Property Let RelatedFields(sRelatedFields)
     vRelatedFields = sRelatedFields
   End Property
   '-------------------------------------------------
   'Add Images Path
   Public Property Let ImagesPath(sImagesPath)
      vImagesPath = sImagesPath
   End Property
   '-----------------------------------------------------
   Public Property Let DetailLinks(sDetailLinks)
      vDetailLinks = sDetailLinks
   End Property
   '-------------------------------------------------------
   Public Property Let DetailLinksParams(sDetailLinksParams)
	  vDetailLinksParams = sDetailLinksParams
   End Property
   '---------------------------------------------------------
   Public Property Let DetailLinksURLs(sDetailLinksURLs)
	  vDetailLinksURLs = sDetailLinksURLs	
   End Property
   '-------------------------------------------------------
   Public Property Let AddDetailLinks(sAddDetailLinks)
      vAddDetailLinks = sAddDetailLinks
   End Property
   
   '#####################################################
   '--------------------------------------------
   '--------------------------------------------
   'Binding Data
   Public Sub BindData()
      'Open DataBase Connection
	  vConStr = "Provider=OraOledb.Oracle;Data Source=IPLASDBPROD;User ID=IPLASPROD;Password=sac"
      vdbConn.Open vConStr
      'Set Cursor Location to Client Side
      vMasterRst.CursorLocation = adUseClient
      'Open recordset
      vMasterRst.Open vMasterSQL,vdbConn,adOpenDynamic,adLockReadOnly 
      'Check If Sorting is Allowed
      IF vAllowSorting Then
         'Set Sorting Field
         IF vActiveColumn = "" Then 
           vMasterRst.Sort = "["&vMasterRst.Fields(0).Name&"]"
         Else
           vMasterRst.Sort = "["&vActiveColumn&"]"
         End IF  
      End IF   
      'If Error The Print Error Description
      IF Err.number <> 0 Then
         Response.Write "<TR><Td Align=center Class="&vNormalTextStyleClass&"><Font Color=Red><b>"&Err.Description&"</td></TR>"&vbCrLf   
         Exit Sub 
      End IF   
      'Begin Data Binding
      Response.Write "<Table Border="&vTableBorder&" Width="&vTableWidth&" bordercolor="&vTableBorderColor&" bgColor="&vTableBackGroundColor&" cellSpacing="&vTableCellSpacing&" cellPadding="&vTableCellPadding&" >"&vbCrLf 
      Call Check_Paging()
      Response.Write "<Tr Class="&vHeaderStyleClass&">"&vbCrLf 
      'Check If Columns Added Automatic
      IF vAutoAddColumns Then
        Call Add_Columns_Auto()
      Else
        'Add Columns Manually
        Call Add_Columns(vAddColumns)
      End IF
      Response.Write "</Table>"&vbCrLf 
   End Sub
   '-------------------------------------------------------------------------
   'Check paging
   Private Sub Check_Paging()
      'If Paging is Allowed
      IF vAllowPaging Then 
         IF vRecordsInPage < 1 OR vRecordsInPage > vMasterRst.RecordCount Then vRecordsInPage = 10
         vMasterRst.PageSize = vRecordsInPage
         vMasterRst.CacheSize = vMasterRst.PageSize
         'bulid paging navigation
         Call Build_Page_Navigation()
      Else
         vMasterRst.PageSize = vMasterRst.RecordCount 
         vMasterRst.CacheSize = vMasterRst.PageSize   
      End IF
      
   End Sub
   '-------------------------------------------------------------------------
   'Add Columns Manually
   Private Sub Add_Columns(MyColumns)
      pMyColumns  = Split(MyColumns,",")
      Response.Write "<td width=""1%"">&nbsp;</td>"
      For vColumnsNumber = LBound(pMyColumns) To UBound(pMyColumns)
         'Check if Sorting is Active
         IF vAllowSorting Then
           IF InStr(vSortedColumns,pMyColumns(vColumnsNumber)) Then 
              Response.Write "<TD><a href="""&vPageName&"?PageNumber="&Request("PageNumber")&"&ActiveCol="&pMyColumns(vColumnsNumber)&"""><Font Color=White>"&pMyColumns(vColumnsNumber)&"</a></Td>"&vbCrLf     
           Else
              Response.Write "<TD>"&pMyColumns(vColumnsNumber)&"</Td>"&vbCrLf     
           End IF
         Else 
           Response.Write "<TD>"&pMyColumns(vColumnsNumber)&"</Td>"&vbCrLf     
         End IF  
      Next  
      
       IF vAddLink Then Response.Write "<td width=""10%"">&nbsp;</td>"&vbCrLf
      
      Response.Write "</Tr>"&vbCrLf
      ReCount = 0
      
      For vPageSize = 0 To vMasterRst.PageSize
         IF ReCount Mod 2 = 0 Then
           BgColor = vEvenRowBgColor
         Else
           BgColor = vOddRowBgColor
         End IF
         Response.Write "<Tr Class="&vNormalTextStyleClass&" bgcolor="&BgColor&">"&vbCrLf 
         Response.Write "<td width=""1%""><a href=""javascript:hidediv('"&vMasterRst(""&vConnectField&"")&"')"" name="""&vMasterRst(""&vConnectField&"")&"link"" id="""&vMasterRst(""&vConnectField&"")&"link""><img src="+vImagesPath+"Minus.gif border=0></a></td>"
         For vColumnsNumber = LBound(pMyColumns) To UBound(pMyColumns)
           On Error Resume Next
           Response.Write "<Td>"&vMasterRst(pMyColumns(vColumnsNumber))&"</Td>"&vbCrLf 
           IF Err.number <> 0 Then 
             Response.Write "<Td><b><font color=red>Column "&pMyColumns(vColumnsNumber)&" Not Found In RecordSet<br>Columns Name Must Be Identical to SQL Statment</td>"&vbCrLf
           End IF
         Next
         IF vAddLink Then 
           
           DLinks = Insert_Links(vLinks,vLinksURLs,vLinksParams,False)
           Response.Write "<Td>"&DLinks&"</Td>"&vbCrLf 
         End IF  
         Response.Write "</Tr>"&vbCrLf 
         BindDetailData()
         vMasterRst.MoveNext
         
         ReCount = ReCount + 1
         IF vMasterRst.EOF Then Exit For
       Next   
      
       
   End Sub
   '-------------------------------------------------------------------------
   'Add Columns Automatic
   Private Sub Add_Columns_Auto()
      
      'Response.Write "<Tr Class="&vHeaderStyleClass&">"&vbCrLf 
      Response.Write "<td width=""1%"">&nbsp;</td>"
      For I = 0 To vMasterRst.Fields.Count - 1
        'Check if Sorting is Active
        IF vAllowSorting Then
          IF InStr(vSortedColumns,vMasterRst.Fields(I).Name) Then
            Response.Write "<TD><a href="""&vPageName&"?PageNumber="&Request("PageNumber")&"&ActiveCol="&vMasterRst.Fields(I).Name&"""><Font Color=White>"&vMasterRst.Fields(I).Name&"</a></Td>"&vbCrLf     
          Else
            Response.Write "<TD>"&vMasterRst.Fields(I).Name&"</Td>"&vbCrLf     
          End IF
        Else
          Response.Write "<TD>"&vMasterRst.Fields(I).Name&"</Td>"&vbCrLf     
        End IF
      Next
      
      IF vAddLink Then Response.Write "<td width=""10%"">&nbsp;</td>"&vbCrLf
      Response.Write "</Tr>"&vbCrLf
      ReCount = 0
      
      For vPageSize = 0 To vMasterRst.PageSize
         IF ReCount mod 2 = 0 Then
           BgColor = vEvenRowBgColor
         Else
           BgColor = vOddRowBgColor
         End IF
         Response.Write "<Tr Class="&vNormalTextStyleClass&" bgcolor="&BgColor&">"&vbCrLf 
         Response.Write "<td width=""1%""><a href=""javascript:hidediv('"&vMasterRst(""&vConnectField&"")&"')"" name="""&vMasterRst(""&vConnectField&"")&"link"" id="""&vMasterRst(""&vConnectField&"")&"link""><img src="+vImagesPath+"Minus.gif border=0></a></td>"
         For I = 0 To vMasterRst.Fields.Count - 1
           IF IsNull(vMasterRst.Fields(I))OR (vMasterRst.Fields(I)="") Then
             Response.Write "<Td>&nbsp;</Td>"&vbCrLf
           Else
             Response.Write "<Td>"&vMasterRst.Fields(I)&"</Td>"&vbCrLf 
           End IF  
         Next
         IF vAddLink Then 
           DLinks = Insert_Links(vLinks,vLinksURLs,vLinksParams,False)
           Response.Write "<Td>"&DLinks&"</Td>"&vbCrLf 
         End IF  
         Response.Write "</Tr>"&vbCrLf 
         BindDetailData()
         vMasterRst.MoveNext
         
         ReCount = ReCount + 1
         IF vMasterRst.EOF Then Exit For
       Next   
   End Sub
   '-------------------------------------------------------------------------
   'Build Paging Navigation
   Private Sub Build_Page_Navigation()
   'Get Page Name For Navigation purpose
      vPageName = Get_Page_Name()
      
      On Error Resume Next
        vPageNumber = CInt(PageNumber)
        IF Err.number <> 0 Then X = X
        
      IF CInt(vPageNumber) > vMasterRst.PageCount Then vPageNumber = vMasterRst.PageCount

	  IF CInt(vPageNumber) < 1 Then vPageNumber = 1  

	  IF vPageNumber = "" Then
	   	 vMasterRst.AbsolutePage = 1
	  Else
		 vMasterRst.AbsolutePage = vPageNumber
	  End IF   	           		 
	  '============================================
	  'Bulid Navigation
	  With Response
	    ColSpan = vMasterRst.Fields.Count + 1
	    IF vShowNavigation Then
			.Write "<Tr>"
			.Write "<Td align=center ColSpan="&ColSpan&">"&vbCrLf 
			.Write " <Table>"&vbCrLf 
			.Write "  <tr>"&vbCrLf 
	 
			IF CInt(vPageNumber) > 1 Then
			     .Write "<td valign=bottom class="&vNormalTextStyleClass&">"&vbCrLf 
				 .Write "<a href="""&vPageName&"?PageNumber="&vMasterRst.AbsolutePage - 1&"&ActiveCol="&Request("ActiveCol")&"""><b>Previous <<<b></a>"&vbCrLf 
				 .Write "</td>"&vbCrLf         
	        Else
				 .Write "<td valign=bottom class="&vNormalTextStyleClass&">"&vbCrLf 
				 .Write "<b>Previous <<<b>"&vbCrLf 
				 .Write "</td>"&vbCrLf
			End IF
	  
			.Write "<Td valign=bottom class="&vNormalTextStyleClass&">"&vbCrLf 
			.Write "Page "&vPageNumber&" of "&vMasterRst.PageCount
			.Write "</Td>"&vbCrLf 
	  
			IF CInt(vPageNumber) < vMasterRst.PageCount Then
			     .Write "<td valign=bottom class="&vNormalTextStyleClass&">"&vbCrLf 
				 .Write "<a href="""&vPageName&"?PageNumber="&vMasterRst.AbsolutePage + 1 &"&ActiveCol="&Request("ActiveCol")&"""><b>Next >></b></a>"&vbCrLf 
				 .Write "</td>"&vbCrLf         
		    Else
				.Write "<td valign=bottom class="&vNormalTextStyleClass&">"&vbCrLf 
				.Write "<b>Next >></b>"&vbCrLf 
				.Write "</td>"&vbCrLf
			End IF
			.Write "</tr>"&vbCrLf 
				 
			.Write "</Table>"&vbCrLf 
			.Write "</td>"&vbCrLf 
			.Write "</tr>"&vbCrLf 
	    End IF
	     		 
		.Write "<Tr>"&vbCrLf 
		.Write "<TD colspan="&ColSpan&" align=center Class="&vNormalTextStyleClass&">"&vbCrLf 
		For J = 1 To vMasterRst.PageCount 
		  IF J = vMasterRst.AbsolutePage Then
			   Response.Write "<B>"&J&"</B>"&vbCrLf 
			Else
			    .Write "<a href=""" & vPageName & "?PageNumber=" & J & "&ActiveCol="&Request("ActiveCol")&"""><b>" & J & "</a>&nbsp;" & VbCrLf
			End IF  
		Next
		.Write "</td>"&vbCrLf 
		.Write "</Tr>" &vbCrLf 
     End With		
   End Sub
   '-------------------------------------------------------------------------
   'Get Pgae Name
   Private Function Get_Page_Name()
      PageName = Server.MapPath(Request.ServerVariables("PATH_INFO"))
      Get_Page_Name = Mid(PageName, InstrRev(PageName, "\") + 1)
   End Function
   '-------------------------------------------------------------------------
   Private Function BindDetailData()
      Dim DLinks
      Dim MySQL
      MySQL = ""
      
      MySQL = vDetailSQL & " WHERE "&vRelatedFields&" = '" & vMasterRst(""&vConnectField&"") &"'"
      vDetailRst.Open MySQL,vdbConn,1,1
      
      'Response.Write "<tr><td colspan=15>"
      Response.Write "<tr><td colspan=15 id="""&vMasterRst(""&vConnectField&"")&""">"
      Response.Write "<Div> "
      'Response.Write "<Div id="""&vMasterRst(""&vConnectField&"")&"""> "
      Response.Write "<table width=""100%"" Border=1 bgcolor=#fafad2 bordercolor=maroon>"
      Response.Write "<tr Class="&vHeaderStyleClass&">"
      For I = 0 To vDetailRst.Fields.Count - 1 
       Response.Write "<td>"&vDetailRst.Fields(I).Name&"</td>"
      Next
      IF vAddDetailLinks Then Response.Write "<td width=""10%"">&nbsp;</td>"&vbCrLf
      Response.Write "</tr>"
      IF Not(vDetailRst.EOF) Then
        While Not(vDetailRst.EOF)
          Response.Write "<tr class="&vNormalTextStyleClass&">"
          For J = 0 To vDetailRst.Fields.Count - 1
            Response.Write "<td>"&vDetailRst.Fields(J)&"</td>"
          Next
          IF vAddDetailLinks Then
             DLinks = Insert_Links(vDetailLinks,vDetailLinksURLs,vDetailLinksParams,True)
			 Response.Write "<Td>"&DLinks&"</Td>"&vbCrLf 
          End IF
          
          Response.Write "</tr>"
          vDetailRst.MoveNext() 
        WEnd
      Else
        Response.Write "<tr class="&vNormalTextStyleClass&"><td align=center colspan="&vDetailRst.Fields.Count&"><b>No Records Foud</td></tr>"
      End IF
      vDetailRst.Close()
      Response.Write "</table>"
      Response.Write "</div>"
      Response.Write "</td></tr>"
   End Function
   
   '-------------------------------------------------------------------------
   'Add Links To the Table
   Private Function Insert_Links(MyLinks,MyLinksValues,MyLinksParams,MyFlag)
     Dim Local_Links
     sMyLinks = Split(MyLinks,",")
     sMyLinksValues = Split(MyLinksValues,",")
     
     IF MyLinksParams <> "" Then
       sParams = Split(MyLinksParams,",")
       For J = LBound(sParams) To UBound(sParams) 
          IF MyFlag = False Then
             sParams(J) = sParams(J) &"="&vMasterRst(""&sParams(J)&"")
          Else
             sParams(J) = sParams(J) &"="&vDetailRst(""&sParams(J)&"")
          End IF   
       Next
     End IF
     
     For I = LBound(sMyLinks) To UBound(sMyLinks)
       IF MyLinksParams <> "" Then
        Local_Links = Local_Links &"<a href="""&sMyLinksValues(I)&"?"&sParams(I)&"""><b>"&sMyLinks(I)&"</a>&nbsp;|&nbsp;"
       Else
        Local_Links = Local_Links &"<a href="""&sMyLinksValues(I)&"""><b>"&sMyLinks(I)&"</a>&nbsp;|&nbsp;"
       End IF
     Next
     On Error Resume Next
      Insert_Links = Left(Local_Links,Len(Local_Links)-7)
      IF Err.number <> 0 Then Response.Write "Error On Links Values" 
   End Function
   
   '----------------------------------------------------------------------------
   Public Sub About()
      Dim strAbout
      strAbout = "<Font Color=darkgreen face=Garamond size=3><b>"
      strAbout = strAbout &" <BR>Master Detail DataGrid Class version 1.0 <BR>"
      strAbout = strAbout &" Done By : Wafi Mohtaseb <BR>"
      strAbout = strAbout &" Email :(wafi_m@hotmail.com)<BR>"
      strAbout = strAbout &" Phone : +962795071467<BR><BR>"
      Response.Write strAbout
   End Sub
   '-------------------------------------------------------------------------
End Class
'======================================================
%>