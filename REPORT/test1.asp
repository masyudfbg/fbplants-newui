<%@Language=Vbscript%>
<% Option Explicit %>
<% Validatesession %>
<!--#Include Virtual="/Conndb1.Inc"-->
<!--#Include Virtual="/Txtfunction.Inc"-->
<!--#Include Virtual="/Adovbs.Inc"-->
 <%
Dim V_Txtyear, V_Txtmonth, V_Txtgangcode, V_Txtgrade
Dim Scriptname
Dim Rsemployee, Stremployee, Rsatt, Stratt, Rskj, Strkj
Dim No, Cclass, Nday
V_Txtyear      =    Request("Txtyear")
V_Txtmonth     =    Request("Txtmonth")
V_Txtgangcode  =    Request("Txtgangcode")
V_Txtgrade     =    Request("Txtgrade")
 Scriptname     =    Request.Servervariables("Script_Name") 
 If Request("Txtyear")="" Then
V_Txtyear      =     Year(Session("V_Employee"))
End If
 If Request("Txtmonth")="" Then
V_Txtmonth     =     Month(Session("V_Employee"))
End If
 Stremployee       = "Select Ep.Fcba," & _
 "To_Char(Ep.Adhocdate,'Yyyy') As Year," & _
 "To_Char(Ep.Adhocdate,'Mm') As Month, " & _
 "Em.Sectionname," & _
 "Em.Gangcode," & _
 "Ep.Empcode, " & _
 "Em.Fcname," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),1, Amount,0))As Day_1," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),2, Amount,0))As Day_2," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),3, Amount,0))As Day_3," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),4, Amount,0))As Day_4," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),5, Amount,0))As Day_5," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),6, Amount,0))As Day_6," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),7, Amount,0))As Day_7," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),8, Amount,0))As Day_8," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),9, Amount,0))As Day_9," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),10,Amount,0))As Day_10," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),11,Amount,0))As Day_11," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),12,Amount,0))As Day_12," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),13,Amount,0))As Day_13," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),14,Amount,0))As Day_14," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),15,Amount,0))As Day_15," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),16,Amount,0))As Day_16," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),17,Amount,0))As Day_17," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),18,Amount,0))As Day_18," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),19,Amount,0))As Day_19," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),20,Amount,0))As Day_20," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),21,Amount,0))As Day_21," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),22,Amount,0))As Day_22," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),23,Amount,0))As Day_23," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),24,Amount,0))As Day_24," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),25,Amount,0))As Day_25," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),26,Amount,0))As Day_26," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),27,Amount,0))As Day_27," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),28,Amount,0))As Day_28," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),29,Amount,0))As Day_29," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),30,Amount,0))As Day_30," & _
 "Sum ( Decode (To_Number (To_Char (Ep.Adhocdate,'Dd')),31,Amount,0))As Day_31," & _
 "Sum ( Ep.Amount) As Total " & _
 "From   Empallded Ep, Allowded Al, Employee Em " & _
 "Where Ep.Fcba = Al.Fcba And Ep.Allowdedcode = Al.Fccode " & _
 "And Ep.Fcba = Em.Fcba And Ep.Empcode = Em.Fccode  " & _
 "And Ep.Allowdedcode In ('A24','A25','A27','A37','A38','A39','A40') " & _
 "And Ep.Fcba ='" & Session ("Fcba") & "' "
 If  Len(Trim(V_Txtyear)) > 0 Then
    Stremployee   =             Stremployee & " And To_Char(Ep.Adhocdate,'Yyyy') = " & V_Txtyear
End If
If  Len(Trim(V_Txtmonth)) > 0 Then
               Stremployee       =             Stremployee & " And To_Char(Ep.Adhocdate,'Mm') = " & V_Txtmonth
End If
If  Len(Trim(V_Txtgangcode)) > 0 Then
    Stremployee   =             Stremployee & " And Em.Gangcode = '" & V_Txtgangcode & "'"
End If
                Stremployee       =             Stremployee &   " Group By Ep.Fcba," & _
 "To_Char(Ep.Adhocdate,'Yyyy'), " & _ 
 "To_Char(Ep.Adhocdate,'Mm') ," & _ 
 "Em.Sectionname," & _
 "Em.Gangcode," &_ 
 "Ep.Empcode," & _
 "Em.Fcname"
                Stremployee       =             Stremployee & " Order By Ep.Empcode Asc"
 Response.Write Stremployee
%>
 <Html>
  
 <Head>
 <Link Rel="Stylesheet" Href="/Css/Stylereport2.Css" Type="Text/Css" />
 <Title>Ipla'S Reporting - Premi Panen </Title>
 <Script Language="Javascript" Src="/Js/Date-Picker.Js"></Script> 
 <Script Language="Javascript">
 Var Newwin;
 Function Launchwin(Winurl,Winname,Winfeatures)
 {
                //This Launches A New Window And Then
                //Focuses It If Window.Focus() Is Supported.
                Newwin = Window.Open(Winurl,Winname,Winfeatures);
  
 }
 </Script>
 </Head>
  
 <Body>
                <Div Id="Wrapper">
                <Div Id="Header-Wrapper">
                <!--#Include File="Menutop.Asp"-->
                <Div Id="Header">
                               <H4>Pt. <%=Session("Company")%></H4>
                               <H5><%=Session("Ba")%></H5>
                </Div>
                <Div Id="Subheader" Class="Centered">
                               <H2>Laporan Hari Kerja Khl</H3>
                               <Br/>
  
<H4>Periode&nbsp;:&nbsp;<%=Textmonth(V_Txtmonth)%>&nbsp;<%=V_Txtyear%></H4>
                </Div>
                </Div>
 <Table Id="Form-Table" Class="Centered-Margin">
                <Form Action="<%=Scriptname%>" Method="Post" Name="Frmentry">
                <Tr>
                               <Td><Font Face="Verdana"
Size="1">Year</Font></Td><Td Width="4"><Input Type="Text" Name="Txtyear"
Width="4" Value="<%=V_Txtyear%>"></Td>
                               <Td><Font Face="Verdana"
Size="1">Month</Font></Td><Td Width="2"><Input Type="Text" Name="Txtmonth"
Width="2" Value="<%=V_Txtmonth%>"></Td>
                               <Td><Font Face="Verdana"
Size="1">Gangcode</Font></Td><Td><Select Name="Txtgangcode" ><Option
Value=""></Option>
                               <%
                               Dim Rsgangcode, Strgangcode
                               Strgangcode = " Select Fccode, Fcname From Gang Where Activation = 'Y' And Fcba = '" & Session("Fcba") & "' Order By Fccode Asc"
                               
                               Set Rsgangcode = Server.Createobject("Adodb.Recordset")
                               Rsgangcode.Open Strgangcode, Connstr, Adopenstatic 
                               
                               If Not Rsgangcode.Eof Then
                               Do While Not Rsgangcode.Eof
                               If Rsgangcode("Fccode") = V_Txtgangcode Then%>
                                       <Option Value="<%=Rsgangcode("Fccode")%>" Selected><%=Rsgangcode("Fccode")%>-<%=Rsgangcode("Fcname")%></Option> 
                                    <%Else%>
                                       <Option Value="<%=Rsgangcode("Fccode")%>"><%=Rsgangcode("Fccode")%>-<%=Rsgangcode("Fcname")%></Option>  
                                   <%End If
                                              Rsgangcode.Movenext
                               Loop
                               End If
                               Rsgangcode.Close
                               Set Rsgangcode = Nothing
                               %>
                               </Select></Td>
                               <Td><Input Type="Submit" Value="View" /></Td>
                </Tr>
                </Form>
 </Table>
  
 <Table Border="0" Width="100%" Id="Product-Table">
                <Tr>
                               <Th>No</Th>
                               <Th>No Pekerja</Th>
                               <Th>Nama Pekerja</Th>
                               <Th>Status</Th>
                               <Th>Gang</Th>
 <%For Nday = 1 To 31%>
 <Th><%=Nday%></Th>
 <%Next%>
                               <Th>Total</Th>
                </Tr>
  
 <%
 Set Rsemployee = Server.Createobject("Adodb.Recordset")
 Rsemployee.Open Stremployee, Connstr, Adopenstatic
  
 Dim Nday_1
 Dim Nday_2
 Dim Nday_3
 Dim Nday_4
 Dim Nday_5
 Dim Nday_6
 Dim Nday_7
 Dim Nday_8
 Dim Nday_9
 Dim Nday_10
 Dim Nday_11
 Dim Nday_12
 Dim Nday_13
 Dim Nday_14
 Dim Nday_15
 Dim Nday_16
 Dim Nday_17
 Dim Nday_18
 Dim Nday_19
 Dim Nday_20
 Dim Nday_21
 Dim Nday_22
 Dim Nday_23
 Dim Nday_24
 Dim Nday_25
 Dim Nday_26
 Dim Nday_27
 Dim Nday_28
 Dim Nday_29
 Dim Nday_30
 Dim Nday_31
 Dim Ntotal
  
 If Rsemployee.Eof And Rsemployee.Eof Then
 %>
                <Tr>
                               <Td Colspan="37">No Data Available</Td>
                </Tr>
 <%End If
                               No = 1
                               
                               Nday_1 = 0
                               Nday_2 = 0
                               Nday_3 = 0
                               Nday_4 = 0
                               Nday_5 = 0
                               Nday_6 = 0
                               Nday_7 = 0
                               Nday_8 = 0
                               Nday_9 = 0
                               Nday_10 = 0
                               Nday_11 = 0
                               Nday_12 = 0
                               Nday_13 = 0
                               Nday_14 = 0
                               Nday_15 = 0
                               Nday_16 = 0
                               Nday_17 = 0
                               Nday_18 = 0
                               Nday_19 = 0
                               Nday_20 = 0
                               Nday_21 = 0
                               Nday_22 = 0
                               Nday_23 = 0
                               Nday_24 = 0
                               Nday_25 = 0
                               Nday_26 = 0
                               Nday_27 = 0
                               Nday_28 = 0
                               Nday_29 = 0
                               Nday_30 = 0
                               Nday_31 = 0
                               Ntotal = 0
                               
                               While Not Rsemployee.Eof
                                              If No  Mod 2 = 1 Then
                                                             Cclass = "Odd"
                                              Else
                                                             Cclass = "Default"
                                              End If
 %>
                                              <Tr Class="<%=Cclass%>">
                                                             <Td
Align="Center"><%=No%></Td>
                                                             <Td
Align="Center"><%=Rsemployee("empcode")%></Td>
                                                             <Td
Align="Left"><%=Rsemployee("Fcname")%></Td>
                                                             <Td
Align="Center"><%=Rsemployee("sectionname")%></Td>
                                                             <Td
Align="Center"><%=Rsemployee("Gangcode")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_1")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_2")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_3")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_4")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_5")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_6")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_7")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_8")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_9")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_10")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_11")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_12")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_13")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_14")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_15")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_16")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_17")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_18")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_19")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_20")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_21")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_22")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_23")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_24")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_25")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_26")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_27")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_28")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_29")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_30")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Day_31")%></Td>
                                                             <Td
Align="Right"><%=Rsemployee("Total")%></Td>
 <%
                                              No = No + 1
                                              
                                              Nday_1 =  Nday_1+Cdbl(Rsemployee("Day_1"))
                                              Nday_2 =  Nday_2+Cdbl(Rsemployee("Day_2"))
                                              Nday_3 =  Nday_3+Cdbl(Rsemployee("Day_3"))
                                              Nday_4 =  Nday_4+Cdbl(Rsemployee("Day_4"))
                                              Nday_5 =  Nday_5+Cdbl(Rsemployee("Day_5"))
                                              Nday_6 =  Nday_6+Cdbl(Rsemployee("Day_6"))
                                              Nday_7 =  Nday_7+Cdbl(Rsemployee("Day_7"))
                                              Nday_8 =  Nday_8+Cdbl(Rsemployee("Day_8"))
                                              Nday_9 =  Nday_9+Cdbl(Rsemployee("Day_9"))
                                              Nday_10 =  Nday_10+Cdbl(Rsemployee("Day_10"))
                                              Nday_11 =  Nday_11+Cdbl(Rsemployee("Day_11"))
                                              Nday_12 =  Nday_12+Cdbl(Rsemployee("Day_12"))
                                              Nday_13 =  Nday_13+Cdbl(Rsemployee("Day_13"))
                                              Nday_14 =  Nday_14+Cdbl(Rsemployee("Day_14"))
                                              Nday_15 =  Nday_15+Cdbl(Rsemployee("Day_15"))
                                              Nday_16 =  Nday_16+Cdbl(Rsemployee("Day_16"))
                                              Nday_17 =  Nday_17+Cdbl(Rsemployee("Day_17"))
                                              Nday_18 =  Nday_18+Cdbl(Rsemployee("Day_18"))
                                              Nday_19 =  Nday_19+Cdbl(Rsemployee("Day_19"))
                                              Nday_20 =  Nday_20+Cdbl(Rsemployee("Day_20"))
                                              Nday_21 =  Nday_21+Cdbl(Rsemployee("Day_21"))
                                              Nday_22 =  Nday_22+Cdbl(Rsemployee("Day_22"))
                                              Nday_23 =  Nday_23+Cdbl(Rsemployee("Day_23"))
                                              Nday_24 =  Nday_24+Cdbl(Rsemployee("Day_24"))
                                              Nday_25 =  Nday_25+Cdbl(Rsemployee("Day_25"))
                                              Nday_26 =  Nday_26+Cdbl(Rsemployee("Day_26"))
                                              Nday_27 =  Nday_27+Cdbl(Rsemployee("Day_27"))
                                              Nday_28 =  Nday_28+Cdbl(Rsemployee("Day_28"))
                                              Nday_29 =  Nday_29+Cdbl(Rsemployee("Day_29"))
                                              Nday_30 =  Nday_30+Cdbl(Rsemployee("Day_30"))
                                              Nday_31 =  Nday_31+Cdbl(Rsemployee("Day_31"))
                                              Ntotal =  Ntotal+Cdbl(Rsemployee("Total"))
                                              
                                              Rsemployee.Movenext
                               Wend
 Rsemployee.Close
 Set Rsemployee = Nothing
 %>
  
                <Tr>
                               <Td Colspan="5"
Align="Center"><Strong>Total</Strong></Td>
                               <Td Align="Right"><%=Nday_1%></Td>
                               <Td Align="Right"><%=Nday_2%></Td>
                               <Td Align="Right"><%=Nday_3%></Td>
                               <Td Align="Right"><%=Nday_4%></Td>
                               <Td Align="Right"><%=Nday_5%></Td>
                               <Td Align="Right"><%=Nday_6%></Td>
                               <Td Align="Right"><%=Nday_7%></Td>
                               <Td Align="Right"><%=Nday_8%></Td>
                               <Td Align="Right"><%=Nday_9%></Td>
                               <Td Align="Right"><%=Nday_10%></Td>
                               <Td Align="Right"><%=Nday_11%></Td>
                               <Td Align="Right"><%=Nday_12%></Td>
                               <Td Align="Right"><%=Nday_13%></Td>
                               <Td Align="Right"><%=Nday_14%></Td>
                               <Td Align="Right"><%=Nday_15%></Td>
                               <Td Align="Right"><%=Nday_16%></Td>
                               <Td Align="Right"><%=Nday_17%></Td>
                               <Td Align="Right"><%=Nday_18%></Td>
                               <Td Align="Right"><%=Nday_19%></Td>
                               <Td Align="Right"><%=Nday_20%></Td>
                               <Td Align="Right"><%=Nday_21%></Td>
                               <Td Align="Right"><%=Nday_22%></Td>
                               <Td Align="Right"><%=Nday_23%></Td>
                               <Td Align="Right"><%=Nday_24%></Td>
                               <Td Align="Right"><%=Nday_25%></Td>
                               <Td Align="Right"><%=Nday_26%></Td>
                               <Td Align="Right"><%=Nday_27%></Td>
                               <Td Align="Right"><%=Nday_28%></Td>
                               <Td Align="Right"><%=Nday_29%></Td>
                               <Td Align="Right"><%=Nday_30%></Td>
                               <Td Align="Right"><%=Nday_31%></Td>
                               <Td Align="Right"><%=Ntotal%></Td>
                </Tr>
 </Table>
  
    <Table Width="90%" Class="Data">
                   <Tr>
                               <Td Width="50%" Align="Center"><Font
Face="Verdana" Size="1">Approved By,</Font></Td>
                               <Td Width="50%" Align="Center"><Font
Face="Verdana" Size="1">Prepared By,</Font></Td>
                   </Tr>
                   <Tr> 
                               <Td>&nbsp;</Td>
                               <Td>&nbsp;</Td>
                   </Tr>
                   <Tr> 
                               <Td>&nbsp;</Td>
                               <Td>&nbsp;</Td>
                   </Tr>
                   <Tr> 
                               <Td>&nbsp;</Td>
                               <Td>&nbsp;</Td>
                   </Tr>
                   <Tr> 
                               <Td Width="50%" Align="Center"><Font
Face="Verdana"
Size="2"><Strong><U><%=Session("Fcmanager")%></U></Strong></Td>
                               <Td Width="50%" Align="Center"><Font
Face="Verdana" Size="2"><Strong><U><%=Session("Ktu")%></U></Strong></Td>
                   </Tr>
                                 <Tr> 
                               <Td Width="50%" Align="Center"><Font
Face="Verdana" Size="1"><Strong><I>Manager</I></Strong></Td>
                               <Td Width="50%" Align="Center"><Font
Face="Verdana" Size="1"><Strong><I>Ktu</I></Strong></Td>
                   </Tr>
    </Table>
    </Body>
    </Html>
