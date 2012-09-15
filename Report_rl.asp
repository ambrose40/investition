<html>
<Head><meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<%b= Server.MapPath("\")%>
<%if request.Cookies("StyleInv")="" then%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<%set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")
  s=servFileStream.ReadLine
  servFileStream.Close%>
<link rel="stylesheet" href="<%=s%>" type="text/css">
<%else%>
<%s=request.Cookies("StyleInv")%>
<link rel="stylesheet" href="<%=s%>" type="text/css">
<%End if%>
<title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title>
</Head>
<body class="report">
<%fotnum=1%>
<%If Request.Form("btn")="OK" Then%>
<%ya=Request.Form("ye")%>
<%Else%>
<%ya=Request.QueryString("ye")%>
<%End if%>
<%If ya="" Then%>
<%ya=Year(Date())%>
<%mo=Month(Date())%>
<%da=Day(Date())%>
<%zz=mo-04%>

<%If zz>=0 Then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=ya-1%>
<%End If%>

<%Else%>
<%If Request.Form("btn")="OK" Then%>
<%ya=Request.Form("ye")%>
<%Else%>
<%ya=Request.QueryString("ye")%>
<%End if%>

<%End If%>
<%XYZ=0%>
<%set mdbo =  Server.CreateObject("ADODB.Connection")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")
  set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%>
<%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
<%mdbo.Open ConnectionString%>
<%set mdbol1 = Server.CreateObject("ADODB.Command")%>
<%set mdborl1 = Server.CreateObject("ADODB.Recordset")%>
<%mdbol1.ActiveConnection = mdbo%>
<%set mdbol2 = Server.CreateObject("ADODB.Command")%>
<%set mdborl2 = Server.CreateObject("ADODB.Recordset")%>
<%mdbol2.ActiveConnection = mdbo%>
<%set mdbol3 = Server.CreateObject("ADODB.Command")%>
<%set mdborl3 = Server.CreateObject("ADODB.Recordset")%>
<%mdbol3.ActiveConnection = mdbo%>
<%set mdbol4 = Server.CreateObject("ADODB.Command")%>
<%set mdborl4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbol4.ActiveConnection = mdbo%>
<%set mdbo4 = Server.CreateObject("ADODB.Command")%>
<%set mdbor4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4.ActiveConnection = mdbo%>
<%set mdbo5 = Server.CreateObject("ADODB.Command")%>
<%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5.ActiveConnection = mdbo%>
<%set mdbog = Server.CreateObject("ADODB.Command")%>
<%set mdborg = Server.CreateObject("ADODB.Recordset")%>
<%mdbog.ActiveConnection = mdbo%>
<%mdbol4.CommandText="SELECT DISTINCT PID FROM MAIN WHERE YEARR>='" & ya & "'"%>
<%mdborl4.Open mdbol4%>
<%Do Until mdborl4.EOF%>
<%DIP=Mdborl4("PID")%>
<%mdbo5.CommandText="UPDATE MAIN SET YEARBEG=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & dip & "' AND YEARR>='" & ya & "') WHERE PID='" & dip & "' AND YEARBEG<'" & ya & "'"%>
<%mdbor5.Open mdbo5%>
<%mdborl4.Movenext%>
<%Loop%>
<%mdborl4.Close%>

<table bordercolor="0F0F0F" border="1"  style="border-collapse: collapse">
<tr bgcolor="AAAAAA">
 <th rowspan="1">Nr</th>
 <th rowspan="1">Projekti Nimetus</th>
 <th rowspan="1">Наименование проекта</th>
</tr>

<tr class="repnum">
<%For nuu=1 to 3%>
 <td><%=nuu%></td>
<%Next%>
</tr>

<%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
<%yy= zo & np%>
<%Select Case yy%>
<%Case zo%>
<%yr=zo%>
<%Case np%>
<%yr=np%>
<%Case zo & np%>
<%yr=zo & " AND " & np%>
<%End Select%>

<%aa=0%><%ab=0%>
<%ac=0%>



<%mdbol1.CommandText="SELECT DISTINCT Pid, ProjCode,PC, OracleCode, PRojName,RusName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
<%mdborl1.Open mdbol1%>

<tr class="boldProjGrup">
<td>


</td>
<td>
 INVESTEERINGUD KOKKU  v&auml;lja arvatud plokkide renoveerimine

</td>
<td>
</td>
</tr>
<tr class="boldProjGrup">
<td>



</td>
<td>
 INVESTEERINGUD KOKKU koos plokkide renoveerimisega

</td>
<td>
</td>

</tr>
<%Do until mdborl1.EOF%>
<tr class="ProjGrup">
<td>

&nbsp;<Font color="0000FF"><%=mdborl1("Pid")%>&nbsp;|&nbsp;

<%=MID(mdborl1("PC"),2,1)%>.

</td>
<td>
<%=mdborl1("ProjName")%>

</td>
<td>
<%=mdborl1("RusName")%>

</td>

</tr>


<%mdbol2.CommandText="SELECT DISTINCT Pid,PC,ProjCode,ProjName,OracleCode,RusName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND  SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY ProjCode"%>
<%mdborl2.Open mdbol2%>

<%Do until mdborl2.EOF%>
<tr class="ProjGrup">
<td>

&nbsp;<Font color="0000FF"><%=mdborl2("Pid")%>&nbsp;|&nbsp;
<%=MID(mdborl2("PC"),2,1) & "." & MID(mdborl2("PC"),5,1)%>.

</td>
<td>
<%=mdborl2("ProjName")%>

</td>
<td>
<%=mdborl2("RusName")%>

</td>
</tr>

<%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentifier='C' AND Yearr>=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "' ORDER BY ENTERPRISE"%>
<%mdborl3.Open mdbol3%>

<%Do until mdborl3.EOF%>
<tr class="Enterp">
<td>


</td>
<td>
<%=mdborl3("EDescr")%>

</td>
<td>
</td>
</tr>
<%mdbol4.CommandText="SELECT DISTINCT dbo.Main.Pid, Main_1.ProjCode as PC FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier AND  Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >='" & ya & "') /*and Main_1.Yearr >='" & ya & "'*/ AND MAin_1.Enterprise='" & Mdborl3("Enterprise") & "' AND MAin_1.IDentifier='C' AND SUBSTRING(MAin_1.ProjCode,4,2)<>'00' AND SUBSTRING(MAin_1.ProjCode,7,2)<>'00' AND  SUBSTRING(MAin_1.ProjCode,1,5)='" & MID(mdborl2("PC"),1,5) & "' ORDER BY Main_1.ProjCode"%>
<%mdborl4.Open mdbol4%>

<%Do Until mdborl4.EOF%>
<%mdbog.CommandText="SELECT DISTINCT RusName,ProjName,PC,RenovBlock,Yearr,FootNote FROM inpl WHERE Yearr >= '" & ya & "' AND Pid = '" & Mdborl4("Pid") & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND IDentifier='C' ORDER BY PC,Yearr"%>
<%mdborg.Open mdbog%>
<%If MDBORl4("Pid")=abcde THEN%>
<%mdborl4.MoveNExt%>
<%ELSE%>
<%Abcde=MDBORl4("Pid")%>
<tr >
<td>
<%=mdborl4("Pid") & " | "%>
         <%if mid(mdborl4("PC"),8,1)=0 and mid(mdborl4("PC"),7,1)<>0 then%>
          <%a=REPLACE(MID(mdborl4("PC"),1,6), "0", "") & MID(mdborl4("PC"),7,2)%>
         <%Else%>
          <%a=REPLACE(mdborl4("PC"), "0", "")%>
	 <%End If%>
          <%If len(a)>=7 Then%>
	  <%If Right(mdborl4("PC"),2)="00" then%>
	   <%=a%>
	  <%Else%>
	   <%=a%>.
	  <%End if%>
	 <%Else%>
	  <%If Right(mdborl4("PC"),2)="00" then%>
	   <%=mid(a,1,6)%>
	  <%Else%>
	   <%=mid(a,1,6)%>.
	  <%End if%>
	 <%End if%>

</td>
<td>
<%If LEN(mdborl4("PC"))>9 and MID(mdborl4("PC"),10,2)="00" Then%>
<%=mdborg("ProjName")%>&nbspsealhulgas:
<%Else%>
<%=mdborg("ProjName")%>
<%End IF%>
</td>
<td>
<%If LEN(mdborl4("PC"))>9 and MID(mdborl4("PC"),10,2)="00" Then%>
<%=mdborg("RusName")%>&nbspв том числе:
<%Else%>
<%=mdborg("RusName")%>
<%End IF%>

&nbsp&nbsp&nbsp
<%If Mdborg("Footnote") & "e" <> "e" then%>
<a name=<%="vira" & Fotnum%>></a>
{<%=Fotnum%>}
<%fotnum=fotnum+1%>
<%end if%>
</td>
</tr>

<%mdborl4.Movenext%>
<%mdborg.Close%>
<%END IF%>
<%loop%>
<%mdborl4.Close%>
<%mdborl3.Movenext%>
<%loop%>

<%mdborl3.Close%>
<%mdborl2.Movenext%>
<%loop%>

<%mdborl2.Close%>
<%mdborl1.Movenext%>

<%Loop%>
<%mdborl1.Close%>



<%set mdbo8 = Server.CreateObject("ADODB.Command")%>
<%set mdbor8 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo8.ActiveConnection = mdbo%>
<%set mdbou = Server.CreateObject("ADODB.Command")%>
<%set mdboru = Server.CreateObject("ADODB.Recordset")%>
<%mdbou.ActiveConnection = mdbo%>

<%mdbo8.CommandText="SELECT DISTINCT PID, YEARR FROM MAIN ORDER BY PID,YEARR"%>
<%mdbor8.Open mdbo8%>
<%Do until mdbor8.EOF%>
<%If MDBOR8("Pid")=abcde THEN%>
<%mdbor8.MoveNExt%>
<%ELSE%>
<%Abcde=MDBOR8("Pid")%>
<%mdbou.CommandText="UPDATE MAIN SET YEARBEG='" & MDBOR8("Yearr") & "' WHERE PID='" & MDBOR8("Pid") & "'"%>
<%mdboru.Open mdbou%>
<%mdbor8.MoveNExt%>
<%END IF%>
<%Loop%>

<%Dim koku(13)%><%Dim kok2(13)%>
<tr>
<td colspan="3">

Kokku ettev&otildette kaupa

</td>
</tr>
<%mdbo4.CommandText="SELECT * FROM Enterprise ORDER BY ENTERPRISE"%>
<%mdbor4.Open mdbo4%>
<%Do until mdbor4.EOF%>
<tr class="boldEnterp">
<td>
</td>
<td>
<%=mdbor4("EDescr")%>
</td>
</tr>
<%mdbor4.Movenext%>
<%Loop%>

<tr class="bold">
<td>
</td>
<td>
Kokku
</td>
</tr>
<tr>
<td colspan="3">

Kokku ettev&otildette kaupa, v&auml;lja arvatud plokkide renoveerimine

</td>
</tr>
<%mdbor4.MoveFirst%>
<%Do until mdbor4.EOF%>
<tr class="boldEnterp">
<td>
</td>
<td>
<%=mdbor4("EDescr")%>
</td>
</tr>
<%mdbor4.Movenext%>
<%Loop%>

<tr class="bold">
<td>
</td>
<td>
Kokku
</td>
</tr>
</table>
</body>
</html>