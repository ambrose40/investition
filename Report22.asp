<%
Response.Expires = 0
Response.AddHeader "pragma", "no-cache"
%> 
<html>
<Head><meta http-equiv="Content-Type" content="text/html; charset=windows-1251"><title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title></Head>
<body bgcolor="FFFFFF">

<%If Request.Form("btn")="OK" Then%>
<%ya=Request.Form("ye")%>
<%Else%>
<%ya=Request.QueryString("ya")%>
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
<%ya=Request.QueryString("ya")%>
<%End if%>

<%End If%>
<img border="0" src="icons/invct.ico" Style=float:Left><p align="center"><p align="center"><a href="Main.asp"  target="_top"><font face="Verdana" Size="5" color="000099"><b> <u><%=ya%>-<%=ya+4%> m.a. SISEARUANNE</font></u></b></a></p><p>

<%zzz=ya%>
<%zzz=zzz-1%>
<%zzz2=zzz+2%>
<Form Method="POST" Action="Report_r.asp?ya=<%=ya%>">
<Input type="Submit" name="btn" size="10" Value="Kopeerimiseks">
<Input type="Submit" name="btn" size="10" Value="Redigeerimiseks">
<hr color="0000F5">
<a href="Report_r.asp?ya=<%=zzz%>"><Img border="0" src="icons/p.ico" Style=float:left></a><a href="Report_r.asp?ya=<%=zzz2%>"><Img border="0" src="icons/n.ico" Style=float:right></a>
<%Dim entt(10,13)%>
<%Dim ent2(10,13)%>
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<%b= Server.MapPath("\")%>
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
<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>
<%set mdbo1 = Server.CreateObject("ADODB.Command")%>
<%set mdbor1 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo1.ActiveConnection = mdbo%>
<%set mdbol4 = Server.CreateObject("ADODB.Command")%>
<%set mdborl4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbol4.ActiveConnection = mdbo%>
<%set mdbo5 = Server.CreateObject("ADODB.Command")%>
<%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5.ActiveConnection = mdbo%>

<%set mdbog = Server.CreateObject("ADODB.Command")%>
<%set mdborg = Server.CreateObject("ADODB.Recordset")%>
<%mdbog.ActiveConnection = mdbo%>
<%set mdbo4 = Server.CreateObject("ADODB.Command")%>
<%set mdbor4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4.ActiveConnection = mdbo%>


<table bordercolor="0F0F0F" border="1"  style="border-collapse: collapse">
<tr bgcolor="AAAAAA">
 <td rowspan="3"><Font Color="000000" Face="Verdana" Size="2">Nr</font></td>
 <td rowspan="3"><Font Color="000000" Face="Verdana"  Size="2">Projekti Nimetus</Font></td>
 <td rowspan="3"><Font Color="000000" Face="Verdana"  Size="2">Наименование проекта</Font></td>
 <td rowspan="2" colspan="2"><Font Color="000000" Face="Verdana"  Size="2">Ehitusperiood kvartal</Font></td>
 <td rowspan="3"><Font Color="000000" Face="Verdana"  Size="2">Kalkuleeritud maksumus kokku</Font></td>
 <td rowspan="3"><Font Color="000000" Face="Verdana"  Size="2">Viie aasta investeeringud kokku</Font></td>
 <td rowspan="3"><Font Color="000000" Face="Verdana"  Size="2">Tehtud seisuga 01.04.<%=ya%></Font></td>
 <td colspan="10" rowspan="1"><Font Color="000000" Face="Verdana"  Size="2">INVESTEERINGUD</Font></td>
</tr>
<tr bgcolor="AAAAAA">
 <td colspan="5"><Font Color="000000" Face="Verdana"  Size="2">Tegelik</Font></td>
 <td colspan="5"><Font Color="000000" Face="Verdana"  Size="2">Prognoos</Font></td>
</tr>

<tr bgcolor="AAAAAA">
<td><Font Color="000000" Face="Verdana"  Size="2">algus</Font></td>
<td><Font Color="000000" Face="Verdana"  Size="2">l&otilde;pp</Font></td>
<%For j=CDbl(ya-5) to Cdbl(ya+4)%>
<td><Font Color="000000" Face="Verdana"  Size="2"><%=j%></Font></td>
<%Next%>
</tr>

<tr  bgcolor="DDEFFF">
<%For nuu=1 to 18%>
 <td><Font Color="000000" Face="Verdana" Size="2"><%=nuu%></font></td>
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

<%mdbol1.CommandText="SELECT DISTINCT Pid, ProjCode,PC, OracleCode, PRojName,RusName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(ProjCode,4,2)='00' ORDER BY ProjCode"%>
<%mdborl1.Open mdbol1%>
<%mdbo1.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE RenovBlock=0 AND m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor1.Open mdbo1%>
<%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE RenovBlock=0 AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT, ISNULL(SUM(ISNULL(PAstSum,0)),0) as PASU FROM Main WHERE RenovBlock=0 AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
<%mdbor4.Open mdbo4%>
<tr bgcolor="FFFFAA">
<td>
<Font Color="000000" Face="Verdana" Size="2">
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b> INVESTEERINGUD KOKKU  v&auml;lja arvatud plokkide renoveerimine</b></Font>
</Font>
</td>
<td>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1c"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="a1c"%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1c"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1d"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="a1d"%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1d"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<b>
<%If mdbor4.EOF=True OR mdbor4("PASU") & "e" = "e" Then%>

<%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
<%sim=mdbor4("PASU")%>
<%Else%>
<%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4("PASU"))%>
<%End If%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1y"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
<%Else%>
  <%a0="a1y"%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="a1y"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1y"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<b>
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1z"%>
  <Font Color="000000" Face="Verdana" size="2"> 
     <%=Request.Form(a0)%>
  </Font>

<%Else%>
  <%a0="a1z"%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="a1z"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b> 
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<b><%If mdbor4.EOF=True OR mdbor4("PASU") & "e" = "e" Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("PASU")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1e"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    </Font>
<%If Request.Form(a0)="" Then%>
<%=sim%>
<%Else%>
<%=Request.Form(a0)%>  
<%End If%>
<%Else%>
  <%a0="a1e"%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="a1e"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1e"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b> 
</Font>
</td>

<%For ja=CDbl(ya-5) to CDbl(ya-1)%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> <B>
<%If mdbor2.BOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor2("Summi")%>
<%End If%>


  <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1f" & ja & "_1x"%>
    <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
    </Font>
    <%If Request.Form(a0)="" Then%>
      <input type="hidden" value="<%=Sim%>" name="<%="a1f" & ja & "_1x"%>">
    <%Else%>
      <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1f" & ja & "_1x"%>">
    <%End If%>
  <%Else%>
    <%a0="a1f" & ja & "_1x"%>
    <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a1f" & ja & "_1x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
    <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1f" & ja & "_1x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
    <%End If%>
  <%End If%>


<%mdbor2.Close%>
</B></Font></td>
<%Next%>

<%For ja=CDbl(ya) to CDbl(ya+4)%>
<td>
<b>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor5.EOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor5("SY")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
   <%a0="a" & ja & "_1x"%>
   <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
   </Font>

<%Else%>
   <%a0="a" & ja & "_1x"%>
   <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a" & ja & "_1x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
   <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "_1x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
   <%End If%>
<%End If%>
</Font>
</b>
</td>
<%If mdbor5.EOF=False Then%>
<%mdbor5.MoveNext%>
<%End If%>
<%Next%>
</tr>
<%mdbor1.Close%>
<%mdbor5.Close%>
<%mdbor4.Close%>
<%mdbo1.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor1.Open mdbo1%>
<%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT, ISNULL(SUM(ISNULL(PAstSum,0)),0) as PASU FROM Main WHERE Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
<%mdbor4.Open mdbo4%>
<tr bgcolor="FFFFAA">
<td>
<Font Color="000000" Face="Verdana" Size="2">

</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b> INVESTEERINGUD KOKKU koos plokkide renoveerimisega</b></Font>
</Font>
</td>
<td>
</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ac"%>">
<%Else%>
  <%a0="ac"%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>

</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ad"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad"%>">
<%Else%>
  <%a0="ad"%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<b>
<%If mdbor4.EOF=True OR mdbor4("PASU") & "e" = "e" Then%>

<%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
<%sim=mdbor4("PASU")%>
<%Else%>
<%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4("PASU"))%>
<%End If%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ay"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ay"%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ay"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<b>
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="az"%>
  <Font Color="000000" Face="Verdana" size="2"> 
     <%=Request.Form(a0)%>
  </Font>

<%Else%>
  <%a0="az"%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="az"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="az"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b> 
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b>
<%If mdbor4.EOF=True OR mdbor4("PASU") & "e" = "e" Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("PASU")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ae"%>
  <Font Color="000000" Face="Verdana" size="2"> 
    </Font>
<%If Request.Form(a0)="" Then%>
<%=sim%>
<%Else%>
<%=Request.Form(a0)%>  

<%End If%>
<%Else%>
  <%a0="ae"%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ae"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b></Font></td>

<%For ja=CDbl(ya-5) to CDbl(ya-1)%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> <B>
<%If mdbor2.BOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor2("Summi")%>
<%End If%>


  <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="af" & ja & "x"%>
    <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
    </Font>

  <%Else%>
    <%a0="af" & ja & "x"%>
    <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="af" & ja & "x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
    <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
    <%End If%>
  <%End If%>


<%mdbor2.Close%>
</B></Font></td>
<%Next%>

<%For ja=CDbl(ya) to CDbl(ya+4)%>
<td>
<b>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor5.EOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor5("SY")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
   <%a0="a1" & ja & "x"%>
   <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
   </Font>

<%Else%>
   <%a0="a1" & ja & "x"%>
   <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a1" & ja & "x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
   <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1" & ja & "x"%>" size="10" style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFAA; border-width:0">
   <%End If%>
<%End If%>
</Font>
</b>
</td>
<%If mdbor5.EOF=False Then%>
<%mdbor5.MoveNext%>
<%End If%>
<%Next%>
</tr>
<%Do until mdborl1.EOF%>
<%mdbor1.Close%>
<%mdbor5.Close%>
<%mdbor4.Close%>
<%mdbo1.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi,LEFT(m.ProjCode,2) FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C') GROUP BY LEFT(m.ProjCode,2) ORDER BY LEFT(m.ProjCode,2)"%>
<%mdbor1.Open mdbo1%>
<%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT, ISNULL(SUM(ISNULL(PAstSum,0)),0) as PASU FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
<%mdbor4.Open mdbo4%>
<tr bgcolor="FFFFAA">
<td>
<Font Color="000000" Face="Verdana" Size="2">
&nbsp;<Font color="0000FF"><%=mdborl1("Pid")%></Font>&nbsp;|&nbsp;<%=mdborl1("PC")%>.
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborl1("ProjName")%></Font>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborl1("RusName")%></Font>
</Font>
</td>
<td>
</td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl1("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
 
<%Else%>
  <%a0="ac" & mdborl1("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ad" & mdborl1("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="ad" & mdborl1("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 

<%If mdbor4.EOF=True Then%>
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=True Then%>
<%sim=mdbor4("PASU")%>
<%Else%>
<%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4("PASU"))%>
<%End If%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ay" & mdborl1("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
<%else%>
  <%a0="ay" & mdborl1("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ay" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="az" & mdborl1("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
     <%=Request.Form(a0)%>
  </Font>

<%Else%>
  <%a0="az" & mdborl1("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="az" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>

 </Font></td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("PASU")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ae" & mdborl1("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ae" & mdborl1("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ae" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</Font></td>

<%For ja=CDbl(ya-5) to CDbl(ya-1)%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor2.BOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor2("Summi")%>
<%End If%>


<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="af" & ja & "x" & mdborl1("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="af" & ja & "x" & mdborl1("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
  <%End If%>
<%End If%>

<%mdbor2.Close%>
 </Font></td>
<%Next%>
<%For ja=CDbl(ya) to CDbl(ya+4)%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor5.EOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor5("SY")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
   <%a0="a" & ja & "x" & mdborl1("Pid")%>
   <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
   </Font>

<%Else%>
   <%a0="a" & ja & "x" & mdborl1("Pid")%>
   <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
   <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl1("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
   <%End If%>
<%End If%>
 </Font>
</td>
<%If mdbor5.EOF=False Then%>
<%mdbor5.MoveNext%>
<%End If%>
<%Next%>
</tr>


<%mdbol2.CommandText="SELECT DISTINCT Pid,PC,ProjCode,ProjName,OracleCode,RusName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(ProjCode,4,2)<>'00' AND SUBSTRING(ProjCode,7,2)='00' AND  SUBSTRING(ProjCode,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY ProjCode"%>
<%mdborl2.Open mdbol2%>


<%Do until mdborl2.EOF%>
<%mdbor1.Close%>
<%mdbor5.Close%>
<%mdbor4.Close%>
<%mdbo1.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi,LEFT(ProjCode,5) FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C') GROUP BY LEFT(ProjCode,5) ORDER BY LEFT(ProjCode,5)"%>
<%mdbor1.Open mdbo1%>
<%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT, ISNULL(SUM(ISNULL(PAstSum,0)),0) as PASU FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
<%mdbor4.Open mdbo4%>
<tr bgcolor="FFFFAA">
<td>
<Font Color="000000" Face="Verdana" Size="2">
&nbsp;<Font color="0000FF"><%=mdborl2("Pid")%></Font>&nbsp;|&nbsp;<%=mdborl2("PC")%>.
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborl2("ProjName")%></Font>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborl2("RusName")%></Font>
</Font>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="ac" & mdborl2("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ad" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="ad" & mdborl2("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>

</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 

<%If mdbor4.EOF=True Then%>
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=True Then%>
<%sim=mdbor4("PASU")%>
<%Else%>
<%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4("PASU"))%>
<%End If%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ay" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ay" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ay" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
</b>
</Font>

</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="az" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
     <%=Request.Form(a0)%>
  </Font>

<%Else%>
  <%a0="az" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="az" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>

 </Font></td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("PASU")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ae" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ae" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ae" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%End If%>
<%End If%>
 </Font></td>

<%For ja=CDbl(ya-5) to CDbl(ya-1)%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor2.BOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor2("Summi")%>
<%End If%>


<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="af" & ja & "x" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="af" & ja & "x" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFAA; border-width:0">
  <%End If%>
<%End If%>

<%mdbor2.Close%>
 </Font></td>
<%Next%>
<%For ja=CDbl(ya) to CDbl(ya+4)%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor5.EOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor5("SY")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
   <%a0="a" & ja & "x" & mdborl2("Pid")%>
   <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
   </Font>

<%Else%>
   <%a0="a" & ja & "x" & mdborl2("Pid")%>
   <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
   <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
   <%End If%>
<%End If%>
 </Font>
</td>
<%If mdbor5.EOF=False Then%>
<%mdbor5.MoveNext%>
<%End If%>
<%Next%>
</tr>

<%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentifier='C' AND Yearr>=" & ya & " AND SUBSTRING(ProjCode,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(ProjCode,1,2)='" & MID(mdborl2("PC"),1,2) & "'"%>
<%mdborl3.Open mdbol3%>

<%Do until mdborl3.EOF%>
<%mdbor1.Close%>
<%mdbor5.Close%>
<%mdbor4.Close%>
<%mdbo1.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi,LEFT(ProjCode,5) FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise = '" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C') GROUP BY LEFT(ProjCode,5),m.enterprise ORDER BY LEFT(ProjCode,5)"%>
<%mdbor1.Open mdbo1%>
<%mdbo5.CommandText="SELECT SUM(SummYe) as SY, Yearr FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT,ISNULL(SUM(ISNULL(PastSum,0)),0) as PASU FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
<%mdbor4.Open mdbo4%>
<tr bgcolor="AAFFFF">
<td>
<Font Color="000000" Face="Verdana" Size="2">
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborl3("EDescr")%></Font>
</Font>
</td>
<td>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #AAFFFF; border-width:0">
<%End If%>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
  
<%Else%>
  <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #AAFFFF; border-width:0">
<%End If%>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=True Then%>
<%sim=mdbor4("PASU")%>
<%Else%>
<%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4("PASU"))%>
<%End If%>
<%End If%>

<%entt(mdborl3("Enterprise"),1)=entt(mdborl3("Enterprise"),1) + CDbl(sim)%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ay" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ay" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ay" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #AAFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #AAFFFF; border-width:0">
<%End If%>
<%End If%>
</b>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%entt(mdborl3("Enterprise"),2)=entt(mdborl3("Enterprise"),2) + Cdbl(sim)%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
     <%=Request.Form(a0)%>
  </Font>

<%Else%>
  <%a0="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #AAFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #AAFFFF; border-width:0">
<%End If%>
<%End If%> 
</Font></td>
<td>
<Font Color="000000" Face="Verdana" Size="2">
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("PASU")%>
<%End If%>
<%entt(mdborl3("Enterprise"),3)=entt(mdborl3("Enterprise"),3) + Cdbl(sim)%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #AAFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #AAFFFF; border-width:0">
<%End If%>
<%End If%>
 </Font></td>
<%jo=4%>
<%For ja=CDbl(ya-5) to CDbl(ya-1)%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor2.BOF=True OR mdbor2("Summi") & "e" ="e" then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor2("Summi")%>
<%End If%>
<%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
<%jo=jo+1%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #AAFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #AAFFFF; border-width:0">
  <%End If%>
<%End If%>

<%mdbor2.Close%>
 </Font></td>
<%Next%>
<%For ja=CDbl(ya) to CDbl(ya+4)%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor5.EOF=True then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor5("SY")%>
<%End If%>
<%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
<%jo=jo+1%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
   <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
   <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
   </Font>

<%Else%>
   <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
   <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #AAFFFF; border-width:0">
   <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #AAFFFF; border-width:0">
   <%End If%>
<%End If%>
 </Font>
</td>
<%If mdbor5.EOF=False Then%>
<%mdbor5.MoveNext%>
<%End If%>
<%Next%>
</tr>
<%mdbol4.CommandText="SELECT DISTINCT Pid,LEFT(PC,8),ne FROM inpl WHERE Enterprise='" & Mdborl3("Enterprise") & "' AND IDentifier='C' AND Yearr>=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)<>'00' AND  SUBSTRING(PC,1,5)='" & MID(mdborl2("PC"),1,5) & "' ORDER BY LEFT(PC,8),NE"%>
<%mdborl4.Open mdbol4%>
<%Do Until mdborl4.EOF%>

<%mdbor1.Close%>
<%mdbor5.Close%>
<%mdbor4.Close%>
<%mdbog.CommandText="SELECT DISTINCT RusName,ProjName,PC,RenovBlock FROM inpl WHERE ne='" & mdborl4("ne") & "' AND Yearr >= '" & ya & "' AND Pid = '" & Mdborl4("Pid") & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND IDentifier='C' "%>
<%mdborg.Open mdbog%>
<%mdbo1.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi,ProjCode FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise = '" & mdborl3("Enterprise") & "' AND m.Pid='" & mdborl4("Pid") & "' AND m.Yearr='" & ya & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C') GROUP BY ProjCode,m.enterprise ORDER BY ProjCode"%>
<%mdbor1.Open mdbo1%>
<%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND Identifier='P' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(SummYe),0) as SYT, ISNULL(SUM(PastSum),0) as PASU FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND Identifier='P' AND Yearr='" & ya & "'"%>
<%mdbor4.Open mdbo4%>

<tr bgcolor="FFFFFF">
<td>
<Font Color="000000" Face="Verdana" Size="2">
&nbsp;<Font color="0000FF"><%=mdborl4("Pid")%></Font>&nbsp;|&nbsp;<%=mdborg("PC")%>.
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborg("ProjName")%></Font>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><%=mdborg("RusName")%></Font>
</Font>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
 <%Else%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFFF; border-width:0">
<%End If%>
</td>
<td>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>
 
<%Else%>
  <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFFF; border-width:0">
<%End If%>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 

<%If mdbor4.EOF=True Then%>
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=True Then%>
<%sim=mdbor4("PASU")%>
<%Else%>
<%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4("PASU"))%>
<%End If%>
<%End If%>
<%If mdborg("RenovBlock")=0 AND MID(mdborg("PC"),10,2)<>"00" Then%>
<%ent2(mdborl3("Enterprise"),1)=ent2(mdborl3("Enterprise"),1)+CDbl(sim)%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ay" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ay" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ay" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
<%End If%>
<%End If%>
</Font>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor4.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("SYT")%>
<%End If%>

<%If mdborg("RenovBlock")=0 AND MID(mdborg("PC"),10,2)<>"00" Then%>
<%ent2(mdborl3("Enterprise"),2)=ent2(mdborl3("Enterprise"),2)+CDbl(sim)%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
     <%=Request.Form(a0)%>
  </Font>

<%Else%>
  <%a0="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
<%End If%>
<%End If%> 
 </Font></td>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor1.EOF=True Then%>
<%sim=0%>
<%Else%>
<%sim=mdbor4("PASU")%>
<%End If%>
<%If mdborg("RenovBlock")=0 AND MID(mdborg("PC"),10,2)<>"00" Then%>
<%ent2(mdborl3("Enterprise"),3)=ent2(mdborl3("Enterprise"),3)+CDbl(sim)%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFFF; border-width:0">
<%End If%>
<%End If%>

 </Font></td>
<%jo=4%>
<%For ja=CDbl(ya-5) to CDbl(ya-1)%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise='" & mdborl3("Enterprise") & "' AND m.Pid='" & mdborl4("Pid") & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor2.BOF=True OR mdbor2("Summi") & "e" ="e" then%>
     <%sim=0%>
<%Else%>
     <%sim=mdbor2("Summi")%>
<%End If%>
<%If mdborg("RenovBlock")=0 AND MID(mdborg("PC"),10,2)<>"00" Then%>
<%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
<%End If%>
<%Jo=jo+1%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
  <Font Color="000000" Face="Verdana" size="2"> 
    <%=Request.Form(a0)%>
   </Font>

<%Else%>
  <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
<%If Request.Form(a0)="" Then%>
<input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFFF; border-width:0">
<%Else%>
<input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; color: #000000; background-color: #FFFFFF; border-width:0">
  <%End If%>
<%End If%>

<%mdbor2.Close%>
 </Font></td>
<%Next%>
<%For ja=CDbl(ya) to CDbl(ya+4)%>
<td>
<Font Color="000000" Face="Verdana" Size="2"> 
<%If mdbor5.EOF=True then%>
     <%sim=0%>
<%Else%>
<%If CDBL(mdbor5("Yearr"))=Ja then%>
     <%sim=mdbor5("SY")%>
<%mdbor5.MoveNext%>
<%Else%>
     <%sim=0%>
<%End If%>
<%End If%>
<%If mdborg("RenovBlock")=0 AND MID(mdborg("PC"),10,2)<>"00" Then%>
<%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
<%End If%>
<%Jo=jo+1%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
   <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
   <Font Color="000000" Face="Verdana" size="2"> 
      <%=Request.Form(a0)%>
   </Font>

<%Else%>
   <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
   <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
   <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
   <%End If%>
<%End If%>
 </Font>
</td>

<%Next%>
</tr>

<%mdborl4.Movenext%>
<%mdborg.Close%>
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
<%mdbor4.Close%>
<%Dim koku(13)%><%Dim kok2(13)%>
<tr>
<td colspan="18">
<Font Color="000000" Face="Verdana" Size="2"><b>
Kokku ettev&otildette kaupa
</b></Font>
</td>
</tr>
<%mdbo4.CommandText="SELECT * FROM Enterprise"%>
<%mdbor4.Open mdbo4%>
<%Do until mdbor4.EOF%>
<tr bgcolor="AAFFFF">
<td>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b><%=mdbor4("EDescr")%></b></font>
</td>
<%For nuu=3 to 4%>
 <td><Font Color="000000" Face="Verdana" Size="2"></font></td>
<%Next%>
<%For nuu=6 to 19%>
 <td><Font Color="000000" Face="Verdana" Size="2"><b><%=entt(Mdbor4("Enterprise"),nuu-6)%></b></font></td>
<%koku(nuu-6)=koku(nuu-6)+entt(Mdbor4("Enterprise"),nuu-6)%>
<%Next%>
</tr>
<%mdbor4.Movenext%>
<%Loop%>

<tr bgcolor="FFFFFF">
<td>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b>Kokku</b></font>
</td>
<%For nuu=3 to 4%>
 <td><Font Color="000000" Face="Verdana" Size="2"></font></td>
<%Next%>
<%For nuu=6 to 19%>
 <td><Font Color="000000" Face="Verdana" Size="2"><b><%=koku(nuu-6)%></b></font></td>
<%Next%>
</tr>
<tr>
<td colspan="18">
<Font Color="000000" Face="Verdana" Size="2"><b>
Kokku ettev&otildette kaupa, v&auml;lja arvatud plokkide renoveerimine
</b></Font>
</td>
</tr>
<%mdbor4.MoveFirst%>
<%Do until mdbor4.EOF%>
<tr bgcolor="AAFFFF">
<td>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b><%=mdbor4("EDescr")%></b></font>
</td>
<%For nuu=3 to 4%>
 <td><Font Color="000000" Face="Verdana" Size="2"></font></td>
<%Next%>
<%For nuu=6 to 19%>
 <td><Font Color="000000" Face="Verdana" Size="2"><b><%=ent2(Mdbor4("Enterprise"),nuu-6)%></b></font></td>
<%kok2(nuu-6)=kok2(nuu-6)+ent2(Mdbor4("Enterprise"),nuu-6)%>
<%Next%>
</tr>
<%mdbor4.Movenext%>
<%Loop%>

<tr bgcolor="FFFFFF">
<td>
</td>
<td>
<Font Color="000000" Face="Verdana" Size="2"><b>Kokku</b></font>
</td>
<%For nuu=3 to 4%>
 <td><Font Color="000000" Face="Verdana" Size="2"></font></td>
<%Next%>
<%For nuu=6 to 19%>
 <td><Font Color="000000" Face="Verdana" Size="2"><b><%=kok2(nuu-6)%></b></font></td>
<%Next%>
</tr>
</Form>
</table>
</body>
</html>