<Html>
 <Head>
  <meta http-equiv="Content-Type" content="text/html; charset=ibm852">
  <title>Invest-IT!on: KULUDE DIAGRAMMID</title>
 </Head>
 <body>
  <%b= Server.MapPath("\inv")%>
<%set mdbo =  Server.CreateObject("ADODB.Connection")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")
  set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%>
<%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Trusted_Connection=yes;Database=invest;UID=;PWD=;"%>
<%mdbo.Open ConnectionString%>
<img border="0" src="icons/graph.ico" Style=float:Left><p align="center">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<a href="#null"><font face="Verdana" Size="5" color="000099"><b><u>KULUDE DIAGRAMMID</font></u></b></a></p>
<p>
<hr color="000000">
<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>
<%set mdbo0 = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdbo0.ActiveConnection = mdbo%>
<%ya=Year(Date())%>
<%mo=Month(Date())%>
<%da=Day(Date())%>
<%zz=mo-04%>
<%If zz>0 Then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=ya-1%>
<%End If%>
<table border="1" bordercolor="000000">
<tr>
<th>
<Font Size="3" face="Verdana">Plaanilised Summad</font>
</th>
<th>
<Font Size="3" face="Verdana">Tegelikud summad</font>
</th>
</tr>
<tr>
<td>
<%mdbo0.CommandText="SELECT SUM(dbo.main.SummYe) AS Kulud, dbo.Enterprise.EDescr FROM dbo.Main INNER JOIN dbo.Enterprise ON dbo.Main.Enterprise = dbo.Enterprise.Enterprise WHERE (dbo.Main.IDentifier = 'P') AND Yearr='" & ya & "' GROUP BY dbo.Main.Enterprise, dbo.Enterprise.EDescr"%>
<%mdbor.Open mdbo0%>
<%Dim ar(30)%>
<%ar(1)="FF0000"%>
<%ar(2)="FFFF00"%>
<%ar(3)="00FF00"%>
<%ar(4)="00FFFF"%>
<%ar(5)="0000FF"%>
<%ar(6)="FF00FF"%>
<%ar(7)="00FFC0"%>
<%ar(8)="C000FF"%>
<%ar(9)="FF00C0"%>
<%ar(10)="FFC000"%>
<%ar(11)="C0FF00"%>
<%ar(12)="00C0FF"%>
<Font Size="3" face="Verdana">Kulude jaotamine ettev&otilde;tetele</font>
<%Do Until mdbor.EOF%>
<%tot=tot + CDBL(Mdbor("kulud"))%>
<%mdbor.Movenext%>
<%Loop%>
<%i=1%>
<%mdbor.movefirst%>
<%Do Until mdbor.EOF%>
<%num=CDbl(Mdbor("kulud"))%>
<%If tot<=0 Then%>
<%per=0%>
<%Else%>
<%per=ROUND(num/tot*100)%>
<%End If%>
<table>
<tr>
<td class="altActive" title="<%=mdbor("EDEscr")%>">
<%k=0%>
<%If Per=0 Then%>
<Font color="<%=ar(i)%>" size="2">
*
</font>
<%End If%>

<%for k=1 to per%>
<%sss=sss+chr(219)%>
<%Next%>
<Font color="<%=ar(i)%>" size="1">
<%=sss%>
</font>
<%sss=""%>
</td>
<td class="altActive" title="<%=mdbor("EDEscr")%>">
&nbsp&nbsp<%=per%>%
</td>
</tr>
</table>
<%i=i+1%>
<%mdbor.Movenext%>
<%Loop%>

<%mdbor.Close%>
<hr color="0000F0">
<%tot=0%>
<%per=0%>

<%mdbo2.CommandText="SELECT ProjCode, ProjName from CODES WHERE SUBSTRING(ProjCode,4,2)='00'"%>
<%mdbor2.Open mdbo2%>
<%mdbo0.CommandText="SELECT DISTINCT SUM(SummaPlan) AS SP, SUM(SummaContract) AS SC, SUM(SummaFact) AS SF, be FROM dbo.Delta WHERE yearr='" & ya & "' GROUP BY be"%>
<%mdbor.Open mdbo0%>
<Font Size="3" face="Verdana">Kulude jaotamine projekti gruppidele</font>
<%i=1%>
<%Do Until mdbor2.EOF%>
<%If mdbor.EOF=True then%>
<%tot=tot + 0%>
<%ELse%>
<%tot=tot + CDbl(Mdbor("SP"))%>
<%mdbor.Movenext%>
<%End If%>
<%mdbor2.Movenext%>
<%Loop%>
<%mdbor.movefirst%><%mdbor2.movefirst%>
<%Do Until mdbor2.EOF%>
<%If mdbor.EOF=True then%>
<%num=0%>
<%ELse%>
<%num=Cdbl(Mdbor("SP"))%>
<%mdbor.Movenext%>
<%End If%>


<%If tot<=0 Then%>
<%per=0%>
<%Else%>
<%per=ROUND(num/tot*100)%>
<%End If%>
<table>
<tr>
<td class="altActive" title="<%=LEFT(mdbor2("ProjCode"),2) & " " & mdbor2("ProjName")%>">
<%k=0%>
<%If Per=0 Then%>
<Font color="<%=ar(i)%>" size="2">
*
</font>
<%End If%>
<%for k=1 to per%>
<%sss=sss+chr(219)%>
<%Next%>
<Font color="<%=ar(i)%>" size="1">
<%=sss%>
</font>
<%sss=""%>
</td>
<td class="altActive" title="<%=LEFT(mdbor2("ProjCode"),2) & " " & mdbor2("ProjName")%>">
&nbsp&nbsp<%=per%>%
</td>
</tr>
</table>
<%i=i+1%>
<%mdbor2.Movenext%>
<%Loop%>
<%mdbor.Close%>
<hr color="0000F0">
<%tot=0%>
<%per=0%>

<%mdbo0.CommandText="SELECT SUM(SummYe) AS SM, Yearr FROM dbo.Main WHERE (IDentifier = 'P') GROUP BY Yearr ORDER BY YEarr"%>
<%mdbor.Open mdbo0%>
<Font Size="3" face="Verdana">Kulude jaotus tegevusaastate l&otilde;ikes </font>
<%i=1%>
<%Do Until mdbor.EOF%>
<%tot=tot + Cdbl(Mdbor("SM"))%>
<%mdbor.Movenext%>
<%Loop%>
<%mdbor.movefirst%>
<%Do Until mdbor.EOF%>
<%num=Cdbl(Mdbor("SM"))%>
<%If tot<=0 Then%>
<%per=0%>
<%Else%>
<%per=ROUND(num/tot*100)%>
<%End If%>
<table>
<tr>
<td class="altActive" title="<%=mdbor("Yearr")%>">
<%k=0%>
<%If Per=0 Then%>
<Font color="<%=ar(i)%>" size="2">
*
</font>
<%End If%>
<%for k=1 to per%>
<%sss=sss+chr(219)%>
<%Next%>
<Font color="<%=ar(i)%>" size="1">
<%=sss%>
</font>
<%sss=""%>
</td>
<td class="altActive" title="<%=mdbor("Yearr")%>">
&nbsp&nbsp<%=per%>%
</td>
</tr>
</table>
<%i=i+1%>
<%mdbor.Movenext%>
<%Loop%>
<%mdbor.Close%>
<%mdbor2.Close%>
</td>

<td>
<%tot=0%>
<%per=0%>

<%mdbo0.CommandText="SELECT SUM(dbo.main.SummYe) AS Kulud, dbo.Enterprise.EDescr FROM dbo.Main INNER JOIN dbo.Enterprise ON dbo.Main.Enterprise = dbo.Enterprise.Enterprise WHERE (dbo.Main.IDentifier = 'F') AND Yearr='" & ya & "' GROUP BY dbo.Main.Enterprise, dbo.Enterprise.EDescr"%>
<%mdbor.Open mdbo0%>
<Font Size="3" face="Verdana">Kulude jaotamine ettev&otilde;tetele</font>
<%Do Until mdbor.EOF%>
<%tot=tot + CDbl(Mdbor("kulud"))%>
<%mdbor.Movenext%>
<%Loop%>
<%i=1%>
<%mdbor.movefirst%>
<%Do Until mdbor.EOF%>
<%num=Cdbl(Mdbor("kulud"))%>
<%If tot<=0 Then%>
<%per=0%>
<%Else%>
<%per=ROUND(num/tot*100)%>
<%End If%>
<table>
<tr>
<td class="altActive" title="<%=mdbor("EDEscr")%>">
<%k=0%>
<%If Per=0 Then%>
<Font color="<%=ar(i)%>" size="1">
*
</font>
<%End If%>
<%for k=1 to per%>
<%sss=sss+chr(219)%>
<%Next%>
<Font color="<%=ar(i)%>" size="1">
<%=sss%>
</font>
<%sss=""%>
</td>
<td class="altActive" title="<%=mdbor("EDEscr")%>">
&nbsp&nbsp<%=per%>%
</td>
</tr>
</table>
<%i=i+1%>
<%mdbor.Movenext%>
<%Loop%>

<%mdbor.Close%>
<hr color="0000F0">
<%tot=0%>
<%per=0%>
<%mdbo2.CommandText="SELECT ProjCode,ProjName from CODES WHERE SUBSTRING(ProjCode,4,2)='00'"%>
<%mdbor2.Open mdbo2%>
<%mdbo0.CommandText="SELECT DISTINCT SUM(SummaPlan) AS SP, SUM(SummaContract) AS SC, SUM(SummaFact) AS SF, be FROM dbo.Delta WHERE yearr='" & ya & "' GROUP BY be"%>
<%mdbor.Open mdbo0%>
<Font Size="3" face="Verdana">Kulude jaotamine projekti gruppidele</font>
<%i=1%>
<%Do Until mdbor2.EOF%>
<%If mdbor.EOF=True then%>
<%tot=tot + 0%>
<%ELse%>
<%tot=tot + CDbl(Mdbor("SF"))%>
<%mdbor.Movenext%>
<%End If%>
<%mdbor2.Movenext%>
<%Loop%>
<%mdbor2.movefirst%><%mdbor.movefirst%>
<%Do Until mdbor2.EOF%>
<%If mdbor.EOF=True then%>
<%num=0%>
<%ELse%>
<%num=Cdbl(Mdbor("SF"))%>
<%mdbor.Movenext%>
<%End If%>
<%If tot<=0 Then%>
<%per=0%>
<%Else%>
<%per=ROUND(num/tot*100)%>
<%End If%>
<table>
<tr>
<td class="altActive" title="<%=LEFT(mdbor2("ProjCode"),2) & " " & mdbor2("ProjName")%>">
<%k=0%>
<%If Per=0 Then%>
<Font color="<%=ar(i)%>" size="2">
*
</font>
<%End If%>
<%for k=1 to per%>
<%sss=sss+chr(219)%>
<%Next%>
<Font color="<%=ar(i)%>" size="1">
<%=sss%>
</font>
<%sss=""%>
</td>
<td class="altActive" title="<%=LEFT(mdbor2("ProjCode"),2) & " " & mdbor2("ProjName")%>">
&nbsp&nbsp<%=per%>%
</td>
</tr>
</table>
<%i=i+1%>
<%mdbor2.Movenext%>
<%Loop%>
<%mdbor.Close%>
<hr color="0000F0">
<%tot=0%>
<%per=0%>

<%mdbo0.CommandText="SELECT SUM(SummYe) AS SM, Yearr FROM dbo.Main WHERE (IDentifier = 'F') GROUP BY Yearr ORDER BY YEarr"%>
<%mdbor.Open mdbo0%>
<Font Size="3" face="Verdana">Kulude jaotus tegevusaastate l&otilde;ikes </font>
<%i=1%>
<%Do Until mdbor.EOF%>
<%tot=tot + Cdbl(Mdbor("SM"))%>
<%mdbor.Movenext%>
<%Loop%>
<%mdbor.movefirst%>
<%Do Until mdbor.EOF%>
<%num=CDbl(Mdbor("SM"))%>
<%If tot<=0 Then%>
<%per=0%>
<%Else%>
<%per=ROUND(num/tot*100)%>
<%End If%>
<table>
<tr>
<td class="altActive" title="<%=mdbor("Yearr")%>">
<%k=0%>
<%If Per=0 Then%>
<Font color="<%=ar(i)%>" size="2">
*
</font>
<%End If%>
<%for k=1 to per%>
<%sss=sss+chr(219)%>
<%Next%>
<Font color="<%=ar(i)%>" size="1">
<%=sss%>
</font>
<%sss=""%>
</td>
<td class="altActive" title="<%=mdbor("Yearr")%>">
&nbsp&nbsp<%=per%>%
</td>
</tr>
</table>
<%i=i+1%>
<%mdbor.Movenext%>
<%Loop%>
</td>
</tr>
</body>
</html>