<html>
<Head>
<link rel="stylesheet" href="STYLE.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>
</title>
</Head>
<body class="Report">

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
<%zzz=ya%>
<%zzz=zzz-1%>
<%zzz2=zzz+2%>
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
<%set mdbo1 = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdbo1.ActiveConnection = mdbo%>
<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>
<%set mdbo2u = Server.CreateObject("ADODB.Command")%>
<%set mdbor2u = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2u.ActiveConnection = mdbo%>
<%set mdbo3 = Server.CreateObject("ADODB.Command")%>
<%set mdbor3 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo3.ActiveConnection = mdbo%>
<%set mdbo3u = Server.CreateObject("ADODB.Command")%>
<%set mdbor3u = Server.CreateObject("ADODB.Recordset")%>
<%mdbo3u.ActiveConnection = mdbo%>
<%set mdbo4 = Server.CreateObject("ADODB.Command")%>
<%set mdbor4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4.ActiveConnection = mdbo%>

<%set mdbo5 = Server.CreateObject("ADODB.Command")%>
<%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5.ActiveConnection = mdbo%>
<%set mdbo6 = Server.CreateObject("ADODB.Command")%>
<%set mdbor6 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo6.ActiveConnection = mdbo%>
<%set mdbo6a = Server.CreateObject("ADODB.Command")%>
<%set mdbor6a = Server.CreateObject("ADODB.Recordset")%>
<%mdbo6a.ActiveConnection = mdbo%>
<%set mdbo7 = Server.CreateObject("ADODB.Command")%>
<%set mdbor7 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo7.ActiveConnection = mdbo%>

<table border=1>
<tr>
 <th>Projekti nr</th>
 <th>nr</a>
 <th>Projekti nimetus</th>
 <th><%=ya%>&nbspm.a&nbspkava</th>
 <th><%=ya%>&nbspm.a&nbsptegelikkult tehtud t&ouml;&ouml;d</th>
 <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.04.<%=ya%></th>
 <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.12.<%=ya%></th>
 <th><%=ya%>&nbspm.a&nbspkokku investeeritud</th>
 <th><%=ya%>&nbspm.a&nbspkokku k&auml;iku antud</th>
 <th>Demontaa&#382;</th>
 <th>L&otilde;petamata ehitus seisuga 01.04.<%=ya%></th>
 <th>L&otilde;petamata ehitus seisuga 01.12.<%=ya%></th>
</tr>
<tr class="Repnum">
<%For nuu=1 to 12%>
 <td><%=nuu%></td>
<%Next%>
</tr>


<%aa=0%><%ab=0%>
<%ac=0%>
<%mdbol1.CommandText="SELECT DISTINCT Pid,ProjCode, OracleCode,PC, PRojName FROM inpl WHERE IDentifier='C' AND Yearr='" & ya & "' AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
<%mdborl1.Open mdbol1%>
<%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND m.OracleCode<>'EJB206' AND (m.IDentifier = 'C')"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND m.Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND m.OracleCode<>'EJB206' AND (m.IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(GP.DEBET,0))/1000),0),0)-ISNULL(ROUND((SUM(ISNULL(GP.CREDIT,0))/1000),0),0) AS summi FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND m.OracleCode='EJB206' AND (m.IDentifier = 'C')"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summc FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND m.OracleCode<>'EJB206' AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND (m.IDentifier = 'C')"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summd FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND (m.IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(SummaPlan,0)),0) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summy FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (((SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (m.IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT DISTINCT ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM,ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(GP.CREDIT,0))/1000,0),0) AS summym FROM dbo.Main m INNER JOIN dbo.glav_project GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND m.RenovBlock=0 AND (SUBSTRING(GP.MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(GP.MES, 2, 2) >= 04) AND (SUBSTRING(GP.MES, 2, 2) < 07) AND (m.IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>
<tr class="Whitetr">
<td>
EJB206

</td>
<td>
01.01.02.00.

</td>
<td>
IVESTEERINGUDKOKKUvar

</td>
<td>

<%=mdbor5("SP")%>

</td>
<td>

<%If mdbor2u.EOF=True THEN%>
<%a1=0%>
<%ELSe%>
<%a1=mdbor2u("Summi")%>
<%End If%>
<%If mdbor2.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor2("Summi")%>
<%End If%>
<%=CDbl(a2)+CDBL(a1)%>


</td>
<td>

<%If mdbor6.EOF=True THEN%>
<%a6=0%>
<%ELSe%>
<%a6=mdbor6("Summy")%>
<%End If%>
<%If mdbor6a.EOF=True THEN%>
<%a6a=0%>
<%ELSe%>
<%a6a=mdbor6a("EM")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1b"%>
  
    <%=Request.Form(a0)%>
  
 <%Else%>
  <%a0="a1b"%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a1b"%>" style="font-family: Verdana; color: #FFFFFF; font-weight:700; background-color: #FFFFFF; border-width:0">
<%End If%>

</td>
<td>


<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=mdbor7("Summym")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1a"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a1a"%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a1a"%>" style="font-family: Verdana; color: #FFFFFF; font-weight:700; background-color: #FFFFFF; border-width:0">
<%End If%>

</td>
<td>

<%If mdbor2u.EOF=True THEN%>
<%a1=0%>
<%ELSe%>
<%a1=mdbor2u("Summi")%>
<%End If%>

<%If mdbor2.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor2("Summi")%>
<%End If%>

<%If mdbor6.EOF=True THEN%>
<%a6=0%>
<%ELSe%>
<%a6=mdbor6("Summy")%>
<%End If%>

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=mdbor7("Summym")%>
<%End If%>

<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("EM")%>
<%End If%>

<%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>

</td>
<td>

<%If mdbor3.EOF=FALSE Then%>
<%a3=mdbor3("Summc")%>
<%Else%>
<%a3=0%>
<%End If%>

<%If mdbor3u.EOF=FALSE Then%>
<%a4=mdbor3u("Summc")%>
<%Else%>
<%a4=0%>
<%End If%>

<%If mdbor2u.EOF=FALSE Then%>
<%a2=mdbor2u("Summi")%>
<%Else%>
<%a2=0%>
<%End If%>
<%=cdbl(a3)+cdbl(a4)%>

</td>
<td>

<%If mdbor4.EOF=False Then%>
<%=mdbor4("Summd")%>
<%Else%>
0
<%End If%>

</td>
<td>

<%If mdbor6a.EOF=True THEN%>
<%a6=0%>
<%ELSe%>
<%a6=mdbor6a("SD")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1c"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a1c"%>
  <input type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a1c"%>" style="font-family: Verdana; color: #FFFFFF; font-weight:700; background-color: #FFFFFF; border-width:0">
<%End If%>

</td>
<td>

<%If mdbor6A.EOF=true then%>
<%a6=0%>
<%ELse%>
<%a6=mdbor6a("SD")%>
<%End iF%>

<%If mdbor3.EOF=true Then%>
<%a3=0%>
<%ELse%>
<%a3=mdbor3("Summc")%>
<%End iF%>

<%If mdbor3u.EOF=true Then%>
<%a9=0%>
<%ELse%>
<%a9=mdbor3u("Summc")%>
<%End iF%>

<%If mdbor2u.EOF=True THEN%>
<%a1=0%>
<%ELSe%>
<%a1=mdbor2u("Summi")%>
<%End If%>

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

<%If mdbor2.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor2("Summi")%>
<%End If%>
<%a0="a1c"%>
<%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor6a.Close%>
<%mdbor7.Close%>
</tr>
</table>
</body>
</html>