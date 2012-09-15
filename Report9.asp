<html>
<Head>
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
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title>
</Head>
<body class="Rrport">

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
<img border="0" src="icons/report.ico" Style=float:Left><p align="center"><a href="Main.asp" target="_top" class="headlink"><%=ya%> majandusaasta investeeringute kava 9 kuu l&otilde;ikes</a></p><p>
<%zzz=ya%>
<%zzz=zzz-1%>
<%zzz2=zzz+2%>
<Form Method="POST" Action="Report9.asp?ye=<%=ya%>">
<Input type="Submit" name="btn" size="10" Value="Kopeerimiseks" class="button">
<Input type="Submit" name="btn" size="10" Value="Parandamiseks" class="button">

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
 <th>nr</th>
 <th>Projekti nimetus</th>
 <th><%=ya%>&nbspa&nbspkava</th>
 <th><%=ya%>&nbspa&nbsptegelikkult tehtud t&ouml;&ouml;d</th>
 <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.04.<%=ya%></th>
 <th>Ettemaksed ja p&otilde;hivara laos seisuga 01.12.<%=ya%></th>
 <th><%=ya%>&nbspa&nbspkokku investeeritud</th>
 <th><%=ya%>&nbspa&nbspkokku k&auml;iku antud</th>
 <th>Demontaa&#382;</th>
 <th>L&otilde;petamata ehitus seisuga 01.04.<%=ya%></th>
 <th>L&otilde;petamata ehitus seisuga 01.12.<%=ya%></th>
</tr>
<tr class="refnum">
<%For nuu=1 to 12%>
 <td><%=nuu%></td>
<%Next%>
</tr>


<%aa=0%><%ab=0%>
<%ac=0%>
<%mdbol1.CommandText="SELECT DISTINCT Pid,ProjCode, OracleCode,PC, PRojName FROM inpl WHERE IDentifier='C' AND Yearr='" & ya & "' AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
<%mdborl1.Open mdbol1%>
<%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Description<>'MAAGAAS' AND (IDentifier = 'C')"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summd FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(SummaPlan) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>
<tr class="boldProjGrup">
<td></td>
<td></td>
<td>IVESTEERINGUD KOKKU v&auml;lja arvatud plokkide renoveerimine</td>
<td><%=mdbor5("SP")%></td>
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
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a1b"%>" class="boldProjGrup">
<%End If%>

</td>
<td>


<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a1a"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a1a"%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a1a"%>" class="boldProjGrup">
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
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=0%>
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
  <input type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a1c"%>" class="boldProjGrup">
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
<%mdbo3u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and  Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'MAAGAAS' AND SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And Yearr='" & ya & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (konto NOT BETWEEN '18410' AND '18433') /*AND Project<>'EJB206' */AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summd FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(ISNULL(SummaPlan,0)) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00'"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00'"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>
<tr class="boldProjGrup">
<td>


</td>
<td>


</td>
<td>
IVESTEERINGUD KOKKU koos plokkide renoveerimisega

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
  <%a0="ab"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ab"%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="ab"%>" class="boldProjGrup">
<%End If%>

</td>
<td>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="aa"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="aa"%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa"%>" class="boldProjGrup">
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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>


<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("EM")%>
<%End If%>

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

<%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>

</td>
<td>

<%If mdbor3.EOF=FALSE Then%>
<%a3=mdbor3("Summc")%>
<%Else%>
<%a3=0%>
<%End If%>
<%If mdbor3U.EOF=FALSE Then%>
<%a4=mdbor3U("Summc")%>
<%Else%>
<%a4=0%>
<%End If%>
<%If mdbor2u.EOF=FALSE Then%>
<%a2=mdbor2u("Summi")%>
<%Else%>
<%a2=0%>
<%End If%>
<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%IF CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))=0 Then%>
<%a7=mdbor7("Summym")%>
<%ELSE%>
<%a7=0%>
<%End If%>
<%End If%>
<%=cdbl(a3)+cdbl(a4)+cdbl(a7)%>

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
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("SD")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac"%>
  
    <%=Request.Form(a0)%>
  
  <%Else%>
  <%a0="ac"%>
  <input type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac"%>" class="boldProjGrup">
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


<%Do until mdborl1.EOF%>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor7.Close%>
<%mdbor6a.Close%>
<%mdbo2.CommandText="SELECT ROUND((SUM(ISNULL(DEBET,0))/1000) ,0) AS summi, SUBSTRING(ProjCode, 1, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor2.Open mdbo2%>
<%mdbo3u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summc, SUBSTRING(ProjCode, 1, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(Project,4,3)='999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2u.CommandText="SELECT ROUND((SUM(ISNULL(DEBET,0))/1000) ,0) AS summi, SUBSTRING(ProjCode, 1, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' and SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(CREDIT,0))/1000,0) AS summc, SUBSTRING(ProjCode, 1, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE  description<>'MAAGAAS' AND SUBSTRING(Project,4,3)<>'999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And Yearr='" & ya & "'  AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "'/* AND Project<>'EJB206' */AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summd, SUBSTRING(ProjCode, 1, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(ISNULL(SummaPlan,0)) AS SP, SUM(ISNULL(SummaContract,0)) AS SC, be FROM dbo.Delta WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,2)='" & MId(mdborl1("PC"),1,2) & "' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2)"%>
<%mdbor7.Open mdbo7%>
<tr class="ProjGrup">
<td>

<%If mdborl1("OracleCode")="N/A" Then%>
&nbsp
<%Else%>
<%=mdborl1("OracleCode")%>
<%End If%>

</td>
<td>
<%a=MID(mdborl1("PC"),1,3)%>
<%=REPLACE(a, "0", "")%>
</td>
<td>
<%=mdborl1("ProjName")%>

</td>
<td>

<%If mdbor5("be")=MId(mdborl1("PC"),1,2) Then%>
<%=mdbor5("SP")%>
<%Else%>
0
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
  <%a0="ab" & mdborl1("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ab" & mdborl1("Pid")%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="ab" & mdborl1("Pid")%>" class="ProjGrup">
<%End If%>

</td>
<td>
<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="aa" & mdborl1("Pid")%>
  
    <%=Request.Form(a0)%>
  
  <%Else%>
<%a0="aa" & mdborl1("Pid")%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl1("Pid")%>" class="ProjGrup">
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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>


<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("EM")%>
<%End If%>

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

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
<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%IF CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))=0 Then%>
<%a7=mdbor7("Summym")%>
<%ELSE%>
<%a7=0%>
<%End If%>
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
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("SD")%>
<%End If%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl1("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ac" & mdborl1("Pid")%>
  <input type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl1("Pid")%>" class="ProjGrup">
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

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

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
<%a0="a1c"%>
<%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>

<%mdbol2.CommandText="SELECT DISTINCT Pid,ProjCode,ProjName,OracleCode,PC FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND  SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY ProjCode"%>
<%mdborl2.Open mdbol2%>

<%Do until mdborl2.EOF%>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor7.Close%>
<%mdbor6a.Close%>
<%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND  Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor2.Open mdbo2%>
<%mdbo3u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summc, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(CREDIT,0))/1000,0) AS summc, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE  description<>'MAAGAAS' AND SUBSTRING(PROJECT,4,3)<>'999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  And Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "')/* AND Project<>'EJB206' */AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summd, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND  (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(ISNULL(SummaPlan,0)) AS SP, be, mi FROM dbo.Delta WHERE yearr='" & ya & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND enn<>'00' AND be='" & Mid(mdborl2("PC"),1,2) & "' GROUP BY be, mi"%>
<%mdbor5.Open mdbo5%>
<%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND  (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')   and Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')  AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND yearr='" & ya & "' AND enn<>'00' GROUP BY be, mi"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2)"%>
<%mdbor7.Open mdbo7%>
<tr class="ProjGrup">
<td>

<%If mdborl2("OracleCode")="N/A" Then%>
&nbsp
<%Else%>
<%=mdborl2("OracleCode")%>
<%End If%>

</td>
<td>
<%a=MID(mdborl2("PC"),1,6)%>
<%=REPLACE(a, "0", "")%>
</td>
<td>
<%=mdborl2("ProjName")%>

</td>
<td>

<%If mdbor5("mi")=MID(mdborl2("PC"),4,2) Then%>
<%=mdbor5("SP")%>
<%Else%>
0
<%End If%>

</td>
<td>

<%a0="aa" & mdborl2("Pid")%>
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
  <%a0="ab" & mdborl2("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ab" & mdborl2("Pid")%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="ab" & mdborl2("Pid")%>" class="ProjGrup">
<%End If%>

</td>
<td>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=mdbor7("Summym")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="aa" & mdborl2("Pid")%>
  
    <%=Request.Form(a0)%>
  
 <%Else%>
  <%a0="aa" & mdborl2("Pid")%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl2("Pid")%>" class="ProjGrup">
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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>


<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("EM")%>
<%End If%>

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%IF CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))=0 Then%>
<%a7=mdbor7("Summym")%>
<%ELSE%>
<%a7=0%>
<%End If%>
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
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("SD")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl2("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ac" & mdborl2("Pid")%>
  <input type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl2("Pid")%>" class="ProjGrup">
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

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

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
<%a0="a1c"%>

<%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>
<%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "'"%>
<%mdborl3.Open mdbol3%>


<%Do until mdborl3.EOF%>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor7.Close%>
<%mdbor6a.Close%>
<%mdbo2.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND  SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor2.Open mdbo2%>
<%mdbo3u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summc, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(Project,4,3)='999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')  and  Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2u.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summi, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2) AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND  (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700')   and Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')  AND (konto NOT BETWEEN '18410' AND '18433') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(CREDIT,0))/1000,0) AS summc, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2), Enterprise AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE  description<>'MAAGAAS' AND SUBSTRING(Project,4,3)<>'999' and (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') and Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "'/* AND Project<>'EJB206' */AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summd, SUBSTRING(ProjCode, 1, 2), SUBSTRING(ProjCode, 4, 2), Enterprise AS be FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND  (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(SummaPlan) AS SP, be, mi, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND enn<>'00' AND be='" & Mid(mdborl2("PC"),1,2) & "' GROUP BY be, mi,Enterprise"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "'  AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE be='" & MID(mdborl1("PC"),1,2) & "' AND mi = '" & Mid(mdborl2("PC"),4,2) & "' AND yearr='" & ya & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND enn<>'00' GROUP BY be, mi,Enterprise"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,5)='" & MId(mdborl2("PC"),1,5) & "' AND Enterprise='" & Mdborl3("Enterprise") & "'  AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY SUBSTRING(ProjCode, 1, 2),SUBSTRING(ProjCode, 4, 2), Enterprise"%>
<%mdbor7.Open mdbo7%>
<tr class="enterp">
<td>
</td>
<td>
</td>
<td>
<%=mdborl3("EDescr")%>

</td>
<td>

<%If mdbor5.EOF=False Then%>
<%=mdbor5("SP")%>
<%Else%>
0
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
  <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" class="Enterp">
<%End If%>

</td>
<td>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" class="Enterp">
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

<%If mdbor4.EOF=faLSE Then%>
<%a4=mdbor4("Summd")%>
<%Else%>
<%a4=0%>
<%End If%>

<%If mdbor6.EOF=True THEN%>
<%a6=0%>
<%ELSe%>
<%a6=mdbor6("Summy")%>
<%End If%>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%IF CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))=0 Then%>
<%a7=mdbor7("Summym")%>
<%ELSE%>
<%a7=0%>
<%End If%>
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
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("SD")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
  <input type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" class="Enterp">
<%End If%>

</td>

<td>

<%If mdbor6a.EOF=true then%>
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

<%If mdbor4.EOF=faLSE Then%>
<%a4=mdbor4("Summd")%>
<%Else%>
<%a4=0%>
<%End If%>

<%If mdbor2.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor2("Summi")%>
<%End If%>
<%a0="a1c"%>
<%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>


<%mdbo1.CommandText="SELECT * FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND SUBSTRING(PC,7,2)<>'00' AND Enterprise='" & Mdborl3("Enterprise") & "' ORDER BY PC"%>
<%mdbor.Open mdbo1%>

<%Do Until mdbor.EOF%>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor7.Close%>
<%mdbor6a.Close%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo2.CommandText="SELECT SUM(SummaPlan) AS SUMMAPLAN FROM delta WHERE LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' and Enterprise='" & Mdbor("Enterprise") & "' AND right(Projcode,2)<>'00' AND yearr='" & ya & "'"%>
<%mdbor2.Open mdbo2%>
<%ELSE%>
<%mdbo2.CommandText="SELECT SummaPlan FROM delta WHERE ROWIDC='" & mdbor("ROWID") & "'"%>
<%mdbor2.Open mdbo2%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0)/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') and Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433')  AND SUBSTRING(ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')"%>
<%mdbor2u.Open mdbo2u%>
<%ELSE%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0)/1000),0),0) AS summi FROM glav_project WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & " AND (konto NOT BETWEEN '18410' AND '18433') AND RIGHT(MES,2) >=04 AND Project='EJB206'  "%>
<%mdbor2u.Open mdbo2u%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo3.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(DEBET,0)/1000,0)),0) AS summd FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' and right(Projcode,2)<>'00' AND Enterprise='" & Mdbor("Enterprise") & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
<%mdbor3.Open mdbo3%>
<%ELSE%>
<%mdbo3.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(DEBET,0)/1000,0)),0) AS summd FROM glav_project WHERE Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & "  AND RIGHT(MES,2) >=04 AND KONTO='43350' AND SUBKONTO='4351' "%>
<%mdbor3.Open mdbo3%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(CREDIT,0)/1000,0)),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE  description<>'MAAGAAS' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' /*AND Project<>'EJB206' */AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor4.Open mdbo4%>
<%ELSE%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ROUND(ISNULL(CREDIT,0)/1000,0)),0) AS summc FROM glav_project WHERE  description<>'MAAGAAS' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & " AND (konto NOT BETWEEN '18410' AND '18433') AND RIGHT(MES,2) >=04 "%>
<%mdbor4.Open mdbo4%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo5.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0)/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') and Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND SUBSTRING(ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor5.Open mdbo5%>
<%ELSE%>
<%mdbo5.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0)/1000),0),0) AS summi FROM glav_project WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND Project='" & mdbor("OracleCode") & "' AND LEFT(MES,1)=" & Mid(ya,4,1) & " AND (konto NOT BETWEEN '18410' AND '18433') AND RIGHT(MES,2) >=04 AND Project<>'EJB206'   "%>
<%mdbor5.Open mdbo5%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND SUBSTRING(ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%ELSE%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Project='" & mdbor("OracleCode") & "' AND Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04)) AND (IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo6a.CommandText="SELECT  ISNULL(SUM(ISNULL(Ettemaks,0)),0) AS EM,ISNULL(SUM(ISNULL(Saldo,0)),0) AS SD FROM dbo.ETTE WHERE LEFT(ProjCode,8)='" & LEFT(mdbor("PC"),8) & "' and Enterprise='" & Mdbor("Enterprise") & "' AND right(Projcode,2)<>'00' AND yearr='" & ya & "'"%>
<%mdbor6a.Open mdbo6a%>
<%ELSE%>
<%mdbo6a.CommandText="SELECT  ISNULL(Ettemaks,0) AS EM,ISNULL(Saldo,0) AS SD FROM dbo.ETTE WHERE ROWID='" & mdbor("ROWID") & "'"%>
<%mdbor6a.Open mdbo6a%>
<%END iF%>
<%IF LEN(mdbor("PC"))>9 AND RIGHT(mdbor("PC"),2)="00" then%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi,ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND SUBSTRING(ProjCode,1,8)='" & MId(mdbor("PC"),1,8) & "' AND Enterprise='" & Mdbor("Enterprise") & "' and right(Projcode,2)<>'00' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>
<%ELSE%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Project='" & mdbor("OracleCode") & "' AND Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>
<%END iF%>
<tr >
<td>

<%If mdbor("OracleCode")="N/A" Then%>
&nbsp
<%Else%>
<%=mdbor("OracleCode")%>
<%End If%>

</td>
<td>
         <%if mid(mdbor("PC"),8,1)=0 and mid(mdbor("PC"),7,1)<>0 then%>
          <%a=REPLACE(MID(mdbor("PC"),1,6), "0", "") & MID(mdbor("PC"),7,2)%>
         <%Else%>
          <%a=REPLACE(mdbor("PC"), "0", "")%>
	 <%End If%>
          <%If len(a)>=7 Then%>
	  <%If Right(mdbor("PC"),2)="00" then%>
	   <%=a%>
	  <%Else%>
	   <%=a%>.
	  <%End if%>
	 <%Else%>
	  <%If Right(mdbor("PC"),2)="00" then%>
	   <%=mid(a,1,6)%>
	  <%Else%>
	   <%=mid(a,1,6)%>.
	  <%End if%>
	 <%End if%>
</td>
<td>
<%=mdbor("ProjName")%>

</td>
<td>

<%=mdbor2("SummaPlan")%>

</td>
<td>

<%If mdbor2u.EOF=True THEN%>
<%a1=0%>
<%ELSe%>
<%a1=mdbor2u("Summi")%>
<%End If%>
<%If mdbor5.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor5("Summi")%>
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
  <%a0="ab" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ab" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="ab" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>" >
<%End If%>

</td>
<td>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>" >
<%End If%>

</td>
<td>

<%If mdbor2u.EOF=True THEN%>
<%a1=0%>
<%ELSe%>
<%a1=mdbor2u("Summi")%>
<%End If%>

<%If mdbor5.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor5("Summi")%>
<%End If%>

<%If mdbor3.EOF=faLSE Then%>
<%a4=mdbor3("Summd")%>
<%Else%>
<%a4=0%>
<%End If%>

<%If mdbor6.EOF=True THEN%>
<%a6=0%>
<%ELSe%>
<%a6=mdbor6("Summy")%>
<%End If%>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>

<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("EM")%>
<%End If%>


<%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>

</td>

<td>


<%If MID(mdbor("OracleCode"),4,3)="999" Then%>
<%If mdbor5.EOF=FALSE Then%>
<%a3=mdbor5("Summi")%>
<%Else%>
<%a3=0%>
<%End If%>
<%Else%>

<%If mdbor4.EOF=FALSE Then%>
<%a3=mdbor4("Summc")%>
<%Else%>
<%a3=0%>
<%End If%>
<%End If%>

<%If mdbor2u.EOF=FALSE Then%>
<%a2=mdbor2u("Summi")%>
<%Else%>
<%a2=0%>
<%End If%>
<%=cdbl(a3)%>

</td>

<td>

<%If mdbor3.EOF=faLSE Then%>
<%=mdbor3("Summd")%>
<%Else%>
0
<%End If%>

</td>

<td>

<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("SD")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="ac" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
  <input type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="ac" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>" >
<%End If%>


</td>
<td>

<%If mdbor6A.EOF=true then%>
<%a6=0%>
<%ELse%>
<%a6=mdbor6a("SD")%>
<%End iF%>

<%If mdbor3.EOF=faLSE Then%>
<%a4=mdbor3("Summd")%>
<%Else%>
<%a4=0%>
<%End If%>

<%If MID(mdbor("OracleCode"),4,3)="999" Then%>
<%If mdbor5.EOF=FALSE Then%>
<%a3=mdbor5("Summi")%>
<%Else%>
<%a3=0%>
<%End If%>
<%Else%>
<%If mdbor4.EOF=FALSE Then%>
<%a3=mdbor4("Summc")%>
<%Else%>
<%a3=0%>
<%End If%>
<%End If%>

<%If mdbor2u.EOF=True THEN%>
<%a1=0%>
<%ELSe%>
<%a1=mdbor2u("Summi")%>
<%End If%>

<%If mdbor5.EOF=True THEN%>
<%a2=0%>
<%ELSe%>
<%a2=mdbor5("Summi")%>
<%End If%>
<%a0="a1c"%>
<%=CDbl(a6)-CDbl(a3)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>
<%mdbor.Movenext%>
<%loop%>

<%mdbor.Close%>
<%mdborl3.Movenext%>
<%loop%>

<%mdborl3.Close%>
<%mdborl2.Movenext%>
<%loop%>

<%mdborl2.Close%>
<%mdborl1.Movenext%>

<%Loop%>

<tr>
<td colspan="12">

Kokku ettev&otilde;tete kaupa

</td>
</tr>

<%mdborl1.Close%>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor7.Close%>
<%mdbor6a.Close%>
<%mdbol1.CommandText="SELECT * FROM Enterprise ORDER BY Enterprise"%>
<%mdborl1.Open mdbol1%>
<%mdbo3u.CommandText="SELECT ROUND((SUM(ISNULL(DEBET,0))/1000),0) AS summc, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ROUND((SUM(ISNULL(DEBET,0))/1000),0) AS summi, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ROUND((SUM(ISNULL(DEBET,0))/1000),0) AS summi, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ROUND(SUM(ISNULL(CREDIT,0))/1000,0) AS summc, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Description<>'MAAGAAS' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ROUND(SUM(ISNULL(DEBET,0))/1000,0) AS summd, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(SummaPlan) AS SP, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD, Enterprise FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor7.Open mdbo7%>

<%Do until mdborl1.EOF%>

<%If mdbor7.EOF=true then%>
<%a7=0%>
<%Else%>
<%IF mdbor7("Enterprise")=mdborl1("Enterprise") Then%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%mdbor7.MoveNext%>
<%ELse%>
<%a7=0%>
<%End iF%>
<%End iF%>

<%If mdbor6a.EOF=true then%>
<%a6a=0%>
<%a33=0%>
<%Else%>
<%IF mdbor6a("Enterprise")=mdborl1("Enterprise") Then%>
<%a6a=mdbor6a("EM")%>
<%a33=mdbor6a("SD")%>
<%mdbor6a.MoveNext%>
<%ELse%>
<%a6a=0%>
<%a33=0%>
<%End iF%>
<%End iF%>
<%If mdbor6.EOF=true then%>
<%a6=0%>
<%Else%>
<%IF mdbor6("Enterprise")=mdborl1("Enterprise") Then%>
<%a6=mdbor6("Summy")%>
<%mdbor6.MoveNext%>
<%ELse%>
<%a6=0%>
<%End iF%>
<%End iF%>
<%If mdbor2u.EOF=true then%>
<%a1=0%>
<%Else%>
<%IF mdbor2u("Enterprise")=mdborl1("Enterprise") Then%>
<%a1=mdbor2u("Summi")%>
<%mdbor2u.MoveNext%>
<%ELse%>
<%a1=0%>
<%End iF%>
<%End iF%> 
<%If mdbor4.EOF=true then%>
<%a4=0%>
<%Else%>
<%IF mdbor4("Enterprise")=mdborl1("Enterprise") Then%>
<%a4=mdbor4("Summd")%>
<%mdbor4.MoveNext%>
<%ELse%>
<%a4=0%>
<%End iF%>
<%End iF%>
<%If mdbor2.EOF=true then%>
<%a2=0%>
<%Else%>
<%IF mdbor2("Enterprise")=mdborl1("Enterprise") Then%>
<%a2=mdbor2("Summi")%>
<%mdbor2.MoveNext%>
<%ELse%>
<%a2=0%>
<%End iF%>
<%End iF%>
<%If mdbor3.EOF=true then%>
<%a3=0%>
<%Else%>
<%IF mdbor3("Enterprise")=mdborl1("Enterprise") Then%>
<%a3=mdbor3("Summc")%>
<%mdbor3.MoveNext%>
<%ELse%>
<%a3=0%>
<%End iF%>
<%End iF%>
<%If mdbor3u.EOF=true then%>
<%a9=0%>
<%Else%>
<%IF mdbor3u("Enterprise")=mdborl1("Enterprise") Then%>
<%a9=mdbor3u("Summc")%>
<%mdbor3u.MoveNext%>
<%ELse%>
<%a9=0%>
<%End iF%>
<%End iF%>

<tr class="boldenterp">
<td>


</td>
<td>


</td>
<td>
<%=Mdborl1("EDescr")%>

</td>
<td>

<%=mdbor5("SP")%>

</td>
<td>

<%=CDbl(a2)+CDBL(a1)%>


</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a8b" & mdborl1("Enterprise")%>
  
    <%=Request.Form(a0)%>
  
 <%Else%>
  <%a0="a8b" & mdborl1("Enterprise")%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a8b" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a8a" & mdborl1("Enterprise")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a8a" & mdborl1("Enterprise")%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a8a" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td>

<%=CDbl(a2)+CDBL(a1)-CDBL(a6a)-CDBL(a6)+CDBL(a7)%>

</td>
<td>

<%=cdbl(a3)+cdbl(a9)%>

</td>
<td>

<%=cdbl(a4)%>

</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a8c" & mdborl1("Enterprise")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a8c" & mdborl1("Enterprise")%>
  <input type="Text" value="<%=CDBL(a33)%>" size="10" name="<%="a8c" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td>

<%=CDbl(a33)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>

<%If mdbor5.EOF <> true Then%>
<%mdbor5.Movenext%>
<%End If%>
<%mdborl1.Movenext%>
<%Loop%>

<%mdbor2.Close%><%mdbor2U.Close%>
<%mdbor3.Close%><%mdbor3U.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor6a.Close%>
<%mdbor7.Close%>

<%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Description<>'MAAGAAS' AND (IDentifier = 'C')"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summd FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(SummaPlan) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>

<tr class="bold">
<td></td>
<td></td>
<td>Kokku ettev&otildetete kaupa</td>
<td><%=mdbor5("SP")%></td>
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
  <%a0="a4b"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a4b"%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a4b"%>" class="bold">
<%End If%>

</td>
<td>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a4a"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a4a"%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a4a"%>" class="bold">
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

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>

<%If mdbor6a.EOF=True THEN%>
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("EM")%>
<%End If%>

<%If mdbor4.EOF=true Then%>
<%a4=0%>
<%ELse%>
<%a4=mdbor4("Summd")%>
<%End iF%>

<%=CDbl(a2)+CDBL(a1)-CDBL(a3)-CDBL(a6)+CDBL(a7)%>

</td>
<td>

<%If mdbor3.EOF=FALSE Then%>
<%a3=mdbor3("Summc")%>
<%Else%>
<%a3=0%>
<%End If%>
<%If mdbor3U.EOF=FALSE Then%>
<%a4=mdbor3U("Summc")%>
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
<%a3=0%>
<%ELSe%>
<%a3=mdbor6a("SD")%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a4c"%>
  
    <%=Request.Form(a0)%>
  
  <%Else%>
  <%a0="a4c"%>
  <input type="Text" value="<%=CDBL(a3)%>" size="10" name="<%="a4c"%>" class="bold">
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
<%a0="a4c"%>
<%=CDbl(a6)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>
<%mdborl1.Close%>
<%mdbor2.Close%><%mdbor2u.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor6a.Close%>
<%mdbor7.Close%>



<%mdbol1.CommandText="SELECT * FROM Enterprise ORDER BY Enterprise"%>
<%mdborl1.Open mdbol1%>

<%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summc, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506'))) AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND RenovBlock=0 AND Project<>'EJB206' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summc, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Description<>'MAAGAAS' AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summd, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351' GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(SummaPlan) AS SP, Enterprise FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD, Enterprise FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0 GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym, Enterprise FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor7.Open mdbo7%>

<%Do until mdborl1.EOF%>

<%If mdbor7.EOF=true then%>
<%a7=0%>
<%Else%>
<%IF mdbor7("Enterprise")=mdborl1("Enterprise") Then%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%mdbor7.MoveNext%>
<%ELse%>
<%a7=0%>
<%End iF%>
<%End iF%>

<%If mdbor6a.EOF=true then%>
<%a6a=0%>
<%a33=0%>
<%Else%>
<%IF mdbor6a("Enterprise")=mdborl1("Enterprise") Then%>
<%a6a=mdbor6a("EM")%>
<%a33=mdbor6a("SD")%>
<%mdbor6a.MoveNext%>
<%ELse%>
<%a6a=0%>
<%a33=0%>
<%End iF%>
<%End iF%>
<%If mdbor6.EOF=true then%>
<%a6=0%>
<%Else%>
<%IF mdbor6("Enterprise")=mdborl1("Enterprise") Then%>
<%a6=mdbor6("Summy")%>
<%mdbor6.MoveNext%>
<%ELse%>
<%a6=0%>
<%End iF%>
<%End iF%>
<%If mdbor2u.EOF=true then%>
<%a1=0%>
<%Else%>
<%IF mdbor2u("Enterprise")=mdborl1("Enterprise") Then%>
<%a1=mdbor2u("Summi")%>
<%mdbor2u.MoveNext%>
<%ELse%>
<%a1=0%>
<%End iF%>
<%End iF%> 
<%If mdbor4.EOF=true then%>
<%a4=0%>
<%Else%>
<%IF mdbor4("Enterprise")=mdborl1("Enterprise") Then%>
<%a4=mdbor4("Summd")%>
<%mdbor4.MoveNext%>
<%ELse%>
<%a4=0%>
<%End iF%>
<%End iF%>
<%If mdbor2.EOF=true then%>
<%a2=0%>
<%Else%>
<%IF mdbor2("Enterprise")=mdborl1("Enterprise") Then%>
<%a2=mdbor2("Summi")%>
<%mdbor2.MoveNext%>
<%ELse%>
<%a2=0%>
<%End iF%>
<%End iF%>
<%If mdbor3.EOF=true then%>
<%a3=0%>
<%Else%>
<%IF mdbor3("Enterprise")=mdborl1("Enterprise") Then%>
<%a3=mdbor3("Summc")%>
<%mdbor3.MoveNext%>
<%ELse%>
<%a3=0%>
<%End iF%>
<%End iF%>
<%If mdbor3u.EOF=true then%>
<%a9=0%>
<%Else%>
<%IF mdbor3u("Enterprise")=mdborl1("Enterprise") Then%>
<%a9=mdbor3u("Summc")%>
<%mdbor3u.MoveNext%>
<%ELse%>
<%a9=0%>
<%End iF%>
<%End iF%>

<tr class="boldenterp">
<td>


</td>
<td>


</td>
<td>
<%=Mdborl1("EDescr")%>

</td>
<td>

<%=mdbor5("SP")%>

</td>
<td>

<%=CDbl(a2)+CDBL(a1)%>


</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a5b" & mdborl1("Enterprise")%>
  
    <%=Request.Form(a0)%>
  
 <%Else%>
  <%a0="a5b" & mdborl1("Enterprise")%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a5b" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a5a" & mdborl1("Enterprise")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a5a" & mdborl1("Enterprise")%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a5a" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td>

<%=CDbl(a2)+CDBL(a1)-CDBL(a6a)-CDBL(a6)+CDBL(a7)%>

</td>
<td>

<%=cdbl(a3)+cdbl(a9)%>

</td>
<td>

<%=cdbl(a4)%>

</td>
<td>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a5c" & mdborl1("Enterprise")%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a5c" & mdborl1("Enterprise")%>
  <input type="Text" value="<%=CDBL(a33)%>" size="10" name="<%="a5c" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td>

<%a0="a5c"%>
<%=CDbl(a33)-CDbl(a3)-CDbl(a9)+CDbl(a1)+CDbl(a2)%>

</td>
</tr>

<%If mdbor5.EOF <> true Then%>
<%mdbor5.Movenext%>
<%End If%>
<%mdborl1.Movenext%>
<%Loop%>

<%mdbor2.Close%><%mdbor2U.Close%>
<%mdbor3.Close%><%mdbor3u.Close%>
<%mdbor4.Close%>
<%mdbor5.Close%>
<%mdbor6.Close%>
<%mdbor6a.Close%>
<%mdbor7.Close%>



<%mdbo3u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)='999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor3u.Open mdbo3u%>
<%mdbo2.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND ((Konto<>'18164' AND Konto<>'18151') OR ((Konto='18164' OR Konto='18151') AND SUBSTRING(PROJECT,4,3)='999') OR (Konto='18164' AND (PROJECT='EJE573' OR PROJECT='EJE516' OR PROJECT='EJB515' OR PROJECT='EJB583') AND (MES='504' OR MES='505' OR MES='506' OR MES='507'))) AND Yearr='" & ya & "' AND (konto NOT BETWEEN '18410' AND '18433') AND (KONTO<>'43350' AND SUBKONTO<>'4351') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project<>'EJB206' AND (IDentifier = 'C')"%>
<%mdbor2.Open mdbo2%>
<%mdbo2u.CommandText="SELECT ISNULL(ROUND((SUM(ISNULL(DEBET,0))/1000),0),0) AS summi FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE description<>'maagaas' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') AND Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Project='EJB206' AND (IDentifier = 'C')"%>
<%mdbor2u.Open mdbo2u%>
<%mdbo3.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summc FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE SUBSTRING(PROJECT,4,3)<>'999' AND (ISNULL(DOCS_REGNR,0)=0 OR DOCS_REGNR<>'437700') And Yearr='" & ya & "' AND RenovBlock=0 AND (konto NOT BETWEEN '18410' AND '18433') AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND Description<>'MAAGAAS' AND (IDentifier = 'C')"%>
<%mdbor3.Open mdbo3%>
<%mdbo4.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) AS summd FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C') AND KONTO='43350' AND SUBKONTO='4351'"%>
<%mdbor4.Open mdbo4%>
<%mdbo5.CommandText="SELECT  SUM(SummaPlan) AS SP FROM dbo.Delta WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor5.Open mdbo5%>
<%mdbo6.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0)-ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summy FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya-1 & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (((SUBSTRING(MES, 1, 1) = '" & Mid(ya-1,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04)) OR (LEFT(MES,1)=" & Mid(ya,4,1) & " AND RIGHT(MES,2) <04))  AND (IDentifier = 'C')"%>
<%mdbor6.Open mdbo6%>
<%mdbo6a.CommandText="SELECT  SUM(ISNULL(Ettemaks,0)) AS EM,SUM(ISNULL(Saldo,0)) AS SD FROM dbo.ETTE WHERE yearr='" & ya & "' AND enn<>'00' AND RenovBlock=0"%>
<%mdbor6a.Open mdbo6a%>
<%mdbo7.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(DEBET,0))/1000,0),0) as summi, ISNULL(ROUND(SUM(ISNULL(CREDIT,0))/1000,0),0) AS summym FROM dbo.glav_project gp INNER JOIN dbo.GPMAIN as gpm ON gp.ROWID = gpm.ROWID INNER JOIN Main m ON GPM.IDPRoj = m.RowID WHERE Yearr='" & ya & "' AND (konto BETWEEN '18410' AND '18433') AND RenovBlock=0 AND (SUBSTRING(MES, 1, 1) = '" & Mid(ya,4,1) & "') AND (SUBSTRING(MES, 2, 2) >= 04) AND (IDentifier = 'C')"%>
<%mdbor7.Open mdbo7%>
<tr class="bold">
<td>


</td>
<td>


</td>
<td>
Kokku ettev&otilde;tete kaupa, v&auml;lja arvatud plokkide renoveerimine 

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
  <%a0="a7b"%>
  
    <%=Request.Form(a0)%>
  
 <%Else%>
  <%a0="a7b"%>
  <input type="Text" value="<%=CDBL(a6)+CDBL(a6a)%>" size="10" name="<%="a7b"%>" class="bold">
<%End If%>

</td>
<td>

<%If mdbor7.EOF=True THEN%>
<%a7=0%>
<%ELSe%>
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
<%End If%>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
  <%a0="a7a"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a7a"%>
  <input type="Text" value="<%=CDBL(a7)%>" size="10" name="<%="a7a"%>" class="bold">
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
<%a7=CDBL(mdbor7("Summi"))-CDBL(mdbor7("Summym"))%>
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
  <%a0="a7c"%>
  
    <%=Request.Form(a0)%>
  
<%Else%>
  <%a0="a7c"%>
  <input type="Text" value="<%=CDBL(a6)%>" size="10" name="<%="a7c"%>" class="bold">
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
</Form>
</table>
</body>
</html>