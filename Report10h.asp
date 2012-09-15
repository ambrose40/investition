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
<link rel="stylesheet" href="scrollbar.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251"><title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title></Head>
<body class="Report">
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
<%zzz=ya%>
<%zzz=zzz-1%>
<%zzz2=zzz+2%>
<%Dim entt(10,13)%>
<%Dim ent2(10,13)%>
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<%set mdbo =  Server.CreateObject("ADODB.Connection")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")
  set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%>
<%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
<%mdbo.Open ConnectionString%>

<%set mdbo5 = Server.CreateObject("ADODB.Command")%>
<%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5.ActiveConnection = mdbo%>
<%set mdbo4 = Server.CreateObject("ADODB.Command")%>
<%set mdbor4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4.ActiveConnection = mdbo%>
<%set mdbo4a = Server.CreateObject("ADODB.Command")%>
<%set mdbor4a = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4a.ActiveConnection = mdbo%>
<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor2 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>

<table border="1" style="border-collapse: collapse">
<tr>
 <th rowspan="3">Nr</th>
 <th rowspan="3">Projekti Nimetus</th>
 <th rowspan="3">NPV</th>
 <th rowspan="3">IRR</th>
 <th rowspan="2" colspan="2">Ehitusperiood kvartal</th>
 <th rowspan="3">Kalkureeritud maksmus kokku</th>
 <th rowspan="3">Kokku viie aasta invest.</th>
 <th rowspan="3">Tehtud seisuga 01.04.<%=ya%></th>
 <th colspan="10" rowspan="1">INVESTEERINGUD</th>
</tr>
<tr>
 <th colspan="4">Tegelik</th>
 <th colspan="6">Prognoos</th>
</tr>
<tr>
<th>algus</th>
 <th>l&otilde;pp</th>
<%For j=CDbl(ya-5) to Cdbl(ya+4)%>
 <th><%=j%></th>
<%Next%>

</tr>
<tr class="Repnum">
<%For nuu=1 to 19%>
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

<%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE RenovBlock=0 AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor5.Open mdbo5%>
<%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE RenovBlock=0 AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
<%mdbor4.Open mdbo4%>
<%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE RenovBlock=0 AND Identifier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
<%mdbor4a.Open mdbo4a%>
<tr class="Whitetr">
  <td>
   i4.2.31.11
  </td>

  <td>
   wwwwwwwwwwwwwwwwww
  </td>

  <td>
   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1a"%>
    <%=Request.Form(a0)%>
   <%Else%>
    <%a0="a1a"%>
    <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1a"%>" size="10" >
   <%End If%>
  </td>

  <td>
   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1b"%>
    <%=Request.Form(a0)%>
   <%Else%>
    <%a0="a1b"%>
    <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1b"%>" size="10" >
   <%End If%>
  </td>

  <td>
   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1c"%>
    <%=Request.Form(a0)%>
   <%Else%>
    <%a0="a1c"%>
    <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1c"%>" size="10" >
   <%End If%>
  </td>

  <td>
   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1d"%>
    <%=Request.Form(a0)%>
   <%Else%>
    <%a0="a1d"%>
    <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1d"%>" size="10" >
   <%End If%>
  </td>

  <td>
   <%If mdbor4a.EOF=True OR mdbor4a("PASU") & "e" = "e" Then%>
    <%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
     <%sim=0%>
    <%Else%>
     <%sim=mdbor4("SYT")%>
    <%End If%>
   <%Else%>
    <%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
     <%sim=mdbor4("PASU")%>
    <%Else%>
     <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
    <%End If%>
   <%End If%>
   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1y"%>  
    <%=Request.Form(a0)%>
   <%Else%>
    <%a0="a1y"%>
    <%If Request.Form(a0)="" Then%>
     <input type="Text" value="<%=sim%>" name="<%="a1y"%>" size="10" >
    <%Else%>
     <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1y"%>" size="10" >
    <%End If%>
   <%End If%>
  </td>

  <td>
   <%If mdbor4.EOF=True Then%>
    <%sim=0%>
   <%Else%>
    <%sim=mdbor4("SYT")%>
   <%End If%>

   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1z"%>
    <%=Request.Form(a0)%> 
   <%Else%>
    <%a0="a1z"%>
    <%If Request.Form(a0)="" Then%>
     <input type="Text" value="<%=sim%>" name="<%="a1z"%>" size="10" >
    <%Else%>
     <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z"%>" size="10" >
    <%End If%>
   <%End If%>
  </td>

  <td>
   <%If mdbor4a.EOF=True OR mdbor4a("PASU") & "e" = "e" Then%>
    <%sim=0%>
   <%Else%>
    <%sim=mdbor4a("PASU")%>
   <%End If%>
   <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%a0="a1e"%>   
    <%If Request.Form(a0)="" Then%>
     <%=sim%>
    <%Else%>
     <%=Request.Form(a0)%>  
    <%End If%>
   <%Else%>
    <%a0="a1e"%>
    <%If Request.Form(a0)="" Then%>
     <input type="Text" value="<%=sim%>" name="<%="a1e"%>" size="10" >
    <%Else%>
     <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1e"%>" size="10" >
    <%End If%>
   <%End If%>
  </td>
  <%For ja=CDbl(ya-5) to CDbl(ya-1)%>
   <%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m  INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
   <%mdbor2.Open mdbo2%>
   <td> 
    <%If mdbor2.BOF=True then%>
     <%sim=0%>
    <%Else%>
     <%sim=mdbor2("Summi")%>
    <%End If%>

    <%If Request.Form("btn")="Kopeerimiseks" Then%>
     <%a0="a1f" & ja & "_1x"%>   
     <%=Request.Form(a0)%>
     <%If Request.Form(a0)="" Then%>
      <input type="hidden" value="<%=Sim%>" name="<%="a1f" & ja & "_1x"%>">
     <%Else%>
      <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1f" & ja & "_1x"%>">
     <%End If%>
    <%Else%>
     <%a0="a1f" & ja & "_1x"%>
     <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a1f" & ja & "_1x"%>" size="10" >
     <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1f" & ja & "_1x"%>" size="10" >
     <%End If%>
    <%End If%>

    <%mdbor2.Close%>
   </td>
  <%Next%>

  <%For ja=CDbl(ya) to CDbl(ya+4)%>
   <td>
    <%If mdbor5.EOF=True then%>
     <%sim=0%>
    <%Else%>
     <%sim=mdbor5("SY")%>
    <%End If%>

    <%If Request.Form("btn")="Kopeerimiseks" Then%>
     <%a0="a" & ja & "_1x"%>
     <%=Request.Form(a0)%>
    <%Else%>
     <%a0="a" & ja & "_1x"%>
     <%If Request.Form(a0)="" Then%>
      <input type="Text" value="<%=sim%>" name="<%="a" & ja & "_1x"%>" size="10" >
     <%Else%>
      <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "_1x"%>" size="10" >
     <%End If%>
    <%End If%>
   </td>
  <%Next%>
  <%mdbor4.close%>
  <%mdbor4a.close%>
  <%mdbor5.close%>
 </tr>
</table>
</body>
</html>