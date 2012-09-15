<html>
 <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
 <Head>
  <%b= Server.MapPath("\")%>
  <%if request.Cookies("StyleInv")="" then%>
   <%set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%else%>
   <%s=request.Cookies("StyleInv")%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%End if%>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
  <title>InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on</title>
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
  <%set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
  <%s=servFileStream.ReadLine%>
  <%i=servFileStream.ReadLine%>
  <%p=servFileStream.ReadLine%>
  <%servFileStream.Close%>
  <%set mdbo =  Server.CreateObject("ADODB.Connection")%>
  <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
  <%mdbo.Open ConnectionString%>
  <%set mdbol1 = Server.CreateObject("ADODB.Command")%>
  <%set mdborl1 = Server.CreateObject("ADODB.Recordset")%>
  <%mdbol1.ActiveConnection = mdbo%>
  <%set mdbo1 = Server.CreateObject("ADODB.Command")%>
  <%set mdbor = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo1.ActiveConnection = mdbo%>
  <%set mdbo5 = Server.CreateObject("ADODB.Command")%>
  <%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
  <%mdbo5.ActiveConnection = mdbo%>
  <table border=1  width="100%">
   <tr>
    <th rowspan="2" width="58" height="44">Nr</th>
    <th rowspan="2" width="210" height="44">Projekti nimetus</th>
    <th rowspan="2" width="102" height="44">Ehitusaastad (m.a.)</th>
    <th rowspan="2" width="119" height="44">Kalkuleeritud maksumus kokku</th>
    <th rowspan="2" width="93" height="44">L&otilde;petatud seisuga 01.04.<%=ya%></th>
    <th rowspan="2" width="41" height="44"><%=Mid(ya,3,2)%>&nbsp;m.a</th>
    <th colspan="4" rowspan="1" width="145" height="19">INVESTEERINGUD</th>
    <th rowspan="2" width="116" height="44">Projektijuht</th>
   </tr>
   <tr>
    <th width="29" height="19">1 kv</th>
    <th width="30" height="19">2 kv</th>
    <th width="29" height="19">3 kv</th>
    <th width="39" height="19">4 kv</th> 
   </tr>
   <tr class="Repnum">
    <%For nuu=1 to 11%>
     <td width="58" height="16"><%=nuu%></td>
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
   <%aa=0%><%ab=0%><%ac=0%>
   
   <%mdbol1.CommandText="SELECT PRojName FROM inpl WHERE LEN(PRojName)=(SELECT MAX(LEN(PRojName)) FROM Aruanne WHERE Yearr='" & ya & "' AND PRojName not LIKE '%#%') AND Yearr='" & ya & "'"%>
   <%mdborl1.Open mdbol1%>
   <%mdbo1.CommandText="SELECT EmplName as em,EmplFname as fm FROM Aruanne WHERE LEN(EmplName + EmplFName)=(SELECT MAX(LEN(EmplNAme + EmplFName)) FROM Aruanne WHERE Yearr='" & ya & "' AND EmplNAme + EmplFName not LIKE '%#%') AND Yearr='" & ya & "'"%>
   <%mdbor.Open mdbo1%>
   <%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY FROM dbo.Main WHERE IDentifier='P' AND RenovBlock=0 AND SUBSTRING(ProjCode,10,2)<>'00' AND yearr='" & ya & "'"%>
   <%mdbor5.Open mdbo5%>

   <tr class="whitetr">
    <td width="58" height="23">1.1.12.1.</td>
    <td width="210" height="23"><%=mdborl1("ProjName")%></td>
    <td width="102" height="23">
     <%If Request.Form("btn")="Kopeerimiseks" Then%>
      <%a0="a1a"%>
      <%=Request.Form(a0)%>
      <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1a"%>">
     <%Else%>
      <%a0="a1a"%>
      <input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="a1a"%>" style="font-family: Verdana; color: #FFFFFF; background-color: #FFFFFF; border-width:0">
     <%End If%>
    </td>
    <td width="119" height="23"></td>
    <td width="93" height="23"></td>
    <td width="41" height="23">
     <%If mdbor5.BOF=False Then%>
      <%=mdbor5("SY")%>
     <%Else%>
      0
     <%End If%>
    </td>
    <td width="29" height="23">
     <%If mdbor5.BOF=False Then%>
      <%=mdbor5("S1")%>
     <%Else%>
      0
     <%End If%>
    </td>
    <td width="30" height="23">
     <%If mdbor5.BOF=False Then%>
      <%=mdbor5("S2")%>
     <%Else%>
      0
     <%End If%>
    </td>
    <td width="29" height="23">
     <%If mdbor5.BOF=False Then%>
      <%=mdbor5("S3")%>
     <%Else%>
      0
     <%End If%>
    </td>
    <td width="39" height="23">
     <%If mdbor5.BOF=False Then%>
      <%=mdbor5("S4")%>
     <%Else%>
      0
     <%End If%>
    </td>
    <td width="116" height="23"><%=mdbor("em")%>&nbsp<%=mdbor("fm")%></td>
   </tr>
   <%mdbor5.Close%>
  </table>
 </body>
</html>