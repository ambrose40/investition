<html>
 <Head>
  <%b= Server.MapPath("\")%>
  <%if request.Cookies("StyleInv")="" then%>
   <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
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
  <img border="0" src="icons/report.ico" Style=float:Left><p align="center"><a href="Main.asp" target="_top" class="Headlink"><%=ya%>-<%=ya+4%> m.a. SISEARUANNE</a></p>
  <%XYZ=0%>
  <Form Method="POST" Action="Report_r.asp?ye=<%=ya%>">
   <Input type="Submit" name="btn" size="10" Value="Kopeerimiseks" class="Button">
   <Input type="Submit" name="btn" size="10" Value="Redigeerimiseks" class="Button">
   <Input type="Submit" name="btn" size="10" Value="Salvestamiseks" class="Button">
   <%Dim entt(10,13)%>
   <%Dim ent2(10,13)%>
   <%Dim ar2(1,200)%>
   <%set mdbo =  Server.CreateObject("ADODB.Connection")%>
   <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%i=servFileStream.ReadLine%>
   <%p=servFileStream.ReadLine%>
   <%servFileStream.Close%>
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
   <%set mdbo4a = Server.CreateObject("ADODB.Command")%>
   <%set mdbor4a = Server.CreateObject("ADODB.Recordset")%>
   <%mdbo4a.ActiveConnection = mdbo%> 


    
   <table bordercolor="0F0F0F" border="1"  style="border-collapse: collapse">
    <tr>
     <th rowspan="3">Nr</th>
     <th rowspan="3">Projekti Nimetus</th>
     <th rowspan="3">Наименование проекта</th>
     <th rowspan="2" colspan="2">Ehitusperiood kvartal</th>
     <th rowspan="3">Kalkuleeritud maksumus kokku</th>
     <th rowspan="3">Viie aasta investeeringud kokku</th>
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
     <%For nuu=1 to 18%>
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
    <%aa=0%><%ab=0%><%ac=0%>
    <%If Request.Form("btn")="Salvestamiseks" Then%>
     <%mdbol4.CommandText="SELECT DISTINCT PID FROM MAIN WHERE YEARR>=" & ya%>
     <%mdborl4.Open mdbol4%>
     <%Do until mdborl4.EOF%>
        <%mdbo4a.CommandText="UPDATE MAIN SET YEARBEG=(SELECT top 1 Yearr FROM MAIN WHERE PID=" & mdborl4(0) & " AND YEARR>=" & ya & ") WHERE PID=" & mdborl4(0) & " AND YEARBEG<" & ya%>
        <%mdbor4a.Open mdbo4a%>
	<%mdborl4.MoveNext%>
     <%Loop%>
     <%mdborl4.Close%>
     <%mdbol4.CommandText="SELECT DISTINCT dbo.Main.Pid, Main_1.ProjCode as PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier AND  Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >='" & ya & "') AND MAin_1.IDentifier='P' ORDER BY Main_1.ProjCode"%>
     <%mdborl4.Open mdbol4%>
     <%Do until Mdborl4.EOF%>
      <%a0="af" & ya-1 & "x" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%a3="ac" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%a4="ad" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%mdbo2.CommandText="UPDATE MAIN SET PROGNTEH='" & REQUEST.FORM(A0) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR='" & ya & "'"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2.CommandText="UPDATE MAIN SET Ealgus='" & REQUEST.FORM(A3) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR='" & ya & "'"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2.CommandText="UPDATE MAIN SET Elopp='" & REQUEST.FORM(A4) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR='" & ya & "'"%>
      <%mdbor2.Open mdbo2%>
      <%Mdborl4.MovenExt%>
     <%Loop%>
     <%Mdborl4.Close%>
     <%mdbol4.CommandText="SELECT PROJCODE,PID,Yearr,Enterprise,Identifier FROM MAIN WHERE SUBSTRING(PROJCODE,10,2)='00'"%>
     <%mdborl4.Open mdbol4%>
     <%Do until mdborl4.EOF%>
      <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(PROGNTEH,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdborl4("PROJCODE"),1,8) & "' and Yearr='" & MDBORl4("Yearr") & "' AND Enterprise='" & MDBORl4("Enterprise") & "' AND IDEntifier='" & MDBORl4("IDentifier") & "'"%>
      <%mdbor4.Open mdbo4%>
      <%mdbo2.CommandText="UPDATE MAIN SET PROGNTEH='" & MDBOR4("PS") & "' WHERE PID='" & MDBORl4("PID") & "' and Yearr='" & MDBORl4("Yearr") & "' AND Enterprise='" & MDBORl4("Enterprise") & "' AND IDEntifier='" & MDBORl4("IDentifier") & "'"%>
      <%mdbor2.Open mdbo2%>
      <%mdborl4.MoveNExt%>
      <%mdbor4.Close%>
     <%Loop%>
     <%mdborl4.Close%>
    <%End if%>
       
    <%mdbol1.CommandText="SELECT DISTINCT Pid, ProjCode,PC, OracleCode, PRojName,RusName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
    <%mdborl1.Open mdbol1%>
    <%sma=0%>
    <%If Request.Form("btn")="Kopeerimiseks"  Then%>
    <%ELSE%>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <%IF ja<2005 then%>
       <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as summi FROM Main WHERE Identifier='F' AND RenovBlock=0 AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ja & "'"%>
       <%mdbor2.Open mdbo2%>
      <%ELSE%>
       <%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m  INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
       <%mdbor2.Open mdbo2%>
      <%End If%>
      <%If mdbor2.EOF=True Then%>
       <%ar2(1,ja-1990)=0%>
       <%sma=sma+0%>
      <%Else%>
       <%ar2(1,ja-1990)=Mdbor2("Summi")%>
       <%sma=sma+CDBL(Mdbor2("Summi"))%>
      <%End If%>
      <%mdbor2.Close%>
     <%Next%>
      
     <%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE RenovBlock=0 AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor5.Open mdbo5%>
     <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE RenovBlock=0 AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor4.Open mdbo4%>
     <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE RenovBlock=0 AND Identifier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
     <%mdbor4a.Open mdbo4a%>
    <%End if%>
    <tr class="boldProjGrup">
     <td></td>
     <td>INVESTEERINGUD KOKKU  v&auml;lja arvatud plokkide renoveerimine</td>
     <td></td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="a1c"%>   
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a1c"%>
       <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1c"%>" size="10" class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="a1d"%> 
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a1d"%>
       <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1d"%>" size="10" class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="a1y"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%If mdbor4a.EOF=True OR mdbor4a("PASU") & "e" = "e" Then%>
        <%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
         <%sim=0%>
        <%Else%>
         <%sim=CDBL(mdbor4("SYT"))%>
        <%End If%>
       <%Else%>
        <%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
         <%sim=CDBL(mdbor4("PASU"))%>
        <%Else%>
         <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
        <%End If%>
       <%End If%>
       <%sim=sim+sma%>
       <%a0="a1y"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="a1y"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1y"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="a1z"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%If mdbor4.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=mdbor4("SYT")%>
       <%End If%>
       <%a0="a1z"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="a1z"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="a1e"%>
       <%If Request.Form(a0)="" Then%>
        <%=sim%>
       <%Else%>
        <%=Request.Form(a0)%>  
       <%End If%>
      <%Else%>
       <%If mdbor4a.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=CDBL(mdbor4a("PASU"))%>
       <%End If%>
       <%sim=sim+sma%>
       <%a0="a1e"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="a1e"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1e"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="a1f" & ja & "_1x"%>
        <%=Request.Form(a0)%>
        <%If Request.Form(a0)="" Then%>
         <input type="hidden" value="<%=Sim%>" name="<%="a1f" & ja & "_1x"%>">
        <%Else%>
         <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1f" & ja & "_1x"%>">
        <%End If%>
       <%Else%>
        <%sim=ar2(1,ja-1990)%>
        <%a0="a1f" & ja & "_1x"%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="a1f" & ja & "_1x"%>" size="10" class="boldProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1f" & ja & "_1x"%>" size="10" class="boldProjGrup">
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%="a1z" & ya-1 & "_1x"%>
       <%If Request.Form(a0)="" Then%>
        <%=sim%>
       <%Else%>
        <%=Request.Form(a0)%>  
       <%End If%>
      <%Else%>
       <%If mdbor4a.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=mdbor4a("PASU")%>
       <%End If%>
       <%a0="a1z" & ya-1 & "_1x"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="a1z" & ya-1 & "_1x"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z" & ya-1 & "_1x"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya) to CDbl(ya+4)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="a" & ja & "_1x"%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%If mdbor5.EOF=True then%>
         <%sim=0%>
        <%Else%>
         <%sim=mdbor5("SY")%>
        <%End If%>
        <%a0="a" & ja & "_1x"%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="a" & ja & "_1x"%>" size="10" class="boldProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "_1x"%>" size="10" class="boldProjGrup">
        <%End If%>
        <%If mdbor5.EOF=False Then%>
         <%mdbor5.MoveNext%>
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
    </tr>
  <%sma=0%>
    <%If Request.Form("btn")="Kopeerimiseks"  Then%>
    <%ELSE%>
     <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <%IF ja<2005 then%>
       <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as summi FROM Main WHERE Identifier='F' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ja & "'"%>
       <%mdbor2.Open mdbo2%>
      <%ELSE%>
       <%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m  INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
       <%mdbor2.Open mdbo2%>
      <%End If%>
      <%If mdbor2.EOF=True Then%>
       <%sma=sma+0%>
       <%ar2(1,ja-1990)=0%>
      <%Else%>
       <%sma=sma+CDBL(Mdbor2("Summi"))%>
       <%ar2(1,ja-1990)=Mdbor2("Summi")%>
      <%End If%>
      <%mdbor2.Close%>
     <%Next%>
     <%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor5.Open mdbo5%>
     <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor4.Open mdbo4%>
     <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE Identifier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
     <%mdbor4a.Open mdbo4a%>
    <%End If%>
    <tr class="boldProjGrup">
     <td></td>
     <td>INVESTEERINGUD KOKKU koos plokkide renoveerimisega</td>
     <td></td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="ac"%>
       <%=Request.Form(a0)%>
       <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ac"%>">
      <%Else%>
       <%a0="ac"%>
       <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac"%>" size="10" class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="ad"%>
       <%=Request.Form(a0)%>
       <input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ad"%>">
      <%Else%>
       <%a0="ad"%>
       <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad"%>" size="10" class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="ay"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%If mdbor4a.EOF=True OR mdbor4a("PASU") & "e" = "e" Then%>
        <%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
         <%sim=0%>
        <%Else%>
         <%sim=CDBL(mdbor4("SYT"))%>
        <%End If%>
       <%Else%>
        <%If mdbor4.EOF=True OR mdbor4("SYT") & "e" = "e" Then%>
         <%sim=CDBL(mdbor4a("PASU"))%>
        <%Else%>
         <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
        <%End If%>
       <%End If%>
       <%sim=sim+sma%>
       <%a0="ay"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="ay"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="az"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%If mdbor4.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=mdbor4("SYT")%>
       <%End If%>
       <%a0="az"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="az"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="az"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="ae"%>
       <%If Request.Form(a0)="" Then%>
        <%=sim%>
       <%Else%>
        <%=Request.Form(a0)%>  
       <%End If%>
      <%Else%>
       <%If mdbor4a.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=CDBL(mdbor4a("PASU"))%>
       <%End If%>
       <%sim=sim+sma%>
       <%a0="ae"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="ae"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="af" & ja & "x"%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%sim=ar2(1,ja-1990)%>
        <%a0="af" & ja & "x"%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="af" & ja & "x"%>" size="10" class="boldProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x"%>" size="10" class="boldProjGrup">
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%a0="a1z" & ya-1 & "x"%>
       <%If Request.Form(a0)="" Then%>
        <%=sim%>
       <%Else%>
        <%=Request.Form(a0)%>  
       <%End If%>
      <%Else%>
       <%If mdbor4a.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=mdbor4a("PASU")%>
       <%End If%>
       <%a0="a1z" & ya-1 & "x"%>
       <%If Request.Form(a0)="" Then%>
        <input type="Text" value="<%=sim%>" name="<%="a1z" & ya-1 & "x"%>" size="10" class="boldProjGrup">
       <%Else%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z" & ya-1 & "x"%>" size="10" class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya) to CDbl(ya+4)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="a1" & ja & "x"%>
        <%=Request.Form(a0)%>   
       <%Else%>
        <%If mdbor5.EOF=True then%>
         <%sim=0%>
        <%Else%>
         <%sim=mdbor5("SY")%>
        <%End If%>
        <%a0="a1" & ja & "x"%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="a1" & ja & "x"%>" size="10" class="boldProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1" & ja & "x"%>" size="10" class="boldProjGrup">
        <%End If%>
        <%If mdbor5.EOF=False Then%>
         <%mdbor5.MoveNext%>
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
    </tr>
    <%Do until mdborl1.EOF%>
     <%sma=0%>
     <%If Request.Form("btn")="Kopeerimiseks"  Then%>
     <%ELSE%>
      <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%> 
      <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
       <%IF ja<2005 then%>
        <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as summi FROM Main WHERE  LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND Identifier='F' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ja & "'"%>
        <%mdbor2.Open mdbo2%>
       <%ELSE%>
        <%mdbo2.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
        <%mdbor2.Open mdbo2%>
       <%End If%>
       <%If mdbor2.EOF=True Then%>
        <%ar2(1,ja-1990)=0%>
        <%sma=sma+0%>
       <%Else%>
        <%ar2(1,ja-1990)=Mdbor2("Summi")%>
        <%sma=sma+CDBL(Mdbor2("Summi"))%>
       <%End If%>
       <%mdbor2.Close%>
      <%NExt%>
      <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
      <%mdbor5.Open mdbo5%>
      <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
      <%mdbor4.Open mdbo4%>
      <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND Identifier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
      <%mdbor4a.Open mdbo4a%>
     <%End If%>
     <tr class="projGrup">
      <td>&nbsp;<%'=mdborl1("Pid")%><%=MID(mdborl1("PC"),2,1)%></td>
      <td><%=mdborl1("ProjName")%></td>
      <td><%=mdborl1("RusName")%></td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="ac" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="ac" & mdborl1("Pid")%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl1("Pid")%>" size="10" class="projGrup">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="ad" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="ad" & mdborl1("Pid")%>
        <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl1("Pid")%>" size="10" class="projGrup">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="ay" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%else%>
        <%If mdbor4a.EOF=True Then%>
         <%If mdbor4.EOF=True Then%>
          <%sim=0%>
         <%Else%>
          <%sim=CDBL(mdbor4("SYT"))%>
         <%End If%>
        <%Else%>
         <%If mdbor4.EOF=True Then%>
          <%sim=CDBL(mdbor4a("PASU"))%>
         <%Else%>
          <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
         <%End If%>
        <%End If%>
        <%sim=sim+sma%>
        <%a0="ay" & mdborl1("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="ay" & mdborl1("Pid")%>" size="10" class="ProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl1("Pid")%>" size="10" class="ProjGrup">
        <%End If%>
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="az" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%If mdbor4.EOF=True Then%>
         <%sim=0%>
        <%Else%>
         <%sim=mdbor4("SYT")%>
        <%End If%>
        <%a0="az" & mdborl1("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="az" & mdborl1("Pid")%>" size="10" class="ProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" & mdborl1("Pid")%>" size="10" class="ProjGrup">
        <%End If%>
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="ae" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%If mdbor4a.EOF=True Then%>
         <%sim=0%>
        <%Else%>
         <%sim=CDBL(mdbor4a("PASU"))%>
        <%End If%>
        <%sim=sim+sma%>
        <%a0="ae" & mdborl1("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="ae" & mdborl1("Pid")%>" size="10" class="projGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" & mdborl1("Pid")%>" size="10" class="projGrup">
        <%End If%>
       <%End If%>
      </td>
      <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="af" & ja & "x" & mdborl1("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%sim=ar2(1,ja-1990)%>
         <%a0="af" & ja & "x" & mdborl1("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl1("Pid")%>" size="10" class="projGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl1("Pid")%>" size="10" class="projGrup">
         <%End If%>
        <%End If%>
       </td>
      <%Next%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
        <%a0="a1z" & ya-1 & "x" & mdborl1("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <%=sim%>
        <%Else%>
         <%=Request.Form(a0)%>  
        <%End If%>
       <%Else%>
        <%If mdbor4a.EOF=True Then%>
         <%sim=0%>
        <%Else%>
         <%sim=mdbor4a("PASU")%>
        <%End If%>
        <%a0="a1z" & ya-1 & "x" & mdborl1("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input type="Text" value="<%=sim%>" name="<%="a1z" & ya-1 & "x" & mdborl1("Pid")%>" size="10" class="boldProjGrup">
        <%Else%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z" & ya-1 & "x" & mdborl1("Pid")%>" size="10" class="boldProjGrup">
        <%End If%>
       <%End If%>
      </td>
      <%For ja=CDbl(ya) to CDbl(ya+4)%>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="a" & ja & "x" & mdborl1("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%If mdbor5.EOF=True then%>
          <%sim=0%>
         <%Else%>
          <%sim=mdbor5("SY")%>
         <%End If%>
         <%a0="a" & ja & "x" & mdborl1("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl1("Pid")%>" size="10" class="ProjGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl1("Pid")%>" size="10" class="ProjGrup">
         <%End If%>
         <%If mdbor5.EOF=False Then%>
          <%mdbor5.MoveNext%>
         <%End If%>
        <%End If%>
       </td>
      <%Next%>
     </tr>
     <%mdbol2.CommandText="SELECT DISTINCT Pid,PC,ProjCode,ProjName,OracleCode,RusName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND  SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY ProjCode"%>
     <%mdborl2.Open mdbol2%>
     <%Do until mdborl2.EOF%>
      <%sma=0%>
      <%If Request.Form("btn")="Kopeerimiseks"  Then%>
      <%ELSE%>
       <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
       <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
        <%IF ja<2005 then%>
         <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as summi FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND SUBSTRING(ProjCode,10,2)<>'00' AND Identifier='F' AND Yearr='" & ja & "'"%>
         <%mdbor2.Open mdbo2%>
        <%ELSE%>
         <%mdbo2.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
         <%mdbor2.Open mdbo2%>
        <%End If%>
        <%If mdbor2.EOF=True Then%>
         <%ar2(1,ja-1990)=0%>
         <%sma=sma+0%>
        <%Else%>
         <%ar2(1,ja-1990)=Mdbor2("Summi")%>
         <%sma=sma+CDBL(Mdbor2("Summi"))%>
        <%End If%>
        <%mdbor2.Close%>
       <%NExt%>
       <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
       <%mdbor5.Open mdbo5%>
       <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
       <%mdbor4.Open mdbo4%>
       <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
       <%mdbor4a.Open mdbo4a%>
      <%End If%>
      <tr class="projGrup">
       <td>&nbsp;<%'=mdborl2("Pid")%><%=MID(mdborl2("PC"),2,1) & "." & MID(mdborl2("PC"),5,1)%>.</td>
       <td><%=mdborl2("ProjName")%></td>
       <td><%=mdborl2("RusName")%></td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="ac" & mdborl2("Pid")%>  
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="ac" & mdborl2("Pid")%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl2("Pid")%>" size="10" class="projGrup">
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="ad" & mdborl2("Pid")%>   
         <%=Request.Form(a0)%>  
        <%Else%>
         <%a0="ad" & mdborl2("Pid")%>
         <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl2("Pid")%>" size="10" class="projGrup">
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="ay" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%If mdbor4a.EOF=True Then%>
          <%If mdbor4.EOF=True Then%>
           <%sim=0%>
          <%Else%>
           <%sim=CDBL(mdbor4("SYT"))%>
          <%End If%>
         <%Else%>
          <%If mdbor4.EOF=True Then%>
           <%sim=CDBL(mdbor4a("PASU"))%>
          <%Else%>
           <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
          <%End If%>
         <%End If%>
         <%sim=sim+sma%>
         <%a0="ay" & mdborl2("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="ay" & mdborl2("Pid")%>" size="10" class="ProjGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl2("Pid")%>" size="10" class="ProjGrup">
         <%End If%>
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="az" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%If mdbor4.EOF=True Then%>
          <%sim=0%>
         <%Else%>
          <%sim=mdbor4("SYT")%>
         <%End If%>
         <%a0="az" & mdborl2("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="az" & mdborl2("Pid")%>" size="10" class="ProjGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" & mdborl2("Pid")%>" size="10" class="ProjGrup">
         <%End If%>
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="ae" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%If mdbor4a.EOF=True Then%>
          <%sim=0%>
         <%Else%>
          <%sim=CDBL(mdbor4a("PASU"))%>
         <%End If%>
         <%sim=sim+sma%>
         <%a0="ae" & mdborl2("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="ae" & mdborl2("Pid")%>" size="10" class="projGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" & mdborl2("Pid")%>" size="10" class="projGrup">
         <%End If%>
        <%End If%>
       </td>
       <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="af" & ja & "x" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%sim=ar2(1,ja-1990)%>
          <%a0="af" & ja & "x" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl2("Pid")%>" size="10" class="projGrup">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl2("Pid")%>" size="10" class="projGrup">
          <%End If%>
         <%End If%>
        </td>
       <%Next%>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%a0="a1z" & ya-1 & "x" & mdborl2("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <%=sim%>
         <%Else%>
          <%=Request.Form(a0)%>  
         <%End If%>
        <%Else%>
         <%If mdbor4a.EOF=True Then%>
          <%sim=0%>
         <%Else%>
          <%sim=mdbor4a("PASU")%>
         <%End If%>
         <%a0="a1z" & ya-1 & "x" & mdborl2("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input type="Text" value="<%=sim%>" name="<%="a1z" & ya-1 & "x" & mdborl2("Pid")%>" size="10" class="boldProjGrup">
         <%Else%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z" & ya-1 & "x" & mdborl2("Pid")%>" size="10" class="boldProjGrup">
         <%End If%>
        <%End If%>
       </td>
       <%For ja=CDbl(ya) to CDbl(ya+4)%>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="a" & ja & "x" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%If mdbor5.EOF=True then%>
           <%sim=0%>
          <%Else%>
           <%sim=mdbor5("SY")%>
          <%End If%>
          <%a0="a" & ja & "x" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl2("Pid")%>" size="10" class="ProjGrup">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl2("Pid")%>" size="10" class="ProjGrup">
          <%End If%>
          <%If mdbor5.EOF=False Then%>
           <%mdbor5.MoveNext%>
          <%End If%>
         <%End If%>
        </td>
       <%Next%>
      </tr>
      <%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentifier='C' AND Yearr>=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "' ORDER BY ENTERPRISE"%>
      <%mdborl3.Open mdbol3%>
      <%Do until mdborl3.EOF%>
       <%sma=0%>
       <%If Request.Form("btn")="Kopeerimiseks"  Then%>
       <%ELSE%>
        <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
        <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
         <%IF ja<2005 then%>
          <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as summi FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND SUBSTRING(ProjCode,10,2)<>'00' AND Identifier='F' AND Yearr='" & ja & "'"%>
          <%mdbor2.Open mdbo2%>
         <%ELSE%>
          <%mdbo2.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ja & "' and ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
          <%mdbor2.Open mdbo2%>
         <%End If%>
         <%If mdbor2.EOF=True Then%>
          <%ar2(1,ja-1990)=0%>
          <%sma=sma+0%>
         <%Else%>
          <%ar2(1,ja-1990)=Mdbor2("Summi")%>
          <%sma=sma+CDBL(Mdbor2("Summi"))%>

         <%End If%>
         <%mdbor2.Close%>
        <%Next%>
        <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND IDentifier='p' and SUBSTRING(PROJCODE,4,2)<>'00' AND SUBSTRING(PROJCODE,7,2)<>'00' AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,5)='" & MID(mdborl2("PC"),1,5) & "' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
        <%mdbor5.Open mdbo5%>
        <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
        <%mdbor4.Open mdbo4%>
        <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND Identifier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
        <%mdbor4a.Open mdbo4a%>
       <%End If%>
       <tr class="Enterp">
        <td></td>
        <td><%=mdborl3("EDescr")%></td>
        <td></td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="ay" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
          <%sim=Request.Form(a0)%>
          <%entt(mdborl3("Enterprise"),1)=entt(mdborl3("Enterprise"),1) + Cdbl(sim)%>
          <%jo=jo+1%>
         <%Else%>
          <%If mdbor4a.EOF=True Then%>
           <%If mdbor4.EOF=True Then%>
            <%sim=0%>
           <%Else%>
            <%sim=CDBL(mdbor4("SYT"))%>
           <%End If%>
          <%Else%>
           <%If mdbor4.EOF=True Then%>
            <%sim=CDBL(mdbor4a("PASU"))%>
           <%Else%>
            <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
           <%End If%>
          <%End If%>
          <%sim=sim+sma%>
          <%entt(mdborl3("Enterprise"),1)=entt(mdborl3("Enterprise"),1) + Cdbl(sim)%>
          <%jo=jo+1%>
          <%a0="ay" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="ay" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
          <%End If%>
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
          <%sim=Request.Form(a0)%>
          <%entt(mdborl3("Enterprise"),2)=entt(mdborl3("Enterprise"),2) + Cdbl(sim)%>
          <%jo=jo+1%>
         <%Else%>
          <%If mdbor4.EOF=True Then%>
           <%sim=0%>
          <%Else%>
           <%sim=mdbor4("SYT")%>
          <%End If%>
          <%entt(mdborl3("Enterprise"),2)=entt(mdborl3("Enterprise"),2) + Cdbl(sim)%>
          <%jo=jo+1%>
          <%a0="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
          <%End If%>
         <%End If%> 
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
          <%sim=Request.Form(a0)%>
          <%entt(mdborl3("Enterprise"),3)=entt(mdborl3("Enterprise"),3) + Cdbl(sim)%>
          <%jo=jo+1%>
         <%Else%>
          <%If mdbor4a.EOF=True Then%>
           <%sim=0%>
          <%Else%>
           <%sim=CDBL(mdbor4a("PASU"))%>
          <%End If%>
          <%sim=sim+sma%>
          <%entt(mdborl3("Enterprise"),3)=entt(mdborl3("Enterprise"),3) + Cdbl(sim)%>
          <%jo=jo+1%>
          <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
          <%End If%>
         <%End If%>
        </td>
        <%jo=4%>
        <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
         <td>
          <%If Request.Form("btn")="Kopeerimiseks"  Then%>
           <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
           <%=Request.Form(a0)%>
           <%sim=Request.Form(a0)%>
           <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
           <%jo=jo+1%>
          <%Else%>
           <%sim=ar2(1,ja-1990)%>
           <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
           <%jo=jo+1%>
           <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
           <%If Request.Form(a0)="" Then%>
            <input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
           <%Else%>
            <input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
           <%End If%>
          <%End If%>
         </td>
        <%Next%>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
          <%a0="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <%=sim%>
          <%Else%>
           <%=Request.Form(a0)%>  
          <%End If%>
         <%Else%>
          <%If mdbor4a.EOF=True Then%>
           <%sim=0%>
          <%Else%>
           <%sim=mdbor4a("PASU")%>
          <%End If%>
          <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
          <%jo=jo+1%>
          <%a0="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input type="Text" value="<%=sim%>" name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="boldEnterp">
          <%Else%>
           <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="boldEnterp">
          <%End If%>
         <%End If%>
        </td>
        <%For ja=CDbl(ya) to CDbl(ya+4)%>
         <td>
          <%If Request.Form("btn")="Kopeerimiseks"  Then%>
           <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
           <%=Request.Form(a0)%>  
           <%sim=Request.Form(a0)%>
           <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
           <%jo=jo+1%>
          <%Else%>
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
           <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
           <%jo=jo+1%>
           <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
           <%If Request.Form(a0)="" Then%>
            <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
           <%Else%>
            <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" class="Enterp">
           <%End If%>
          <%End If%>
         </td>
        <%Next%>
       </tr>
       <%mdbol4.CommandText="SELECT DISTINCT dbo.Main.Pid, Main_1.ProjCode as PC FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier AND  Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >='" & ya & "') /*and Main_1.Yearr >='" & ya & "'*/ AND MAin_1.Enterprise='" & Mdborl3("Enterprise") & "' AND MAin_1.IDentifier='C' AND SUBSTRING(MAin_1.ProjCode,4,2)<>'00' AND SUBSTRING(MAin_1.ProjCode,7,2)<>'00' AND  SUBSTRING(MAin_1.ProjCode,1,5)='" & MID(mdborl2("PC"),1,5) & "' ORDER BY Main_1.ProjCode"%>
       <%mdborl4.Open mdbol4%>

       <%Do Until mdborl4.EOF%>
        <%sma=0%>
        <%If MDBORl4("Pid")=abcde THEN%>
         <%mdborl4.MoveNExt%>
        <%ELSE%>
         <%Abcde=MDBORl4("Pid")%>
         <%mdbog.CommandText="SELECT DISTINCT RusName,ProjName,PC,RenovBlock,Yearr,FootNote FROM inpl WHERE Yearr >= '" & ya & "' AND Pid = '" & Mdborl4("Pid") & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND IDentifier='C' ORDER BY PC,Yearr"%>
         <%mdborg.Open mdbog%>
         <%If Request.Form("btn")="Kopeerimiseks"  Then%>
         <%ELSE%>
          <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
          <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
           <%IF ja<2005 then%>
            <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as summi FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND Identifier='F' AND Yearr='" & ja & "'"%>
            <%mdbor2.Open mdbo2%>
           <%ELSE%>
            <%mdbo2.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentifier = Main_1.IDentifier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise='" & mdborl3("Enterprise") & "' AND m.Pid='" & mdborl4("Pid") & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C')"%>
            <%mdbor2.Open mdbo2%>
           <%End If%>
           <%If mdbor2.EOF=True Then%>
            <%ar2(1,ja-1990)=0%>
            <%sma=sma+0%>
           <%Else%>
            <%ar2(1,ja-1990)=Mdbor2("Summi")%>
            <%sma=sma+CDBL(Mdbor2("Summi"))%>
           <%End If%>
           <%mdbor2.Close%>
          <%Next%>
          <%mdbo5.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as sy, Yearr FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND Identifier='P' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
          <%mdbor5.Open mdbo5%>
          <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND Identifier='P' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
          <%mdbor4.Open mdbo4%>
          <%mdbo4a.CommandText="SELECT ISNULL(PROGNTEH,0) as PASU,Ealgus,Elopp FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND Identifier='P' AND Yearr='" & ya & "'"%>
          <%mdbor4a.Open mdbo4a%>
         <%End If%>
         <tr>
          <td>
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
           <%End IF%>&nbsp&nbsp&nbsp
           <%If Mdborg("Footnote") & "e" <> "e" then%>
            <a name=<%="vira" & Fotnum%>></a>{<%=Fotnum%>}<%fotnum=fotnum+1%>
           <%end if%>
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks"  Then%>
            <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%=Request.Form(a0)%>
           <%Else%>
            <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If mdbor4a.EOF=true then%>
             <%sim=""%>
            <%Else%>
             <%sim=mdbor4a("Ealgus")%>
            <%End If%>
            <%If Request.Form(a0)="" Then%>
             <input type="Text" value="<%=sim%>" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%Else%>
             <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%End If%>
           <%End If%>
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks"  Then%>
            <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%=Request.Form(a0)%>
           <%Else%>
            <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If mdbor4a.EOF=true then%>
             <%sim=""%>
            <%Else%>
             <%sim=mdbor4a("Elopp")%>
            <%End If%>
            <%If Request.Form(a0)="" Then%>
             <input type="Text" value="<%=sim%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%Else%>
             <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%End If%>
           <%End If%>
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks"  Then%>
            <%a0="ay" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%=Request.Form(a0)%>
            <%sim=Request.Form(a0)%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),1)=ent2(mdborl3("Enterprise"),1)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
           <%Else%>
            <%If mdbor4a.EOF=True Then%>
             <%If mdbor4.EOF=True Then%>
              <%sim=0%>
             <%Else%>
              <%sim=CDbl(mdbor4("SYT"))%>
             <%End If%>
            <%Else%>
             <%If mdbor4.EOF=True Then%>
              <%sim=CDbl(mdbor4a("PASU"))%>
             <%Else%>
              <%sim=CDbl(mdbor4("SYT")) + CDbl(mdbor4a("PASU"))%>
             <%End If%>
            <%End If%>
            <%sim=sim+sma%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),1)=ent2(mdborl3("Enterprise"),1)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
            <%a0="ay" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If Request.Form(a0)="" Then%>
             <input type="Text" value="<%=sim%>" name="<%="ay" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%Else%>
             <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ay" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%End If%>
           <%End If%>
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks"  Then%>
            <%a0="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%=Request.Form(a0)%> 
            <%sim=Request.Form(a0)%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),2)=ent2(mdborl3("Enterprise"),2)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
           <%Else%>
            <%If mdbor4.EOF=True Then%>
             <%sim=0%>
            <%Else%>
             <%sim=mdbor4("SYT")%>
            <%End If%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),2)=ent2(mdborl3("Enterprise"),2)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
            <%a0="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If Request.Form(a0)="" Then%>
             <input type="Text" value="<%=sim%>" name="<%="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%Else%>
             <input type="Text" value="<%=Request.Form(a0)%>" name="<%="az" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%End If%>
           <%End If%> 
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks"  Then%>
            <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%=Request.Form(a0)%>
            <%sim=Request.Form(a0)%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
           <%Else%>
            <%If mdbor4a.EOF=True Then%>
             <%sim=0%>
            <%Else%>
             <%sim=CDBL(mdbor4a("PASU"))%>
            <%End If%>
            <%sim=sim+sma%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),3)=ent2(mdborl3("Enterprise"),3)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
            <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If Request.Form(a0)="" Then%>
	     <%'=mdborl3("Enterprise")%>
             <input type="Text" value="<%=sim%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%Else%>
             <input type="Text" value="<%=Request.Form(a0)%>" name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%End If%>
           <%End If%>
          </td>
          <%jo=4%>
          <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
           <td>
            <%If Request.Form("btn")="Kopeerimiseks"  Then%>
             <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
             <%=Request.Form(a0)%>
             <%sim=Request.Form(a0)%>
             <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
              <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
             <%End If%>
             <%Jo=jo+1%>
            <%Else%>
             <%sim=ar2(1,ja-1990)%>
             <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
              <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
             <%End If%>
             <%Jo=jo+1%>
             <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
             <%If Request.Form(a0)="" Then%>
              <input type="Text" value="<%=sim%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
             <%Else%>
              <input type="Text" value="<%=Request.Form(a0)%>" name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
             <%End If%>
            <%End If%>
           </td>
          <%Next%>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks"  Then%>
            <%a0="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If Request.Form(a0)="" Then%>
             <%=sim%>
            <%Else%>
             <%=Request.Form(a0)%>  
            <%End If%>
           <%Else%>
            <%If mdbor4a.EOF=True Then%>
             <%sim=0%>
            <%Else%>
             <%sim=mdbor4a("PASU")%>
            <%End If%>
            <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
             <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
            <%End If%>
            <%Jo=jo+1%>
            <%a0="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If Request.Form(a0)="" Then%>
             <input type="Text" value="<%=sim%>" name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" class="bold">
            <%Else%>
             <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" class="bold">
            <%End If%>
           <%End If%>
          </td>
          <%For ja=CDbl(ya) to CDbl(ya+4)%>
           <td>
            <%If Request.Form("btn")="Kopeerimiseks"  Then%>
             <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
             <%=Request.Form(a0)%>
             <%sim=Request.Form(a0)%>
             <%If mdborg("RenovBlock")<>0 AND (MID(mdborg("PC"),10,2)<>"00") Then%>
              <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
             <%End If%>
             <%Jo=jo+1%>
            <%Else%>
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
             <%If mdborg("RenovBlock")<>0 AND MID(mdborg("PC"),10,2)<>"00" Then%>
              <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
             <%End If%>
             <%Jo=jo+1%>
             <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
             <%If Request.Form(a0)="" Then%>
              <input type="Text" value="<%=sim%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
             <%Else%>
              <input type="Text" value="<%=Request.Form(a0)%>" name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
             <%End If%>
            <%End If%>
           </td>
          <%Next%>
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
    <%If Request.Form("btn")="Kopeerimiseks"  Then%>
    <%ELSE%>
     <%mdbor4.Close%>
    <%End If%>
     
    <%'mdbol4.CommandText="EXEC Yearbegrep @ya=" & ya%>
    <%'mdborl4.Open mdbol4%>
     
     <%Dim koku(13)%><%Dim kok2(13)%>
     <tr class="bold">
      <td colspan="18">Kokku ettev&otildette kaupa</td>
     </tr>
     <%mdbo4.CommandText="SELECT * FROM Enterprise ORDER BY ENTERPRISE"%>
     <%mdbor4.Open mdbo4%>
     <%Do until mdbor4.EOF%>
      <tr class="boldenterp">
       <td></td>
       <td><%=mdbor4("EDescr")%></td>
       <%For nuu=3 to 5%>
        <td></td>
       <%Next%>
       <%For nuu=6 to 18%>
        <td><%=entt(Mdbor4("Enterprise"),nuu-5)-ent2(Mdbor4("Enterprise"),nuu-5)%></td>
        <%koku(nuu-5)=koku(nuu-5)+entt(Mdbor4("Enterprise"),nuu-5)-ent2(Mdbor4("Enterprise"),nuu-5)%>
       <%Next%>
      </tr>
      <%mdbor4.Movenext%>
     <%Loop%>
     <tr class="bold">
      <td></td>
      <td>Kokku</td>
      <%For nuu=3 to 5%>
       <td></td>
      <%Next%>
      <%For nuu=6 to 18%>
       <td><%=koku(nuu-5)%></td>
      <%Next%>
     </tr>
     <tr class="bold">
      <td colspan="18">Kokku ettev&otildette kaupa, v&auml;lja arvatud plokkide renoveerimine</td>
     </tr>
     <%mdbor4.MoveFirst%>
     <%Do until mdbor4.EOF%>
      <tr class="boldEnterp">
       <td></td>
       <td><%=mdbor4("EDescr")%></td>
       <%For nuu=3 to 5%>
        <td></td>
       <%Next%>
       <%For nuu=6 to 18%>
        <td><%=entt(Mdbor4("Enterprise"),nuu-5)%></td>
        <%kok2(nuu-5)=kok2(nuu-5)+entt(Mdbor4("Enterprise"),nuu-5)%>
       <%Next%>
      </tr>
      <%mdbor4.Movenext%>
     <%Loop%>
     <tr class="bold">
      <td></td>
      <td>Kokku</td>
      <%For nuu=3 to 5%>
       <td></td>
      <%Next%>
      <%For nuu=6 to 18%>
       <td><%=kok2(nuu-5)%></td>
      <%Next%>
     </tr>
  
    
    <%mdbo1.CommandText="SELECT DISTINCT Footnote,PC FROM inpl WHERE IDentifier='C' AND Yearr>='" & ya & "' AND footnote iS NOT NULL AND Footnote<>'' ORDER BY PC"%>
    <%mdborg.Open mdbo1%>
    <%Fotnum=1%>
    <tr>
     <td Colspan="21">{} M&Auml;RKUSED:</td>
    </tr>
    <%Do until Mdborg.EOF%>
     <tr>
      <td colspan="21"><a href=<%="report_r.asp?" & Request.QueryString & "#vira" & fotnum%>>{<%=fotnum%>}&nbsp<%=Mdborg("footnote")%></a>
       <%fotnum=fotnum+1%>
       <%mdborg.MoveNext%>
      </td>
     </tr>
    <%Loop%>
    <%mdborg.close%>
   </table>
  </Form> 
 </body>
</html>