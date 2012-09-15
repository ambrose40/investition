<Html>
 <Head>
  <%b= Server.MapPath("\")%>
  <%If Request.Cookies("StyleInv")="" Then%>
   <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\Style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <Link Rel="Stylesheet" href="<%=s%>" Type="text/css">
  <%Else%>
   <%s=Request.Cookies("StyleInv")%>
   <Link Rel="Stylesheet" href="<%=s%>" Type="text/css">
  <%End If%>
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Title>
   InFormatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
  </Title>
 </Head>
 <Body Class="Report">
  <%fotnum=1%>
  <%If Request.Form("btn")="OK" Then%>
   <%ya=Request.Form("ye")%>
  <%Else%>
   <%ya=Request.QueryString("ye")%>
  <%End If%>

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
   <%End If%>
  <%End If%>
  <img border="0" src="icons/report.ico" Style=float:Left><p align="center"><a href="Main.asp" Class="HeadLink" target="_top">10 aasta investeerimiskava <%=ya%>-<%=ya+4%> aastate kaupa</a></p>
  <br>
  <%XYZ=0%>
  <Form Method="POST" Action="Report10.asp?ye=<%=ya%>">
   <Input Type="Submit" Name="btn" size="10" Value="Kopeerimiseks" Class="Button">
   <Input Type="Submit" Name="btn" size="10" Value="Redigeerimiseks" Class="Button">
   <Input Type="Submit" Name="btn" size="10" Value="Salvestamiseks" Class="Button">
   <hr>
   <%Dim entt(10,13)%>
   <%Dim ent2(10,13)%>
   <%Dim ar2(1,200)%>
   <%Set mdbo =  Server.CreateObject("ADODB.Connection")%>
   <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%i=servFileStream.ReadLine%>
   <%p=servFileStream.ReadLine%>
   <%servFileStream.Close%>

   <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
   <%mdbo.Open ConnectionString%>
   <%Set mdbol1 = Server.CreateObject("ADODB.Command")%>
   <%Set mdborl1 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbol1.ActiveConnection = mdbo%>
   <%Set mdbol2 = Server.CreateObject("ADODB.Command")%>
   <%Set mdborl2 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbol2.ActiveConnection = mdbo%>
   <%Set mdbol3 = Server.CreateObject("ADODB.Command")%>
   <%Set mdborl3 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbol3.ActiveConnection = mdbo%>
   <%Set mdbo3 = Server.CreateObject("ADODB.Command")%>
   <%Set mdbor3 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbo3.ActiveConnection = mdbo%>
   <%Set mdbo2 = Server.CreateObject("ADODB.Command")%>
   <%Set mdbor2 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbo2.ActiveConnection = mdbo%>
   <%Set mdbo1 = Server.CreateObject("ADODB.Command")%>
   <%Set mdbor1 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbo1.ActiveConnection = mdbo%>
   <%Set mdbol4 = Server.CreateObject("ADODB.Command")%>
   <%Set mdborl4 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbol4.ActiveConnection = mdbo%>
   <%Set mdbo5 = Server.CreateObject("ADODB.Command")%>
   <%Set mdbor5 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbo5.ActiveConnection = mdbo%>
   <%Set mdbog = Server.CreateObject("ADODB.Command")%>
   <%Set mdborg = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbog.ActiveConnection = mdbo%>
   <%Set mdbo4 = Server.CreateObject("ADODB.Command")%>
   <%Set mdbor4 = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbo4.ActiveConnection = mdbo%>
   <%Set mdbo4a = Server.CreateObject("ADODB.Command")%>
   <%Set mdbor4a = Server.CreateObject("ADODB.RecordSet")%>
   <%mdbo4a.ActiveConnection = mdbo%>
  <%'mdbol4.CommandText="EXEC Yearbegrep @ya=" & ya%>
   <%'mdborl4.Open mdbol4%>
   <Table border="1"  Style="border-collapse: collapse">
    <tr>
     <th rowspan="3">Nr</th>
     <th rowspan="3">Projekti Nimetus</th>
     <th rowspan="3">NPV</th>
     <th rowspan="3">IRR</th>
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
    <tr Class="Repnum">
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
    <%If Request.Form("btn")="Salvestamiseks" Then%>
     <%mdbol4.CommandText="SELECT DISTINCT PID FROM MAIN WHERE YEARR>=" & ya%>
     <%mdborl4.Open mdbol4%>

     <%Do until mdborl4.EOF%>
        <%mdbo4a.CommandText="UPDATE MAIN SET YEARBEG=(SELECT top 1 Yearr FROM MAIN WHERE PID=" & mdborl4(0) & " AND YEARR>=" & ya & ") WHERE PID=" & mdborl4(0) & " AND YEARBEG<" & ya%>
        <%mdbor4a.Open mdbo4a%>
	<%mdborl4.MoveNext%>
     <%Loop%>
     <%mdborl4.Close%>
     <%mdbol4.CommandText="SELECT DISTINCT dbo.Main.Pid, Main_1.ProjCode as PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND  Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >='" & ya & "') AND MAin_1.IDentIfier='P' ORDER BY Main_1.ProjCode"%>
     <%mdborl4.Open mdbol4%>

     <%Do until Mdborl4.EOF%>
      <%a0="a1z" & ya-1 & "x" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%a1="aa" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%a2="ab" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%a3="ac" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%a4="ad" & mdborl4("Enterprise") & "_" & mdborl4("Pid")%>
      <%'="UPDATE MAIN SET PROGNTEH='" & Request.FORM(A0) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & mdborl4("Pid") & "' AND YEARR>='" & ya & "')"%>
      <%mdbo2.CommandText="UPDATE MAIN SET PROGNTEH='" & Request.FORM(A0) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & mdborl4("Pid") & "' AND YEARR>='" & ya & "')"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2.CommandText="UPDATE MAIN SET NPV='" & Request.FORM(A1) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & mdborl4("Pid") & "' AND YEARR>='" & ya & "')"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2.CommandText="UPDATE MAIN SET IRR='" & Request.FORM(A2) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & mdborl4("Pid") & "' AND YEARR>='" & ya & "')"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2.CommandText="UPDATE MAIN SET Ealgus='" & Request.FORM(A3) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & mdborl4("Pid") & "' AND YEARR>='" & ya & "')"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo2.CommandText="UPDATE MAIN SET Elopp='" & Request.FORM(A4) & "' WHERE Pid='" & mdborl4("Pid") & "' AND Enterprise='" & mdborl4("Enterprise") & "' AND YEARR=(SELECT top 1 Yearr FROM MAIN WHERE PID='" & mdborl4("Pid") & "' AND YEARR>='" & ya & "')"%>
      <%mdbor2.Open mdbo2%>

      <%Mdborl4.MovenExt%>
     <%Loop%>
     <%Mdborl4.Close%>
     <%mdbol4.CommandText="SELECT PROJCODE,PID,Yearr,Enterprise,IdentIfier FROM MAIN WHERE SUBSTRING(PROJCODE,10,2)='00'"%>
     <%mdborl4.Open mdbol4%>
     <%Do until mdborl4.EOF%>
      <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(PROGNTEH,0)),0) as PS FROM MAIN WHERE LEN(PROJCODE)>9  AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,8)='" & MID(mdborl4("PROJCODE"),1,8) & "' and Yearr='" & MDBORl4("Yearr") & "' AND Enterprise='" & MDBORl4("Enterprise") & "' AND IDEntIfier='" & MDBORl4("IDentIfier") & "'"%>
      <%mdbor2.Open mdbo2%>
      <%mdbo4a.CommandText="UPDATE MAIN SET PROGNTEH='" & MDBOR2("PS") & "' WHERE PID='" & MDBORl4("PID") & "' and Yearr='" & MDBORl4("Yearr") & "' AND Enterprise='" & MDBORl4("Enterprise") & "' AND IDEntIfier='" & MDBORl4("IDentIfier") & "'"%>
      <%mdbor4a.Open mdbo4a%>
      <%mdborl4.MoveNExt%>
      <%mdbor2.Close%>
     <%Loop%>
     <%mdborl4.Close%>
    <%End If%>

    <%mdbol1.CommandText="SELECT DISTINCT Pid, ProjCode,PC, PRojName FROM inpl WHERE IDentIfier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='00' ORDER BY ProjCode"%>
    <%mdborl1.Open mdbol1%>
    <%sma=0%>
    <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%Else%>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <%If ja<2005 Then%>
       <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(MAin.SummYE,0)),0) as Summi FROM (SELECT DISTINCT TOP 100 PERCENT dbo.Main.Pid, Main_1.ProjCode AS PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >= '" & ya & "') AND (Main_1.IDentIfier = 'F')) allp INNER JOIN dbo.Main ON allp.Pid = dbo.Main.Pid WHERE (dbo.Main.Yearr = '" & ja & "') AND SUBSTRING(Main.ProjCode,10,2)<>'00' AND Main.RenovBlock='0' AND (dbo.Main.IDentIfier = 'F')"%>
       <%mdbor2.Open mdbo2%>
      <%Else%>
       <%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m  INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.RenovBlock=0 AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentIfier = 'C')"%>
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
     <%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE RenovBlock=0 AND IdentIfier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor5.Open mdbo5%>
     <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE RenovBlock=0 AND IdentIfier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor4.Open mdbo4%>
     <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE RenovBlock=0 AND IdentIfier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
     <%mdbor4a.Open mdbo4a%>
    <%End If%>
    <tr Class="boldProjGrup">
     <td></td>
     <td>INVESTEERINGUD KOKKU  v&auml;lja arvatud plokkide renoveerimine</td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a1a"%> 
       <%=Request.Form(a0)%>  
      <%Else%>
       <%a0="a1a"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1a"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a1b"%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="a1b"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1b"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a1c"%>
       <%=Request.Form(a0)%>  
      <%Else%>
       <%a0="a1c"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1c"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a1d"%>
       <%=Request.Form(a0)%> 
      <%Else%>
       <%a0="a1d"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1d"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="a1y"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1y"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="a1z"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="a1e"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1e"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a1f" & ja & "_1x"%>    
        <%=Request.Form(a0)%>   
        <%If Request.Form(a0)="" Then%>
         <input Type="hidden" Value="<%=Sim%>" Name="<%="a1f" & ja & "_1x"%>">
        <%Else%>
         <input Type="hidden" Value="<%=Request.Form(a0)%>" Name="<%="a1f" & ja & "_1x"%>">
        <%End If%>
       <%Else%>
        <%sim=ar2(1,ja-1990)%>
        <%a0="a1f" & ja & "_1x"%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="a1f" & ja & "_1x"%>" size="10" Class="boldProjGrup">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1f" & ja & "_1x"%>" size="10" Class="boldProjGrup">
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="a1z" & ya-1 & "_1x"%>  
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
        <input Type="Text" Value="<%=sim%>" Name="<%="a1z" & ya-1 & "_1x"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z" & ya-1 & "_1x"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya) to CDbl(ya+4)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a" & ja & "_1x"%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%If mdbor5.EOF=True Then%>
         <%sim=0%>
        <%Else%>
         <%sim=mdbor5("SY")%>
        <%End If%>
        <%a0="a" & ja & "_1x"%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="a" & ja & "_1x"%>" size="10" Class="boldProjGrup">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a" & ja & "_1x"%>" size="10" Class="boldProjGrup">
        <%End If%>
        <%If mdbor5.EOF=False Then%>
         <%mdbor5.MoveNext%>
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
    </tr>
    <%sma=0%>
    <%If Request.Form("btn")="Kopeerimiseks" Then%>
    <%Else%>
     <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%> 
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <%If ja<2005 Then%>
       <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(MAin.SummYE,0)),0) as Summi FROM (SELECT DISTINCT TOP 100 PERCENT dbo.Main.Pid, Main_1.ProjCode AS PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >= '" & ya & "') AND (Main_1.IDentIfier = 'F')) allp INNER JOIN dbo.Main ON allp.Pid = dbo.Main.Pid WHERE (dbo.Main.Yearr = '" & ja & "') AND SUBSTRING(Main.ProjCode,10,2)<>'00' AND (dbo.Main.IDentIfier = 'F')"%>
       <%mdbor2.Open mdbo2%>
      <%Else%>
       <%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m  INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentIfier = 'C')"%>
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
     <%mdbo5.CommandText="SELECT SUM(SummYe) as sy,Yearr FROM Main WHERE IdentIfier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor5.Open mdbo5%>
     <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE IdentIfier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
     <%mdbor4.Open mdbo4%>
     <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE IdentIfier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
     <%mdbor4a.Open mdbo4a%>
    <%End If%>
    <tr Class="boldProjGrup">
     <td></td>
     <td>INVESTEERINGUD KOKKU koos plokkide renoveerimisega</td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="aa"%>
       <%=Request.Form(a0)%>
       <input Type="hidden" Value="<%=Request.Form(a0)%>" Name="<%="aa"%>">
      <%Else%>
       <%a0="aa"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="aa"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ab"%>
       <%=Request.Form(a0)%>
       <input Type="hidden" Value="<%=Request.Form(a0)%>" Name="<%="ab"%>">
      <%Else%>
       <%a0="ab"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ab"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ac"%>
       <%=Request.Form(a0)%>
       <input Type="hidden" Value="<%=Request.Form(a0)%>" Name="<%="ac"%>">
      <%Else%>
       <%a0="ac"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ac"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ad"%>
       <%=Request.Form(a0)%>
       <input Type="hidden" Value="<%=Request.Form(a0)%>" Name="<%="ad"%>">
      <%Else%>
       <%a0="ad"%>
       <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ad"%>" size="10" Class="boldProjGrup">
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="ay"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ay"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="az"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="az"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="ae"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ae"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="af" & ja & "x"%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%sim=ar2(1,ja-1990)%>
        <%a0="af" & ja & "x"%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="af" & ja & "x"%>" size="10" Class="boldProjGrup">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="af" & ja & "x"%>" size="10" Class="boldProjGrup">
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
        <input Type="Text" Value="<%=sim%>" Name="<%="a1z" & ya-1 & "x"%>" size="10" Class="boldProjGrup">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z" & ya-1 & "x"%>" size="10" Class="boldProjGrup">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya) to CDbl(ya+4)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a1" & ja & "x"%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%If mdbor5.EOF=True Then%>
         <%sim=0%>
        <%Else%>
         <%sim=mdbor5("SY")%>
        <%End If%>
        <%a0="a1" & ja & "x"%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="a1" & ja & "x"%>" size="10" Class="boldProjGrup">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1" & ja & "x"%>" size="10" Class="boldProjGrup">
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
     <%If Request.Form("btn")="Kopeerimiseks" Then%>
     <%Else%>
      <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
      <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
       <%If ja<2005 Then%>
        <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(MAin.SummYE,0)),0) as Summi FROM (SELECT DISTINCT TOP 100 PERCENT dbo.Main.Pid, Main_1.ProjCode AS PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >= '" & ya & "') AND (Main_1.IDentIfier = 'F')) allp INNER JOIN dbo.Main ON allp.Pid = dbo.Main.Pid WHERE LEFT(Main.ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND (dbo.Main.Yearr = '" & ja & "') AND SUBSTRING(Main.ProjCode,10,2)<>'00' AND (dbo.Main.IDentIfier = 'F')"%>
        <%mdbor2.Open mdbo2%>
       <%Else%>
        <%mdbo2.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentIfier = 'C')"%>
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
      <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND IdentIfier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
      <%mdbor5.Open mdbo5%>
      <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND IdentIfier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
      <%mdbor4.Open mdbo4%>
      <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE LEFT(ProjCode,2)='" & MID(mdborl1("PC"),1,2) & "' AND IdentIfier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
      <%mdbor4a.Open mdbo4a%>
     <%End If%>

     <tr Class="ProjGrup">
      <td><%=MID(mdborl1("PC"),2,1)%></td>
      <td><%=mdborl1("ProjName")%></td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="aa" & mdborl1("Pid")%>
        <%=Request.Form(a0)%> 
       <%Else%>
        <%a0="aa" & mdborl1("Pid")%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="aa" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="ab" & mdborl1("Pid")%>
        <%=Request.Form(a0)%> 
       <%Else%>
        <%a0="ab" & mdborl1("Pid")%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ab" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="ac" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>
       <%Else%>
        <%a0="ac" & mdborl1("Pid")%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ac" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="ad" & mdborl1("Pid")%>
        <%=Request.Form(a0)%>  
       <%Else%>
        <%a0="ad" & mdborl1("Pid")%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ad" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="ay" & mdborl1("Pid")%>
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
        <%a0="ay" & mdborl1("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="ay" & mdborl1("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ay" & mdborl1("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
        <%End If%>
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
         <input Type="Text" Value="<%=sim%>" Name="<%="az" & mdborl1("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="az" & mdborl1("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
        <%End If%>
       <%End If%>
      </td>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
         <input Type="Text" Value="<%=sim%>" Name="<%="ae" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ae" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
        <%End If%>
       <%End If%>
      </td>
      <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="af" & ja & "x" & mdborl1("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%sim=ar2(1,ja-1990)%>
         <%a0="af" & ja & "x" & mdborl1("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input Type="Text" Value="<%=sim%>" Name="<%="af" & ja & "x" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
         <%Else%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="af" & ja & "x" & mdborl1("Pid")%>" size="10" Class="ProjGrup">
         <%End If%>
        <%End If%>
       </td>
      <%Next%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
         <input Type="Text" Value="<%=sim%>" Name="<%="a1z" & ya-1 & "x" & mdborl1("Pid")%>" size="10" Class="boldProjGrup">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z" & ya-1 & "x" & mdborl1("Pid")%>" size="10" Class="boldProjGrup">
        <%End If%>
       <%End If%>
      </td>
      <%For ja=CDbl(ya) to CDbl(ya+4)%>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="a" & ja & "x" & mdborl1("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%If mdbor5.EOF=True Then%>
          <%sim=0%>
         <%Else%>
          <%sim=mdbor5("SY")%>
         <%End If%>
         <%a0="a" & ja & "x" & mdborl1("Pid")%>
         <%If Request.Form(a0)="" Then%>
          <input Type="Text" Value="<%=sim%>" Name="<%="a" & ja & "x" & mdborl1("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
         <%Else%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a" & ja & "x" & mdborl1("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
         <%End If%>
         <%If mdbor5.EOF=False Then%>
          <%mdbor5.MoveNext%>
         <%End If%>
        <%End If%>
       </td>
      <%Next%>
     </tr>

     <%mdbol2.CommandText="SELECT DISTINCT Pid,PC,ProjCode,ProjName FROM inpl WHERE IDentIfier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND  SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY ProjCode"%>
     <%mdborl2.Open mdbol2%>

     <%Do until mdborl2.EOF%>
      <%sma=0%>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
      <%Else%>
       <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
          
       <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
        <%If ja<2005 Then%>
         <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(MAin.SummYE,0)),0) as Summi FROM (SELECT DISTINCT TOP 100 PERCENT dbo.Main.Pid, Main_1.ProjCode AS PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >= '" & ya & "') AND (Main_1.IDentIfier = 'F')) allp INNER JOIN dbo.Main ON allp.Pid = dbo.Main.Pid WHERE LEFT(Main.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND (dbo.Main.Yearr = '" & ja & "') AND SUBSTRING(Main.ProjCode,10,2)<>'00' AND (dbo.Main.IDentIfier = 'F')"%>
         <%mdbor2.Open mdbo2%>
        <%Else%>
         <%mdbo2.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentIfier = 'C')"%>
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
       <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND IdentIfier='P'  AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
       <%mdbor5.Open mdbo5%>
       <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND IdentIfier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
       <%mdbor4.Open mdbo4%>
       <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND IdentIfier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
       <%mdbor4a.Open mdbo4a%>
      <%End If%>
      <tr Class="ProjGrup">
       <td><%=MID(mdborl2("PC"),2,1) & "." & MID(mdborl2("PC"),5,1)%>.</td>
       <td><%=mdborl2("ProjName")%></td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="aa" & mdborl2("Pid")%>   
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="aa" & mdborl2("Pid")%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="aa" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="ab" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="ab" & mdborl2("Pid")%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ab" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="ac" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="ac" & mdborl2("Pid")%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ac" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%a0="ad" & mdborl2("Pid")%>
         <%=Request.Form(a0)%>
        <%Else%>
         <%a0="ad" & mdborl2("Pid")%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ad" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
          <input Type="Text" Value="<%=sim%>" Name="<%="ay" & mdborl2("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
         <%Else%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ay" & mdborl2("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
         <%End If%>
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
          <input Type="Text" Value="<%=sim%>" Name="<%="az" & mdborl2("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
         <%Else%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="az" & mdborl2("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
         <%End If%>
        <%End If%>
       </td>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
          <input Type="Text" Value="<%=sim%>" Name="<%="ae" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
         <%Else%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ae" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
         <%End If%>
        <%End If%>
       </td>
       <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="af" & ja & "x" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%sim=ar2(1,ja-1990)%>
          <%a0="af" & ja & "x" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input Type="Text" Value="<%=sim%>" Name="<%="af" & ja & "x" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
          <%Else%>
           <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="af" & ja & "x" & mdborl2("Pid")%>" size="10" Class="ProjGrup">
          <%End If%>
         <%End If%>
        </td>
       <%Next%>
       <td>
        <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
          <input Type="Text" Value="<%=sim%>" Name="<%="a1z" & ya-1 & "x" & mdborl2("Pid")%>" size="10" Class="boldProjGrup">
         <%Else%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z" & ya-1 & "x" & mdborl2("Pid")%>" size="10" Class="boldProjGrup">
         <%End If%>
        <%End If%>
       </td>
       <%For ja=CDbl(ya) to CDbl(ya+4)%>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="a" & ja & "x" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%If mdbor5.EOF=True Then%>
           <%sim=0%>
          <%Else%>
           <%sim=mdbor5("SY")%>
          <%End If%>
          <%a0="a" & ja & "x" & mdborl2("Pid")%>
          <%If Request.Form(a0)="" Then%>
           <input Type="Text" Value="<%=sim%>" Name="<%="a" & ja & "x" & mdborl2("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
          <%Else%>
           <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a" & ja & "x" & mdborl2("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFAA; border-width:0">
          <%End If%>
          <%If mdbor5.EOF=False Then%>
           <%mdbor5.MoveNext%>
          <%End If%>
         <%End If%>
        </td>
       <%Next%>
      </tr>
      <%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentIfier='C' AND Yearr>=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "' ORDER BY ENTERPRISE"%>
      <%mdborl3.Open mdbol3%>
      <%Do until mdborl3.EOF%>
       <%sma=0%>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%Else%>
        <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
        <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
         <%If ja<2005 Then%>
          <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(MAin.SummYE,0)),0) as Summi FROM (SELECT DISTINCT TOP 100 PERCENT dbo.Main.Pid, Main_1.ProjCode AS PC, Main_1.Enterprise FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >= '" & ya & "') AND (Main_1.IDentIfier = 'F')) allp INNER JOIN dbo.Main ON allp.Pid = dbo.Main.Pid WHERE  Main.Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(Main.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND (dbo.Main.Yearr = '" & ja & "') AND SUBSTRING(Main.ProjCode,10,2)<>'00' AND (dbo.Main.IDentIfier = 'F')"%>
          <%mdbor2.Open mdbo2%>
         <%Else%>
          <%mdbo2.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Yearr='" & ja & "' and ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentIfier = 'C')"%>
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
        <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND IDentIfier='p' and SUBSTRING(PROJCODE,4,2)<>'00' AND SUBSTRING(PROJCODE,7,2)<>'00' AND SUBSTRING(PROJCODE,10,2)<>'00' AND SUBSTRING(PROJCODE,1,5)='" & MID(mdborl2("PC"),1,5) & "' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
        <%mdbor5.Open mdbo5%>
        <%mdbo4.CommandText="SELECT ISNULL(SUM(ISNULL(SummYe,0)),0) as SYT FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND IdentIfier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
        <%mdbor4.Open mdbo4%>
        <%mdbo4a.CommandText="SELECT ISNULL(SUM(ISNULL(PrognTeh,0)),0) as PASU FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND LEFT(ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND IdentIfier='P' AND SUBSTRING(ProjCode,7,2)<>'00' AND SUBSTRING(ProjCode,10,2)<>'00' AND Yearr='" & ya & "'"%>
        <%mdbor4a.Open mdbo4a%>
       <%End If%>
       <tr Class="Enterp">
        <td></td>
        <td><%=mdborl3("EDescr")%></td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ab" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
         <%Else%>
          <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
           <input Type="Text" Value="<%=sim%>" Name="<%="ay" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
          <%Else%>
           <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ay" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
          <%End If%>
         <%End If%>
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
           <input Type="Text" Value="<%=sim%>" Name="<%="az" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
          <%Else%>
           <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="az" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
          <%End If%>
         <%End If%> 
        </td>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
           <input Type="Text" Value="<%=sim%>" Name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
          <%Else%>
           <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
          <%End If%>
         <%End If%>
        </td>
        <%jo=4%>
        <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
         <td>
          <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
            <input Type="Text" Value="<%=sim%>" Name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
           <%Else%>
            <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
           <%End If%>
          <%End If%>
         </td>
        <%Next%>
        <td>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
          <%a0="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
          <%=Request.Form(a0)%>
          <%sim=Request.Form(a0)%>
          <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
          <%jo=jo+1%>
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
           <input Type="Text" Value="<%=sim%>" Name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="boldEnterp">
          <%Else%>
           <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="boldEnterp">
          <%End If%>
         <%End If%>
        </td>
        <%For ja=CDbl(ya) to CDbl(ya+4)%>
         <td>
          <%If Request.Form("btn")="Kopeerimiseks" Then%>
           <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
           <%=Request.Form(a0)%>
           <%sim=Request.Form(a0)%>
           <%entt(mdborl3("Enterprise"),jo)=entt(mdborl3("Enterprise"),jo) + CDbl(sim)%>
           <%jo=jo+1%>
          <%Else%>
           <%If mdbor5.EOF=True Then%>
            <%sim=0%>
           <%Else%>
            <%If CDBL(mdbor5("Yearr"))=Ja Then%>
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
            <input Type="Text" Value="<%=sim%>" Name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
           <%Else%>
            <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" size="10" Class="Enterp">
           <%End If%>
          <%End If%>
         </td>
        <%Next%>
       </tr>
          
       <%mdbol4.CommandText="SELECT DISTINCT dbo.Main.Pid, Main_1.ProjCode as PC FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier AND  Main_1.Yearr = dbo.Main.YearBeg WHERE (dbo.Main.Yearr >='" & ya & "') AND MAin.Enterprise='" & Mdborl3("Enterprise") & "' AND MAin.IDentIfier='C' AND SUBSTRING(MAin.ProjCode,4,2)<>'00' AND SUBSTRING(MAin.ProjCode,7,2)<>'00' AND  SUBSTRING(MAin.ProjCode,1,5)='" & MID(mdborl2("PC"),1,5) & "' ORDER BY Main_1.ProjCode"%>
       <%mdborl4.Open mdbol4%>
       <%Do Until mdborl4.EOF%>
        <%If MDBORl4("Pid")=abcde THEN%>
         <%mdborl4.MoveNExt%>
        <%Else%>
         <%Abcde=MDBORl4("Pid")%>
             
         <%mdbog.CommandText="SELECT DISTINCT ProjName,PC,RenovBlock,Yearr,FootNote,NPV,IRR,Ealgus,Elopp FROM inpl WHERE Yearr >= '" & ya & "' AND Pid = '" & Mdborl4("Pid") & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND IDentIfier='C' ORDER BY PC,Yearr"%>
         <%mdborg.Open mdbog%>
         <%sma=0%>
         <%If Request.Form("btn")="Kopeerimiseks" Then%>
         <%Else%>
          <%mdbor5.Close%><%mdbor4.Close%><%mdbor4a.Close%>
          <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
           <%If ja<2005 Then%>
            <%mdbo2.CommandText="SELECT ISNULL(SUM(ISNULL(SummYE,0)),0) as Summi FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND IdentIfier='F' AND Yearr='" & ja & "'"%>
            <%mdbor2.Open mdbo2%>
           <%Else%>
            <%mdbo2.CommandText="SELECT ISNULL(ROUND(SUM(ISNULL(GP.DEBET,0))/1000,0),0) AS summi FROM (SELECT DISTINCT Main_1.* FROM dbo.Main INNER JOIN dbo.Main Main_1 ON dbo.Main.Pid = Main_1.Pid AND dbo.Main.Enterprise = Main_1.Enterprise AND dbo.Main.IDentIfier = Main_1.IDentIfier WHERE (dbo.Main.Yearr >= '" & ya & "')) AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Enterprise='" & mdborl3("Enterprise") & "' AND m.Pid='" & mdborl4("Pid") & "' AND m.Yearr='" & ja & "' AND ((LEFT(MES,1)='" & MID(ja,4,1) & "') OR (LEFT(MES,1)='" & MID(ja+1,4,1) & "' AND RIGHT(MES,1)<04)) AND (GP.DEBET IS NOT NULL) AND (m.IDentIfier = 'C')"%>
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
          
          <%mdbo5.CommandText="SELECT SUM(SummYe) as sy, Yearr FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND IdentIfier='P' GROUP BY Yearr HAVING Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
          <%mdbor5.Open mdbo5%>
          <%mdbo4.CommandText="SELECT ISNULL(SUM(SummYe),0) as SYT FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND IdentIfier='P' AND Yearr>'" & ya-1 & "' AND Yearr<='" & ya+4 & "'"%>
          <%mdbor4.Open mdbo4%>
          <%mdbo4a.CommandText="SELECT ISNULL(PROGNTEH,0) as PASU FROM Main WHERE Enterprise='" & mdborl3("Enterprise") & "' AND Pid='" & mdborl4("Pid") & "' AND IdentIfier='P' AND Yearr>='" & ya & "'"%>
          <%mdbor4a.Open mdbo4a%>
         <%End If%>
         <tr>
          <td>
           <%If LEN(mdborl4("PC"))>9 Then%>
            <%If MID(mdborl4("PC"),10,2)="00" Then%>
             <%If Mid(mdborl4("PC"),7,1)="0" THen%>
              <%Endi=MId(mdborl4("PC"),8,1)%>
             <%Else%>
              <%Endi=MId(mdborl4("PC"),7,2)%>
             <%End If%>
             <%=MID(mdborl4("PC"),2,1) & "." & MID(mdborl4("PC"),5,1) & "." & Endi%>.
            <%Else%>
             <%If Mid(mdborl4("PC"),7,1)="0" THen%>
              <%Endi=MId(mdborl4("PC"),8,1)%>
             <%Else%>
              <%Endi=MId(mdborl4("PC"),7,2)%>
             <%End If%>
             <%If Mid(mdborl4("PC"),10,1)="0" THen%>
              <%Endy=MId(mdborl4("PC"),11,1)%>
             <%Else%>
              <%Endy=MId(mdborl4("PC"),10,2)%>
             <%End If%>
             <%=MID(mdborl4("PC"),2,1) & "." & MID(mdborl4("PC"),5,1) & "." & Endi & "." & Endy%>.
            <%End If%> 
           <%Else%>
            <%If Mid(mdborl4("PC"),7,1)="0" THen%>
             <%Endi=MId(mdborl4("PC"),8,1)%>
            <%Else%>
             <%Endi=MId(mdborl4("PC"),7,2)%>
            <%End If%>
            <%=MID(mdborl4("PC"),2,1) & "." & MID(mdborl4("PC"),5,1) & "." & Endi%>.
           <%End If%>
          </td>
          <td>
           <%If LEN(mdborl4("PC"))>9 and MID(mdborl4("PC"),10,2)="00" Then%>
            <%=mdborg("ProjName")%>&nbspsealhulgas:
           <%Else%>
            <%=mdborg("ProjName")%>
           <%End If%>&nbsp&nbsp&nbsp
           <%If Mdborg("Footnote") & "e" <> "e" Then%>
            <a Name=<%="vira" & Fotnum%>></a>{<%=Fotnum%>}
            <%fotnum=fotnum+1%>
           <%End If%>
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks" Then%>
            <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>  
            <%=Request.Form(a0)%>
           <%Else%>
            <%a0="aa" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
            <%If mdborg.EOF=True Then%>
             <%sim=0%>
            <%Else%>
             <%sim=mdborg("NPV")%>
            <%End If%>
            <%If Request.Form(a0)="" Then%>
             <input Type="Text" Value="<%=sim%>" Name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%Else%>
             <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
            <%End If%>
           <%End If%>
          </td>
          <td>
           <%If Request.Form("btn")="Kopeerimiseks" Then%>
            <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="ab" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If mdborg.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=mdborg("IRR")%>
       <%End If%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="ab" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ab" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If mdborg.EOF=True Then%>
        <%sim=""%>
       <%Else%>
        <%sim=mdborg("Ealgus")%>
       <%End If%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ac" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%=Request.Form(a0)%>
      <%Else%>
       <%a0="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If mdborg.EOF=True Then%>
        <%sim=""%>
       <%Else%>
        <%sim=mdborg("Elopp")%>
       <%End If%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ad" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ay" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%=Request.Form(a0)%>
       <%sim=Request.Form(a0)%>
       <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
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
       <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
        <%ent2(mdborl3("Enterprise"),1)=ent2(mdborl3("Enterprise"),1)+CDbl(sim)%>
       <%End If%>
       <%Jo=jo+1%>
       <%a0="ay" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="ay" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ay" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
       <%End If%>
      <%End If%>
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%=Request.Form(a0)%> 
       <%sim=Request.Form(a0)%>  
       <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
        <%ent2(mdborl3("Enterprise"),2)=ent2(mdborl3("Enterprise"),2)+CDbl(sim)%>
       <%End If%>
       <%Jo=jo+1%>
      <%Else%>
       <%If mdbor4.EOF=True Then%>
        <%sim=0%>
       <%Else%>
        <%sim=mdbor4("SYT")%>
       <%End If%>
       <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
        <%ent2(mdborl3("Enterprise"),2)=ent2(mdborl3("Enterprise"),2)+CDbl(sim)%>
       <%End If%>
       <%Jo=jo+1%>
       <%a0="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="az" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="az" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
       <%End If%>
      <%End If%> 
     </td>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
       <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%=Request.Form(a0)%>
       <%sim=Request.Form(a0)%>
       <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
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
       <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
        <%ent2(mdborl3("Enterprise"),3)=ent2(mdborl3("Enterprise"),3)+CDbl(sim)%>
       <%End If%>
       <%Jo=jo+1%>
       <%a0="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="ae" &  mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
       <%End If%>
      <%End If%>
     </td>
     <%jo=4%>
     <%For ja=CDbl(ya-5) to CDbl(ya-2)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
        <%=Request.Form(a0)%>
        <%sim=Request.Form(a0)%>
        <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
         <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
        <%End If%>
        <%Jo=jo+1%>
       <%Else%>
        <%sim=ar2(1,ja-1990)%>
        <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
         <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
        <%End If%>
        <%Jo=jo+1%>
        <%a0="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="af" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" >
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
     <td>
      <%If Request.Form("btn")="Kopeerimiseks" Then%>
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
       <%a0="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
       <%If Request.Form(a0)="" Then%>
        <input Type="Text" Value="<%=sim%>" Name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFFF; border-width:0">
       <%Else%>
        <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a1z" & ya-1 & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:700; color: #000000; background-color: #FFFFFF; border-width:0">
       <%End If%>
      <%End If%>
     </td>
     <%For ja=CDbl(ya) to CDbl(ya+4)%>
      <td>
       <%If Request.Form("btn")="Kopeerimiseks" Then%>
        <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
        <%=Request.Form(a0)%>
        <%sim=Request.Form(a0)%>
        <%If mdborg("RenovBlock")<>0 AND (MID(mdborl4("PC"),10,2)<>"00") Then%>
         <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
        <%End If%>
        <%Jo=jo+1%>
       <%Else%>
        <%If mdbor5.EOF=True Then%>
         <%sim=0%>
        <%Else%>
         <%If CDBL(mdbor5("Yearr"))=Ja Then%>
          <%sim=mdbor5("SY")%>
          <%mdbor5.MoveNext%>
         <%Else%>
          <%sim=0%>
         <%End If%>
        <%End If%>
        <%If mdborg("RenovBlock")<>0 AND MID(mdborl4("PC"),10,2)<>"00" Then%>
         <%ent2(mdborl3("Enterprise"),jo)=ent2(mdborl3("Enterprise"),jo)+CDbl(sim)%>
        <%End If%>
        <%Jo=jo+1%>
        <%a0="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>
        <%If Request.Form(a0)="" Then%>
         <input Type="Text" Value="<%=sim%>" Name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
        <%Else%>
         <input Type="Text" Value="<%=Request.Form(a0)%>" Name="<%="a" & ja & "x" & mdborl3("Enterprise") & "_" & mdborl4("Pid")%>" size="10" Style="font-family: Verdana; font-weight:400; color: #000000; background-color: #FFFFFF; border-width:0">
        <%End If%>
       <%End If%>
      </td>
     <%Next%>
    </tr>
    <%mdborl4.MoveNext%><%mdborg.Close%>
   <%End If%>
  <%Loop%>
  <%mdborl4.Close%>
  <%mdborl3.MoveNext%>
 <%Loop%>
 <%mdborl3.Close%>
 <%mdborl2.MoveNext%>
<%Loop%>

<%mdborl2.Close%>
<%mdborl1.MoveNext%>

<%Loop%>
<%mdborl1.Close%>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%Else%>
 <%mdbor4.Close%>
<%mdbor5.Close%><%mdbor4a.Close%>
<%End If%>


 <%'mdbol4.CommandText="EXEC Yearbegrep @ya=" & ya%>
<%'mdborl4.Open mdbol4%>

<%Dim koku(13)%><%Dim kok2(13)%>
<tr Class="bold">
 <td colspan="18">Kokku ettev&otildette kaupa</td>
</tr>
<%mdbo4.CommandText="SELECT * FROM Enterprise ORDER BY ENTERPRISE"%>
<%mdbor4.Open mdbo4%>
<%Do until mdbor4.EOF%>
 <tr Class="boldEnterp">
  <td></td>
  <td><%=mdbor4("EDescr")%></td>
  <%For nuu=3 to 5%>
   <td></td>
  <%Next%>
  <%For nuu=6 to 19%>
   <td><%=entt(Mdbor4("Enterprise"),nuu-6)-ent2(Mdbor4("Enterprise"),nuu-6)%></td>
   <%koku(nuu-6)=koku(nuu-6)+entt(Mdbor4("Enterprise"),nuu-6)-ent2(Mdbor4("Enterprise"),nuu-6)%>
  <%Next%>
 </tr>
 <%mdbor4.MoveNext%>
<%Loop%>
<tr>
 <td></td>
 <td Class="bold">Kokku</td>
 <%For nuu=3 to 5%>
  <td></td>
 <%Next%>
 <%For nuu=6 to 19%>
  <td><%=koku(nuu-6)%></td>
 <%Next%>
</tr>
<tr Class="bold">
 <td colspan="18">Kokku ettev&otildette kaupa, v&auml;lja arvatud plokkide renoveerimine</td>
</tr>
<%mdbor4.MoveFirst%>
<%Do until mdbor4.EOF%>
 <tr Class="boldEnterp">
  <td></td>
  <td><%=mdbor4("EDescr")%></td>
  <%For nuu=3 to 5%>
   <td></td>
  <%Next%>
  <%For nuu=6 to 19%>
 <td><%=entt(Mdbor4("Enterprise"),nuu-6)%></td>
<%kok2(nuu-6)=kok2(nuu-6)+entt(Mdbor4("Enterprise"),nuu-6)%>
<%Next%>
</tr>
<%mdbor4.MoveNext%>
<%Loop%>

<tr Class="bold">
<td>
</td>
<td>
Kokku
</td>
<%For nuu=3 to 5%>
 <td></td>
<%Next%>
<%For nuu=6 to 19%>
 <td><%=kok2(nuu-6)%></td>
<%Next%>
</tr>
</Form>
<%mdbor4.Close%>
<%mdbo4.CommandText="SELECT DISTINCT Footnote,PC FROM inpl WHERE IDentIfier='C' AND Yearr>='" & ya & "' AND footnote iS NOT NULL AND Footnote<>'' ORDER BY PC"%>
<%mdbor4.Open mdbo4%>
<%Fotnum=1%>
<tr bordercolor="FFFFFF">
<td Colspan="21" bordercolor="FFFFFF">
{} M&Auml;RKUSED:
</td>
</tr>
<%Do until Mdbor4.EOF%>
<tr bordercolor="FFFFFF">
<td colspan="21" bordercolor="FFFFFF">
<a href=<%="report10.asp?" & Request.QueryString & "#vira" & fotnum%>>{<%=fotnum%>}&nbsp
<%=Mdbor4("footnote")%></a>
<%fotnum=fotnum+1%>
<%mdbor4.MoveNext%>
</td>
</tr>
<%Loop%>
<%mdbor4.Close%>

</Table>

</Body>
</Html>