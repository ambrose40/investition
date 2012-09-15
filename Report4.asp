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
<title>
InformatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on
</title>
</Head>
<body class="REPORT">

<img border="0" src="icons/report.ico" Style=float:Left><p align="center"><a href="Main.asp" target="_top" class="headlink">MAJANDUSAASTA INVESTEERINGUTE KAVA</a></p>
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
<Form Method="POST" Action="Report4.asp?ye=<%=ya%>">
<Input type="Submit" name="btn" size="10" Value="Kopeerimiseks" class="button">
<Input type="Submit" name="btn" size="10" Value="Parandamiseks" class="button">

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


<%set mdbo5 = Server.CreateObject("ADODB.Command")%>
<%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5.ActiveConnection = mdbo%>



<table border="1" width="100%" >
<tr>
 <th rowspan="2" width="67">Nr</td>
 <th rowspan="2" width="250">Projekti nimetus</td>
 <th rowspan="2" width="115">Ehitusaastad (m.a.)
 <th rowspan="2" width="117">Kalkuleeritud maksumus kokku</td>
 <th rowspan="2" width="92">L&otilde;petatud seisuga 01.04.<%=ya%></th>
 <th rowspan="2" width="50"><%=Mid(ya,3,2)%>&nbsp;m.a</th>
 <th colspan="4" rowspan="1" width="164">INVESTEERINGUD</th>
<th rowspan="2" width="170">Projektijuht</th>
</tr>
<tr>
 <th width="37">1 kv</th>
 <th width="36">2 kv</th>
 <th width="36">3 kv</th>
 <th width="37">4 kv</th> 
</tr>
<tr class="RepNum">
<%For nuu=1 to 11%>
 <td width="67"><%=nuu%></td>
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


<%mdbol1.CommandText="SELECT DISTINCT Pid,PC, OracleCode, PRojName FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='00' ORDER BY PC"%>
<%mdborl1.Open mdbol1%>

<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY FROM dbo.Main WHERE IDentifier='P' AND RenovBlock=0 AND SUBSTRING(ProjCode,10,2)<>'00' AND yearr='" & ya & "'"%>
<%mdbor5.Open mdbo5%>

<tr class="boldProjGrup">
<td width="250" colspan=2>
INVESTEERINGUD KOKKU v&auml;lja arvatud plokkide renoveerimine
</td>
<td width="115" colspan=3>
<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="a1a"%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a1a"%>">
<%Else%>
<%a0="a1a"%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="a1a"%>" class="boldProjGrup">
<%End If%>
</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>


</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>

<%mdbor5.Close%>
<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY FROM dbo.Main WHERE IDentifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND yearr='" & ya & "'"%>
<%mdbor5.Open mdbo5%>

<tr class="boldProjGrup">
<td width="250" colspan=2>
INVESTEERINGUD KOKKU koos plokkide renoveerimisega

</td>
<td width="115" colspan=3>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="aa"%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="aa"%>">
<%Else%>
<%a0="aa"%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="aa"%>" class="boldProjGrup">
<%End If%>

</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>


</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>
<%mdbor5.Close%>

<%Do until mdborl1.EOF%>
<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY, SUBSTRING(ProjCode,1,2) as be FROM dbo.Main WHERE SUBSTRING(ProjCode,1,2)='" & MID(mdborl1("PC"),1,2) & "' AND IDentifier='P' AND SUBSTRING(ProjCode,10,2)<>'00' AND yearr='" & ya & "' GROUP BY SUBSTRING(ProjCode,1,2)"%>
<%mdbor5.Open mdbo5%>
<tr class="ProjGrup">
<td width="67">
<%a=MID(mdborl1("PC"),1,3)%>
<%=REPLACE(a, "0", "")%>
</td>
<td width="250">
<%=mdborl1("ProjName")%>

</td>
<td width="115" colspan=3>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="aa" & mdborl1("Pid")%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" size="10" name="<%="aa" & mdborl1("Pid")%>">
<%Else%>
<%a0="aa" & mdborl1("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="aa" & mdborl1("Pid")%>" class="boldProjGrup">
<%End If%>

</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>

<%mdbol2.CommandText="SELECT DISTINCT Pid,PC,ProjName,OracleCode FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)<>'00' AND SUBSTRING(PC,7,2)='00' AND SUBSTRING(PC,10,2)<>'00' AND SUBSTRING(PC,1,2)='" & MID(mdborl1("PC"),1,2) & "' ORDER BY PC"%>
<%mdborl2.Open mdbol2%>


<%Do until mdborl2.EOF%>
<%mdbor5.Close%>

<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY, SUBSTRING(ProjCode,1,2) as be, SUBSTRING(ProjCode,4,2) as mi FROM dbo.Main WHERE SUBSTRING(ProjCode,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND SUBSTRING(ProjCode,10,2)<>'00' AND SUBSTRING(ProjCode,1,2)='" & MID(mdborl2("PC"),1,2) & "' AND IDentifier='P' AND yearr='" & ya & "' GROUP BY SUBSTRING(ProjCode,1,2), SUBSTRING(ProjCode,4,2)"%>
<%mdbor5.Open mdbo5%>

<tr class="ProjGrup">
<td width="67">
<%a=MID(mdborl2("PC"),1,6)%>
<%=REPLACE(a, "0", "")%>
</td>
<td width="250">
<%=mdborl2("ProjName")%>

</td>
<td width="115" colspan=3>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="aa" & mdborl2("Pid")%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="aa" & mdborl2("Pid")%>">
<%Else%>
<%a0="aa" & mdborl2("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="aa" & mdborl2("Pid")%>" class="ProjGrup">
<%End If%>

</td>

<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>

<%mdbol3.CommandText="SELECT DISTINCT Enterprise,Edescr FROM inpl WHERE IDentifier='C' AND Yearr=" & ya & " AND SUBSTRING(PC,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND  SUBSTRING(PC,1,2)='" & MID(mdborl2("PC"),1,2) & "'"%>
<%mdborl3.Open mdbol3%>


<%Do until mdborl3.EOF%>
<%mdbor5.Close%>

<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY, SUBSTRING(ProjCode,1,2) as be, SUBSTRING(ProjCode,4,2) as mi,Enterprise FROM dbo.Main WHERE SUBSTRING(ProjCode,4,2)='" & MID(mdborl2("PC"),4,2) & "' AND SUBSTRING(ProjCode,10,2)<>'00' AND SUBSTRING(ProjCode,1,2)='" & MID(mdborl2("PC"),1,2) & "' AND Enterprise='" & Mdborl3("Enterprise") & "' AND IDentifier='P' AND yearr='" & ya & "' GROUP BY SUBSTRING(ProjCode,1,2), SUBSTRING(ProjCode,4,2),Enterprise"%>
<%mdbor5.Open mdbo5%>

<tr class="enterp">

<td width="250" colspan=2>
<%=mdborl3("EDescr")%>

</td>
<td width="115" colspan=3>

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>">
<%Else%>
<%a0="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdborl2("Pid")%>" class="Enterp">
<%End If%>

</td>

<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>
<%mdbo1.CommandText="SELECT DISTINCT Pid,ProjCode as PC,RusName as ProjName,PastSum,SummYe,Ikvartal,IIkvartal,IIIkvartal,IVkvartal,EmplName,EmplFname FROM Aruanne WHERE Yearr='" & ya & "' AND Enterprise='" & mdborl3("Enterprise") & "' AND  par='" & MID(mdborl2("PC"),1,5) & "' ORDER BY ProjCode"%>
<%mdbor.Open mdbo1%>
<%mdbo2.CommandText="SELECT ROUND(SUM(GP.DEBET)/1000,0) AS summi, m.ProjCode as PC FROM dbo.Main AS m INNER JOIN dbo.glav_project AS GP ON m.OracleCode = GP.PROJECT WHERE m.Yearr='" & ya & "' AND LEFT(m.ProjCode,5)='" & MID(mdborl2("PC"),1,5) & "' AND m.Enterprise='" & mdborl3("Enterprise") & "' AND ((LEFT(MES,1)<='" & MID(ya-1,4,1) & "') OR (LEFT(MES,1)='" & MID(ya,4,1) & "' AND RIGHT(MES,1)<04) OR (LEFT(MES,1)='9')) AND (GP.DEBET IS NOT NULL) AND (m.IDentifier = 'C') GROUP BY m.ProjCode"%>
<%mdbor2.Open mdbo2%>
<%Do Until mdbor.EOF%>
<tr>
<td width="67">
<%a=REPLACE(mdbor("PC"), ".00", ".")%>
<%a=Mid(REPLACE(a, ".0", "."),2)%>
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
<td width="250" class=altActive title="<%=mdbo1.CommandText%>">
<%IF LEN(mdbor("PC"))>9 AND MID(mdbor("PC"),10,2)="00" then%>
<%=mdbor("ProjName")%> в том числе:

<%Else%>
<%=mdbor("ProjName")%>
<%End if%>

</td>
<td width="115">
<%If Request.Form("btn")="Kopeerimiseks" Then%>

<%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>">
<%Else%>
<%a0="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="aa" & mdborl3("Enterprise") & "_" & mdbor("Pid")%>" >
<%End If%>

</td>
<td width="117">

<%If mdbor2.EOF=False Then%>
<%If mdbor("PC")=mdbor2("PC") Then%>
<%=CDbl(mdbor2("Summi"))+CDbl(mdbor("SummYe"))%>
<%Else%>
<%=mdbor("SummYe")%>
<%End If%>
<%Else%>
<%=mdbor("SummYe")%>
<%End If%>

</td>

<td width="92">

<%If mdbor2.EOF=False Then%>
<%If mdbor("PC")=mdbor2("PC") Then%>
<%=mdbor("PastSum")%>
<%Else%>
0
<%End If%>
<%Else%>
0
<%End If%>

</td>

<td width="50">

<%=mdbor("SummYe")%>
</td>
<td width="37">

<%=mdbor("Ikvartal")%>

</td>
<td width="36">

<%=mdbor("IIkvartal")%>

</td>
<td width="36">

<%=mdbor("IIIkvartal")%>
</td>
<td width="37">

<%=mdbor("IVkvartal")%>

</td>
<td width="170">

<%=mdbor("EmplName")%>&nbsp<%=mdbor("EmplFName")%>

</td>
</tr>

<%mdbor.Movenext%>
<%If mdbor2.EOF = False Then%>
<%mdbor2.Movenext%>
<%Else%>
<%End If%>
<%loop%>

<%mdbor.Close%>
<%mdbor2.Close%>
<%mdborl3.Movenext%>
<%loop%>

<%mdborl3.Close%>
<%mdborl2.Movenext%>
<%loop%>

<%mdborl2.Close%>
<%mdborl1.Movenext%>
<%mdbor5.Close%>

<%Loop%>
<%mdborl1.Close%>
<tr>
<td colspan=11 class="bold">

Kokku ettev&otildetete kaupa

</td>
</tr>
<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY,Enterprise FROM dbo.Main WHERE IDentifier='P' AND yearr='" & ya & "' AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor5.Open mdbo5%>
<%mdbol1.CommandText="SELECT * FROM Enterprise ORDER BY Enterprise"%>
<%mdborl1.Open mdbol1%>
<%Do until Mdbor5.EOF%>
<tr class="boldEnterp">
<td width="67">



</td>
<td width="250">
<%=mdborl1("EDescr")%>

</td>
<td width="115">

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="ab" & mdborl1("Enterprise")%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="ab" & mdborl1("Enterprise")%>">
<%Else%>
<%a0="ab" & mdborl1("Enterprise")%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="ab" & mdborl1("Enterprise")%>"  class="boldEnterp" >
<%End If%>

</td>
<td width="117">
<td width="92">
</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>


</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>
<%mdbor5.Movenext%>
<%mdborl1.Movenext%>
<%Loop%>
<%mdbor5.Close%>

<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY FROM dbo.Main WHERE IDentifier='P' AND yearr='" & ya & "' AND SUBSTRING(ProjCode,10,2)<>'00'"%>
<%mdbor5.Open mdbo5%>

<tr  class="bold">
<td width="67">



</td>
<td width="250">
Kokku

</td>
<td width="115">

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="a3a"%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a3a"%>">
<%Else%>
<%a0="a3a"%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="a3a"%>" class="bold">
<%End If%>

</td>
<td width="117">
<td width="92">
</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>


</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>
<%mdbor5.Close%>
<%mdborl1.Close%>
<tr class="bold">
<td colspan="11" width="1067">

Kokku ettev&otildetete kaupa, v&auml;lja arvatud plokkide renoveerimine

</td>
</tr>
<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY,Enterprise FROM dbo.Main WHERE IDentifier='P' AND yearr='" & ya & "' AND RenovBlock=0 AND SUBSTRING(ProjCode,10,2)<>'00' GROUP BY Enterprise ORDER BY Enterprise"%>
<%mdbor5.Open mdbo5%>
<%mdbol1.CommandText="SELECT * FROM Enterprise ORDER BY Enterprise"%>
<%mdborl1.Open mdbol1%>
<%Do until Mdbor5.EOF%>
<tr class="boldEnterp">
<td width="67">



</td>
<td width="250">
<%=mdborl1("EDescr")%>

</td>
<td width="115">

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="a2b" & mdborl1("Enterprise")%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a2b" & mdborl1("Enterprise")%>">
<%Else%>
<%a0="a2b" & mdborl1("Enterprise")%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="a2b" & mdborl1("Enterprise")%>" class="boldEnterp">
<%End If%>

</td>
<td width="117">
<td width="92">
</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>


</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>
<%mdbor5.Movenext%>
<%mdborl1.Movenext%>
<%Loop%>
<%mdbor5.Close%>

<%mdbo5.CommandText="SELECT SUM(Ikvartal) AS S1, SUM(IIkvartal) AS S2, SUM(IIIkvartal) AS S3, SUM(IVkvartal) AS S4,SUM(SummYe) AS SY FROM dbo.Main WHERE IDentifier='P' AND RenovBlock=0 AND yearr='" & ya & "' AND SUBSTRING(ProjCode,10,2)<>'00'"%>
<%mdbor5.Open mdbo5%>

<tr class="bold">
<td width="67">



</td>
<td width="250">
Kokku

</td>
<td width="115">

<%If Request.Form("btn")="Kopeerimiseks" Then%>
<%a0="a4a"%>
<%=Request.Form(a0)%>
<input type="hidden" value="<%=Request.Form(a0)%>" name="<%="a4a"%>">
<%Else%>
<%a0="a4a"%>
<input type="Text" value="<%=Request.Form(a0)%>" size="10" name="<%="a4a"%>" class="bold">
<%End If%>

</td>
<td width="117">
<td width="92">
</td>
<td width="50">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("SY")%>
<%Else%>
0
<%End If%>


</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S1")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S2")%>
<%Else%>
0
<%End If%>

</td>
<td width="36">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S3")%>
<%Else%>
0
<%End If%>

</td>
<td width="37">

<%If mdbor5.BOF=False Then%>
<%=mdbor5("S4")%>
<%Else%>
0
<%End If%>

</td>
<td width="170">
</td>
</tr>
</Form>
</table>
</body>
</html>