<%Response.CacheControl="Public"%>
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
Invest-IT!on: INVESTEERIMISKAVA
</title>
<SCRIPT LANGUAGE="JavaScript">

function onKeyPress () {
var keycode;
if (window.event) keycode = window.event.keyCode;
else if (e) keycode = e.which;
else return true;
if (keycode == 13) {
newWindow ('Help.mht','','','','scrollbars');
return false
}
return true 
}
document.onkeypress = onKeyPress;

<!-- Begin
var win = null;
function newWindow(mypage,myname,w,h,features) {
  var winl = 0;
  var wint = 0;
  if (winl < 0) winl = 0;
  if (wint < 0) wint = 0;
  var settings = 'height=' + h + ',';
  settings += 'width=' + w + ',';
  settings += 'top=' + wint + ',';
  settings += 'left=' + winl + ',';
  settings += features;
  win = window.open(mypage,myname,settings);
  win.window.focus();
}
//  End -->
</script>
<SCRIPT LANGUAGE="VBScript">
Sub subme
  Dim TheForm
  Set TheForm = Document.forms("forma")
  TheForm.Submit
End Sub
</SCRIPT>
</Head>

<body background="icons/back.gif" class="Main">

<%If request.QueryString("y")="" Then%>
<%ya=Year(Date())%>
<%mo=Month(Date())%>
<%da=Day(Date())%>
<%zz=mo-04%>
<%If zz>=0 Then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=ya-1%>
<%End If%>
<%ia="Yearr='" & ya & "'"%>
<%Else%>
<%ia=request.QueryString("y")%>
<%End If%>

<%zzz=ya-1%>
<%zzz="Yearr='" & zzz & "'"%>
<%zzz2=ya+1%>
<%zzz2="Yearr='" & zzz2 & "'"%>

<%srt=Request.QueryString("sr")%>
<%If Request.form("ye")<> "" Then%>
<%zo="Yearr='" & Request.form("ye") & "'"%>
<%zzz=Request.form("ye")-1%>
<%zzz="Yearr='" & zzz & "'"%>
<%zzz2=Request.form("ye")+1%>
<%zzz2="Yearr='" & zzz2 & "'"%>
<%n=3%>
<%Else%>
<%zo=Request.QueryString("y")%>
<%n=Request.QueryString("no")%>
<%End If%>
<%If Request.QueryString("Entt")="" Then%>
<%np=Request.QueryString("e3")%>
<%Else%>

<%If Request.QueryString("Entt")="All" Then%>
<%np=""%>
<%Else%>
<%np="Enterprise='" & Request.QueryString("entt") & "'"%>
<%End If%>

<%End If%>

<%If Request.QueryString("entt")="" Then%>
<%co=Request.QueryString("s")%>
<%Else%>
<%co=Request.QueryString("s")%>
<%End If%>

<%If Request.QueryString("wrkk")="" Then%>
<%pb=Request.QueryString("em")%>
<%Else%>

<%If Request.QueryString("Wrkk")="All" Then%>
<%pb=""%>
<%Else%>
<%pb="EmployeeID='" & Request.QueryString("wrkk") & "'"%>
<%End If%>

<%End If%>

<%If zo="" Then%>
<%zzz=ya-1%>
<%zzz="Yearr='" & zzz & "'"%>
<%zzz2=ya+1%>
<%zzz2="Yearr='" & zzz2 & "'"%>
<%zo="Yearr='" & ya & "'"%>
<%Else%>
<%yo=Mid(zo,8,4)%>
<%zzz=yo-1%>
<%zzz="Yearr='" & zzz & "'"%>
<%zzz2=yo+1%>
<%zzz2="Yearr='" & zzz2 & "'"%>
<%End If%>



<a href="Invest.asp?&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zzz%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=Request.QueryString("so")%>"><Img border="0" src="icons/p.ico" Style=float:left></a><a href="Invest.asp?sr=<%=srt & dd%>&no=<%=n%>&y=<%=zzz2%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=Request.QueryString("so")%>"><Img border="0" src="icons/n.ico" Style=float:right></a>

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
<%set mdbo0 = Server.CreateObject("ADODB.Command")%>
<%set mdbor0 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo0.ActiveConnection = mdbo%>
                                                                                                                                                                                                                                                                                                                             

<%set mdbode = Server.CreateObject("ADODB.Command")%>
<%set mdborde = Server.CreateObject("ADODB.Recordset")%>
<%mdbode.ActiveConnection = mdbo%>
<%set mdbode3 = Server.CreateObject("ADODB.Command")%>
<%set mdborde3 = Server.CreateObject("ADODB.Recordset")%>
<%mdbode3.ActiveConnection = mdbo%>
<%If Request.QueryString("Del") = "1" AND Request.FORM("btn")<>"    Kinnita" Then%>
<%set mdbod = Server.CreateObject("ADODB.Command")%>
<%set mdbord = Server.CreateObject("ADODB.Recordset")%>
<%mdbod.ActiveConnection = mdbo%>
<%If MID(Request.Form("pid"),9,1)="." Then%>
<%piid=Mid(Request.Form("pid"),12,5)%>
<%pood=Mid(Request.Form("pid"),1,11)%>
<%Else%>
<%piid=MID(Request.Form("pid"),9,5)%>
<%pood=Mid(Request.Form("pid"),1,8)%>
<%End If%>
<%mdbod.CommandText="DELETE Codes WHERE Pid='" & piid & "'"%>
<%mdbord.Open mdbod%>
<%End If%>
<%If Request.Form("btn") = "    Kinnita" Then%>
<%set mdboap = Server.CreateObject("ADODB.Command")%>
<%set mdborap = Server.CreateObject("ADODB.Recordset")%>
<%mdboap.ActiveConnection = mdbo%>
<%set mdboo = Server.CreateObject("ADODB.Command")%>
<%set mdboro = Server.CreateObject("ADODB.Recordset")%>
<%b= Server.MapPath("/")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<%set servFileStream=servcfg.createTextFile(b & "/cookie.cfg")%>
<%yas=Request.Form("yra")%>
<%servFileStream.WriteLine yas%>
<%servFileStream.Close%><p>
<%set mdboym = Server.CreateObject("ADODB.Command")%>
<%set mdborym = Server.CreateObject("ADODB.Recordset")%>
<%mdboym.ActiveConnection = mdbo%>
<%set mdboy = Server.CreateObject("ADODB.Command")%>
<%set mdbory = Server.CreateObject("ADODB.Recordset")%>
<%mdboy.ActiveConnection = mdbo%>

<%mdboy.CommandText="SELECT DISTINCT Yearr FROM Main WHERE Pid ='" & piid & "'"%>
<%mdbory.Open mdboy%>


<%ftor=1%>
<%If MID(Request.Form("pid"),9,1)="." Then%>
<%piid=Mid(Request.Form("pid"),12,5)%>
<%pood=Mid(Request.Form("pid"),1,11)%>
<%Else%>
<%piid=MID(Request.Form("pid"),9,5)%>
<%pood=Mid(Request.Form("pid"),1,8)%>
<%End If%>
<%mdboym.CommandText="SELECT MAX(Yearr) as my FROM Main WHERE Pid ='" & piid & "'"%>
<%mdborym.Open mdboym%>

<%If mdborym.EOF or (mdborym("my") & "e")="e" then%>
<%ftor=1%>
<%Else%>
<%ftor=5%>
<%iii=mdborym("my")%>
<%End if%>

<%If ftor="5" Then%>
<%set mdboo = Server.CreateObject("ADODB.Command")%>
<%set mdboro = Server.CreateObject("ADODB.Recordset")%>
<%mdboo.ActiveConnection = mdbo%>
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='C' AND Yearr='" & iii & "' AND Pid ='" & piid & "' AND Enterprise='" & Request.Form("ena") & "'"%>
<%mdboro.Open mdboo%>
<%If mdboro.EOF="True" Then%>
<%mdboro.Close%>
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='C' AND Yearr='" & iii & "' AND Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>
<%End If%>
<%aa=mdboro("SummTot")%>
<%set mdboo = Server.CreateObject("ADODB.Command")%>
<%set mdboro = Server.CreateObject("ADODB.Recordset")%>
<%mdboo.ActiveConnection = mdbo%>
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='P' AND Yearr='" & iii & "' AND Pid='" & piid & "' AND Enterprise='" & Request.Form("ena") & "'"%>
<%mdboro.Open mdboo%>
<%If mdboro.EOF="True" Then%>
<%mdboro.Close%>
<%mdboo.CommandText="SELECT SummTot FROM Main WHERE Identifier='P' AND Yearr='" & iii & "' AND Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>
<%End If%>
<%bb=mdboro("SummTot")%>

<%set mdboo = Server.CreateObject("ADODB.Command")%>
<%set mdboro = Server.CreateObject("ADODB.Recordset")%>
<%mdboo.ActiveConnection = mdbo%>
<%mdboo.CommandText="SELECT SummTot,OracleCode,RusName FROM Main WHERE Identifier='F' AND Yearr='" & iii & "' AND Pid='" & piid & "' AND Enterprise='" & Request.Form("ena") & "'"%>
<%mdboro.Open mdboo%>
<%If mdboro.EOF="True" Then%>
<%mdboro.Close%>
<%mdboo.CommandText="SELECT SummTot,OracleCode,RusName FROM Main WHERE Identifier='F' AND Yearr='" & iii & "' AND Pid ='" & piid & "'"%>
<%mdboro.Open mdboo%>
<%End If%>
<%cc=mdboro("SummTot")%>
<%oco=mdboro("OracleCode")%>
<%dcp=mdboro("RusName")%>
<%Else%>
<%aa=0%>
<%bb=0%>
<%cc=0%>
<%oco="N/A"%>
<%dcp="---"%>
<%End If%>
<%mdboap.CommandText="INSERT INTO Main (ProjCode,Pid,Yearr,Enterprise,PastSum,IKvartal,IIkvartal,IIIKvartal,IVKvartal,Identifier,OracleCode,RusName) VALUES ('" & pood & "','" & piid & "', '" & Request.Form("yra") & "', '" & Request.Form("ena") & "'," & aa & ",0,0,0,0,'C','" & oco & "','" & dcp & "')"%>
<%mdborap.Open mdboap%>
<%mdboap.CommandText="INSERT INTO Main (ProjCode,Pid,Yearr,Enterprise,PastSum,IKvartal,IIkvartal,IIIKvartal,IVKvartal,Identifier,OracleCode,RusName) VALUES ('" & pood & "','" & piid & "', '" & Request.Form("yra") & "', '" & Request.Form("ena") & "'," & cc & ",0,0,0,0,'F','" & oco & "','" & dcp & "')"%>
<%mdborap.Open mdboap%>
<%mdboap.CommandText="INSERT INTO Main (ProjCode,Pid,Yearr,Enterprise,PastSum,IKvartal,IIkvartal,IIIKvartal,IVKvartal,Identifier,OracleCode,RusName) VALUES ('" & pood & "','" & piid & "', '" & Request.Form("yra") & "', '" & Request.Form("ena") & "'," & bb & ",0,0,0,0,'P','" & oco & "','" & dcp & "')"%>
<%mdborap.Open mdboap%>
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
<%End If%>

<%set mdbobi = Server.CreateObject("ADODB.Command")%>
<%set mdborbi = Server.CreateObject("ADODB.Recordset")%>
<%mdbobi.ActiveConnection = mdbo%>
<%set mdbou = Server.CreateObject("ADODB.Command")%>
<%set mdboru = Server.CreateObject("ADODB.Recordset")%>
<%mdbou.ActiveConnection = mdbo%>

<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>
<%set mdbo2P = Server.CreateObject("ADODB.Command")%>
<%set mdbor2P = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2P.ActiveConnection = mdbo%>

<%set mdbo3 = Server.CreateObject("ADODB.Command")%>
<%set mdbor3 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo3.ActiveConnection = mdbo%>
<%set mdbo3P = Server.CreateObject("ADODB.Command")%>
<%set mdbor3P = Server.CreateObject("ADODB.Recordset")%>
<%mdbo3P.ActiveConnection = mdbo%>

<%set mdbo4 = Server.CreateObject("ADODB.Command")%>
<%set mdbor4 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4.ActiveConnection = mdbo%>
<%set mdbo4P = Server.CreateObject("ADODB.Command")%>
<%set mdbor4P = Server.CreateObject("ADODB.Recordset")%>
<%mdbo4P.ActiveConnection = mdbo%>

<%set mdbo5 = Server.CreateObject("ADODB.Command")%>
<%set mdbor5 = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5.ActiveConnection = mdbo%>
<%set mdbo5P = Server.CreateObject("ADODB.Command")%>
<%set mdbor5P = Server.CreateObject("ADODB.Recordset")%>
<%mdbo5P.ActiveConnection = mdbo%>




<table border="1" style="border-collapse: collapse">
<tr>
<th class="inv">

No

</th>
<th class="inv">
<a href="invest.asp?so=<%=Request.QueryString("so")%>&y=<%=zo%>"class="th">*</a>
</th>

<%If request.QueryString("so")="all" Then%><th class="inv"colspan="1"><%Else%><th class="inv"colspan="2"><%End If%>
<form method="POST" action="insert.asp" target="_blanck">
<input type="Submit" class="inv" name="btnn" value="    " style="background-image:url('icons/insert.png'); background-repeat: no-repeat; background-position: center;" onmouseover='window.status="Avaneb lehek&uuml;lg proektide sisse panemiseks";'onmouseout='window.status="";'>
<%so=Request.QueryString("so")%>

<%If srt="" then%>

 <a href="invest.asp?sr=<%="ProjCode,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>" class="th">Proekti kood ja nimetus</a>

<%Else%>
 <%l=Len(srt)%>
 <%o=1%>
 <%j1=1%>
 <%f=2%>
 <%bz=""%>
 <%Do Until j1>l%>
  <%a=Mid(srt,j1,1)%>
  <%If a="," Then%>
   <%a2=Mid(srt,o,j1-o)%>
   <%o=j1+1%>
   <%If a2="ProjCode" Then%>
    <%bz=bz & a2 & " DESC,"%>
    <%j1=j1+1%>
    <%f=1%>
   <%Else%>
    <%bz=bz & a2 & ","%>
    <%j1=j1+1%>
   <%End If%>
  <%Else%>
   <%If a=" " Then%>
    <%a2=Mid(srt,o,j1-o)%>
    <%o=j1+6%>
    <%If a2="ProjCode" Then%>
     <%j1=j1+6%>
     <%f=4%>
    <%Else%>
     <%bz=bz & a2 & " DESC,"%>
    <%j1=j1+6%>
    <%End if%>
   <%Else%> 
    <%j1=j1+1%>
   <%End If%> 
  <%End If%>
 <%Loop%>
 <%If f=1 Then%>
  <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Proektide kood ja nimetus <img border="0" src="icons/down.png"></a>
 <%Else%>
  <%If f=4 Then%>
   <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Proektide kood ja nimetus <img border="0" src="icons/up.png"></a>
  <%Else%> 
   <a href="invest.asp?sr=<%=bz & "ProjCode,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Proektide kood ja nimetus </a>
  <%End If%>
 <%End If%>
<%End If%>
</th> 
</form>
<Form Method="PUT" ID="forma" class="inv Action="Invest.asp?sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>#<%="vira" & i8 & i9%>">

<th class="inv">
<input type="hidden" name="sr" value="<%=srt & dd%>">
<input type="hidden" name="no" value="<%=n%>">
<input type="hidden" name="y" value="<%=zo%>">
<input type="hidden" name="s" value="<%=co%>">
<input type="hidden" name="em" value="<%=pb%>">
<input type="hidden" name="e3" value="<%=np%>">
<input type="hidden" name="so" value="<%=so%>">
Oracle kood
</th>

<th class="inv">
 Aasta</a>
</th>
<%set mdboe = Server.CreateObject("ADODB.Command")%>
<%set mdbore = Server.CreateObject("ADODB.Recordset")%>
<%mdboe.ActiveConnection = mdbo%>
<%mdboe.CommandText="SELECT Enterprise,EDescr from Enterprise"%>
<%mdbore.Open mdboe%>


<%If srt="" then%>
 <th class="inv">
 <a href="invest.asp?sr=<%="Enterprise,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Ettev&otilde;te </a>
<select size="1" name="entt" class="inv" onChange="subme()">
<%mdbore.Movefirst%>
<%If np="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdbore.EOF%>
<%If "Enterprise='" & mdbore("Enterprise") & "'"=np Then%>
<option value="<%=mdbore("Enterprise")%>" selected="True"><%=mdbore("EDescr")%></option>
<%Else%>
<option value="<%=mdbore("Enterprise")%>"><%=mdbore("EDescr")%></option>
<%End If%>
<%mdbore.movenext%>
<%Loop%>
</select>
 </th>
<%Else%>
 <%l=Len(srt)%>
 <%bz=""%>
 <%o=1%>
 <%j1=1%>
 <%f=2%>
 <%Do Until j1>l%>
  <%a=Mid(srt,j1,1)%>
  <%If a="," Then%>
   <%a2=Mid(srt,o,j1-o)%>
   <%o=j1+1%>
   <%If a2="Enterprise" Then%>
    <%bz=bz & a2 & " DESC,"%>
    <%j1=j1+1%>
    <%f=1%>
   <%Else%>
    <%bz=bz & a2 & ","%>
    <%j1=j1+1%>
   <%End If%>
  <%Else%>
   <%If a=" " Then%>
    <%a2=Mid(srt,o,j1-o)%>
    <%o=j1+6%>
    <%If a2="Enterprise" Then%>
     <%j1=j1+6%>
     <%f=4%>
    <%Else%>
     <%j1=j1+6%>
     <%bz=bz & a2 & " DESC,"%>
    <%End if%>
   <%Else%> 
    <%j1=j1+1%>
   <%End If%> 
  <%End If%>
 <%Loop%>

 <%If f=1 Then%>
  <th class="inv">
  <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Ettev&otilde;te <img border="0" src="icons/down.png"></a>
<select size="1" name="entt" class="Inv" onChange="subme()">
<%mdbore.Movefirst%>
<%If np="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdbore.EOF%>
<%If "Enterprise='" & mdbore("Enterprise") & "'"=np Then%>
<option value="<%=mdbore("Enterprise")%>" selected="True"><%=mdbore("EDescr")%></option>
<%Else%>
<option value="<%=mdbore("Enterprise")%>"><%=mdbore("EDescr")%></option>
<%End If%>
<%mdbore.movenext%>
<%Loop%>
</select>
  </th> 
 <%Else%>
  <%If f=4 Then%>
   <th class="inv">
   <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Ettev&otilde;te <img border="0" src="icons/up.png"></a>
<select size="1" name="entt" class="inv" onChange="subme()">
<%mdbore.Movefirst%>
<%If np="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdbore.EOF%>
<%If "Enterprise='" & mdbore("Enterprise") & "'"=np Then%>
<option value="<%=mdbore("Enterprise")%>" selected="True"><%=mdbore("EDescr")%></option>
<%Else%>
<option value="<%=mdbore("Enterprise")%>"><%=mdbore("EDescr")%></option>
<%End If%>
<%mdbore.movenext%>
<%Loop%>
</select>
   </th>  
  <%Else%> 
  <th class="inv">
  <a href="invest.asp?sr=<%=bz & "Enterprise,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Ettev&otilde;te</a>
<select size="1" name="entt" class="inv" onChange="subme()">
<%mdbore.Movefirst%>
<%If np="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdbore.EOF%>
<%If "Enterprise='" & mdbore("Enterprise") & "'"=np Then%>
<option value="<%=mdbore("Enterprise")%>" selected="True"><%=mdbore("EDescr")%></option>
<%Else%>
<option value="<%=mdbore("Enterprise")%>"><%=mdbore("EDescr")%></option>
<%End If%>
<%mdbore.movenext%>
<%Loop%>
</select>
  </th>
  <%End If%>
 <%End If%>
<%End If%>


<%If srt="" then%>
 <th class="inv">
 <a href="invest.asp?sr=<%="SummYe,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Planeeritud summa</a>
 </th>
 <th class="inv">
 <a href="invest.asp?sr=<%="SummYe,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Lepingup&otilde;hine summa </a>
 </th>
<%Else%>
 <%l=Len(srt)%>
 <%bz=""%>
 <%o=1%>
 <%j1=1%>
 <%f=2%>
 <%Do Until j1>l%>
  <%a=Mid(srt,j1,1)%>
  <%If a="," Then%>
   <%a2=Mid(srt,o,j1-o)%>
   <%o=j1+1%>
   <%If a2="SummYe" Then%>
    <%bz=bz & a2 & " DESC,"%>
    <%j1=j1+1%>
    <%f=1%>
   <%Else%>
    <%bz=bz & a2 & ","%>
    <%j1=j1+1%>
   <%End If%>
  <%Else%>
   <%If a=" " Then%>
    <%a2=Mid(srt,o,j1-o)%>
    <%o=j1+6%>
    <%If a2="SummYe" Then%>
     <%j1=j1+6%>
     <%f=4%>
    <%Else%>
     <%j1=j1+6%>
     <%bz=bz & a2 & " DESC,"%>
    <%End if%>
   <%Else%> 
    <%j1=j1+1%>
   <%End If%> 
  <%End If%>
 <%Loop%>
 <%If f=1 Then%>
  <th class="inv">
  <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Planeeritud summa <img border="0" src="icons/down.png"></a>
  </th> 
  <th class="inv">
  <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Lepingup&otilde;hine summa <img border="0" src="icons/down.png"></a>
  </th> 
 <%Else%>
  <%If f=4 Then%>
   <th class="inv">
   <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Planeeritud summa <img border="0" src="icons/up.png"></a>
   </th>  
   <th class="inv">
   <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Lepingup&otilde;hine summa <img border="0" src="icons/up.png"></a>
   </th>  
  <%Else%> 
  <th class="inv">
  <a href="invest.asp?sr=<%=bz & "SummYe,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Planeeritud summa</a>
  </th>
  <th class="inv">
  <a href="invest.asp?sr=<%=bz & "SummYe,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Lepingup&otilde;hine summa </a>
  </th>
  <%End If%>
 <%End If%>
<%End If%>


<%If srt="" then%>
 <th class="inv">
 <a href="invest.asp?sr=<%="StatusId,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Seisund</a>
 </th>
<%Else%>
 <%l=Len(srt)%>
 <%bz=""%>
 <%o=1%>
 <%j1=1%>
 <%f=2%>
 <%Do Until j1>l%>
  <%a=Mid(srt,j1,1)%>
  <%If a="," Then%>
   <%a2=Mid(srt,o,j1-o)%>
   <%o=j1+1%>
   <%If a2="StatusId" Then%>
    <%bz=bz & a2 & " DESC,"%>
    <%j1=j1+1%>
    <%f=1%>
   <%Else%>
    <%bz=bz & a2 & ","%>
    <%j1=j1+1%>
   <%End If%>
  <%Else%>
   <%If a=" " Then%>
    <%a2=Mid(srt,o,j1-o)%>
    <%o=j1+6%>
    <%If a2="StatusId" Then%>
     <%j1=j1+6%>
     <%f=4%>
    <%Else%>
     <%j1=j1+6%>
     <%bz=bz & a2 & " DESC,"%>
    <%End if%>
   <%Else%> 
    <%j1=j1+1%>
   <%End If%> 
  <%End If%>
 <%Loop%>
 <%If f=1 Then%>
  <th class="inv">
  <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Seisund <img border="0" src="icons/down.png"></a>
  </th> 
 <%Else%>
  <%If f=4 Then%>
   <th class="inv">
   <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Seisund <img border="0" src="icons/up.png"></a>
   </th>  
  <%Else%> 
  <th class="inv">
  <a href="invest.asp?sr=<%=bz & "StatusId,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Seisund</a>
  </th>
  <%End If%>
 <%End If%>
<%End If%>

<%set mdbow = Server.CreateObject("ADODB.Command")%>
<%set mdborw = Server.CreateObject("ADODB.Recordset")%>
<%mdbow.ActiveConnection = mdbo%>
<%mdbow.CommandText="SELECT * from Worker ORDER BY EmplFName,Emplname"%>
<%mdborw.Open mdbow%>

<%If srt="" then%>
 <th class="inv">
 <a href="invest.asp?sr=<%="EmployeeId,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Vastutav t&ouml&oumltaja</a>

<select size="1" name="wrkk" class="inv" onChange="subme()">
<%mdborw.Movefirst%>
<%If pb="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdborw.EOF%>
<%If "EmployeeID='" & mdborw("EmployeeID") & "'"=pb Then%>
<option value="<%=mdborw("EmployeeID")%>" selected="True"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%Else%>
<option value="<%=mdborw("EmployeeID")%>"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%End If%>
<%mdborw.movenext%>
<%Loop%>
</select>
 </th>
<%Else%>
 <%l=Len(srt)%>
 <%bz=""%>
 <%o=1%>
 <%j1=1%>
 <%f=2%>
 <%Do Until j1>l%>
  <%a=Mid(srt,j1,1)%>
  <%If a="," Then%>
   <%a2=Mid(srt,o,j1-o)%>
   <%o=j1+1%>
   <%If a2="EmployeeId" Then%>
    <%bz=bz & a2 & " DESC,"%>
    <%j1=j1+1%>
    <%f=1%>
   <%Else%>
    <%bz=bz & a2 & ","%>
    <%j1=j1+1%>
   <%End If%>
  <%Else%>
   <%If a=" " Then%>
    <%a2=Mid(srt,o,j1-o)%>
    <%o=j1+6%>
    <%If a2="EmployeeId" Then%>
     <%j1=j1+6%>
     <%f=4%>
    <%Else%>
     <%j1=j1+6%>
     <%bz=bz & a2 & " DESC,"%>
    <%End if%>
   <%Else%> 
    <%j1=j1+1%>
   <%End If%> 
  <%End If%>
 <%Loop%>
 <%If f=1 Then%>
  <th class="inv">
  <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Vastutav t&ouml&oumltaja <img border="0" src="icons/down.png"></a>
<select size="1" name="wrkk" class="inv" onChange="subme()">
<%mdborw.Movefirst%>
<%If pb="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdborw.EOF%>
<%If "EmployeeID='" & mdborw("EmployeeID") & "'"=pb Then%>
<option value="<%=mdborw("EmployeeID")%>" selected="True"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%Else%>
<option value="<%=mdborw("EmployeeID")%>"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%End If%>
<%mdborw.movenext%>
<%Loop%>
</select>
  </th> 
 <%Else%>
  <%If f=4 Then%>
   <th class="inv">
   <a href="invest.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Vastutav t&ouml&oumltaja <img border="0" src="icons/up.png"></a>
<select size="1" name="wrkk" class="inv" onChange="subme()">
<%mdborw.Movefirst%>
<%If pb="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdborw.EOF%>
<%If "EmployeeID='" & mdborw("EmployeeID") & "'"=pb Then%>
<option value="<%=mdborw("EmployeeID")%>" selected="True"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%Else%>
<option value="<%=mdborw("EmployeeID")%>"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%End If%>
<%mdborw.movenext%>
<%Loop%>
</select>
   </th>  
  <%Else%> 
  <th class="inv">
  <a href="invest.asp?sr=<%=bz & "EmployeeId,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"class="th">Vastutav t&ouml&oumltaja</a>
<select size="1" name="wrkk" class="inv" onChange="subme()">
<%mdborw.Movefirst%>
<%If pb="" Then%>
<option value="All" selected="True">K&otilde;ik</option>
<%Else%>
<option value="All">K&otilde;ik</option>
<%End If%>
<%Do Until mdborw.EOF%>
<%If "EmployeeID='" & mdborw("EmployeeID") & "'"=pb Then%>
<option value="<%=mdborw("EmployeeID")%>" selected="True"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%Else%>
<option value="<%=mdborw("EmployeeID")%>"><%=mdborw("EmplFName")%>&nbsp<%=mdborw("EmplName")%></option>
<%End If%>
<%mdborw.movenext%>
<%Loop%>
</select>
  </th>
  <%End If%>
 <%End If%>
<%End If%>
</tr>
</form>

<%If srt<>"" Then%>
<%l2=len(srt)%>
<%srt=Mid(srt,1,l2-1)%>
<%End If%>
<%j=1%>
<%yg=zo & co & pb & np%>
<%Select Case yg%>
<%Case ""%>
<%y=""%>
<%Case zo%>
<%y=zo%>
<%ia=zo%>
<%Case np%>
<%y=np%>
<%Case co%>
<%y=co%>
<%Case pb%>
<%y=pb%>
<%Case zo & np%>
<%y=zo & " AND " & np%>
<%Case co & np%>
<%y=co & " AND " & np%>
<%Case pb & np%>
<%y=pb & " AND " & np%>
<%Case zo & pb%>
<%y=zo & " AND " & pb%>
<%Case zo & co%>
<%y=zo & " AND " & co%>
<%Case co & pb%>
<%y=co & " AND " & pb%>
<%Case zo & co & pb%>
<%y=zo & " AND " & co & " AND " & pb%>
<%Case zo & co & np%>
<%y=zo & " AND " & co & " AND " & np%>
<%Case zo & pb & np%>
<%y=zo & " AND " & pb & " AND " & np%>
<%Case co & pb & np%>
<%y=co & " AND " & pb & " AND " & np%>
<%Case zo & co & pb & np%>
<%y=zo & " AND " & co & " AND " & pb & " AND " & np%>
<%End Select%>
<%d=Year(Date()) & "-" & Month(Date()) & "-" & Day(Date())%>
<img border="0" src="icons/invpla.ico" Style=float:Left><p align="Center"><a href="Main.asp" class="headlink"><b>INVESTEERIMISKAVA <%=LEFT(RIGHT(ia,5),4)%> m.a.</b></a></p><p>
<%If request.QueryString("so")="all" Then%>
<%If srt="" Then%>
<%mdbo4.CommandText="SELECT * from inpl1 WHERE " & y & " AND Identifier='C' AND ProjCode NOT LIKE '__.00.00' AND ProjCode NOT LIKE '__.__.__.00' AND ProjCode NOT LIKE '__.__.00' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='1' OR StatusID='7')"%>
<%mdbor4.Open mdbo4%>
<%mdbo4P.CommandText="SELECT * from inpl1 WHERE " & y & " AND Identifier='P' AND ProjCode NOT LIKE '__.00.00' AND ProjCode NOT LIKE '__.__.__.00' AND ProjCode NOT LIKE '__.__.00' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='1' OR StatusID='7')"%>
<%mdbor4P.Open mdbo4P%>
<%Else%>
<%mdbo4.CommandText="SELECT * from inpl1 WHERE " & y & " AND Identifier='C' AND ProjCode NOT LIKE '__.00.00' AND ProjCode NOT LIKE '__.__.__.00' AND ProjCode NOT LIKE '__.__.00' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='1' OR StatusID='7') ORDER BY " & srt%>
<%mdbor4.Open mdbo4%>
<%mdbo4P.CommandText="SELECT * from inpl1 WHERE " & y & " AND Identifier='P' AND ProjCode NOT LIKE '__.00.00' AND ProjCode NOT LIKE '__.__.__.00' AND ProjCode NOT LIKE '__.__.00' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='1' OR StatusID='7') ORDER BY " & srt%>
<%mdbor4P.Open mdbo4P%>
<%End If%>
<%j=1%>
<%If srt="" Then%>
<%dd=""%>
<%Else%>
<%dd=","%>
<%End If%>

<%Do Until mdbor4.EOF%>

<tr>
<td class="inv" class="inv" colspan="1">
<%=j%>
</td>
<td class="inv" class="inv" colspan="1">
<%=mdbor4("Pid")%>
</td>
<td class="inv" class="inv" class='altActive' title="<%=mdbor4("RusName")%>">
<font style="font-size:9">
<%If len(mdbor4("ProjCode"))>9 and RIGHT(mdbor4("ProjCode"),2)="00" Then%>
<%=MId(mdbor4("ProjCode"),1,8)%>./<%=mdbor4("ProjName")%>
<%Else%>
<%ref="newWindow('ProjCard.asp?pc=" &mdbor4("Yearr") & mdbor4("Enterprise") & mdbor4("Pid") & "&sr=" & srt & dd & "&no=" & n & "&s=" & co & "&em=" & pb & "&so="& Request.QueryString("so") & "','','800','600','')"%>
<a style="font-size:9" href="#null" onClick="<%=ref%>"><%=mdbor4("ProjCode")%>./<%=mdbor4("ProjName")%></a>
<%End If%>
<br>
<font style="color:#000000">
<%=mdbor4("RusName")%>
</font>
</font>
</td>
<td class="inv" class="inv" class="inv">
<%=mdbor4("OracleCode")%>
</td>
<td class="inv" class="inv">
<%=mdbor4("Yearr")%>
</td>
<td class="inv" class="inv">
<a href="invest.asp?sr=<%=srt & dd%>&e3=<%="Enterprise='" & mdbor4("Enterprise") & "'"%>&no=<%="3"%>&so=<%=so%>&s=<%=co%>&em=<%=pb%>&y=<%=zo%>"><%=mdbor4("Edescr")%></a>
</td>
<td class="inv" class="inv">
<%=mdbor4P("SummYe")%>
<%sup=sup+Cdbl(mdbor4P("SummYe"))%>
</td>
<td class="inv" class="inv">
<%=mdbor4("SummYe")%>
<%suc=suc+Cdbl(mdbor4("SummYe"))%>
</td>
<td class="inv" class="inv">


<%If mdbor4("StatusId")="6" or mdbor4("StatusId")="7" Then%>
<img border="0" src="icons/stoi.png">
<%End If%>
<%a32="?" & mdbor4("StatusId")%>
<%If a32 = "?" Then%>
<img border="0" src="icons/que.png">
<%End If%>
<a href="invest.asp?sr=<%=srt & dd%>&s=<%="StatusID='" & mdbor4("StatusID") & "'"%>&no=<%="3"%>&so=<%=so%>&y=<%=zo%>&em=<%=pb%>&e3=<%=np%>"><%=mdbor4("StatusName")%></a>
</td>
<td class="inv" class="inv">

<%a33="?" & mdbor4("EmployeeID")%>
<%If a33 = "?" Then%>
<a href="invest.asp?sr=<%=srt & dd%>&em=<%="EmployeeID IS NULL "%>&so=<%=Request.QueryString("So")%>&no=<%="3"%>&y=<%=zo%>&s=<%=co%>&e3=<%=np%>"><%=mdbor4("EmplFName")%>&nbsp<%=mdbor4("EmplName")%></a>
<%Else%>
<a href="invest.asp?sr=<%=srt & dd%>&em=<%="EmployeeID='" & mdbor4("EmployeeID") & "'"%>&so=<%=Request.QueryString("So")%>&no=<%="3"%>&y=<%=zo%>&s=<%=co%>&e3=<%=np%>"><%=mdbor4("EmplFName")%>&nbsp<%=mdbor4("EmplName")%></a>
<%End If%>
</td>
</tr>


<%mdbor4.MoveNext%>
<%mdbor4P.MoveNext%>
<%j=j+1%>
<%loop%>
<%mdbor4.Close%>
<%mdbor4P.Close%>
<%Else%>
<%mdbo2.CommandText="SELECT DISTINCT Pid,ProjCode,ProjName,Yearr from inpl1 WHERE Identifier='C' AND " & ia & " AND ProjCode LIKE '__.00.00' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='7')"%>
<%mdbor.Open mdbo2%>
<%mdbode.CommandText="SELECT DISTINCT SUM(SummaPlan) AS SP, SUM(SummaContract) AS SC, be FROM dbo.Delta WHERE " & y & "  AND SUBSTRING(PRojCOde,10,2)<>'00' GROUP BY be"%>
<%mdborde.Open mdbode%>
<%m=1%>
<%If srt="" Then%>
<%dd=""%>
<%Else%>
<%dd=","%>
<%End If%>

<Form Method="POST" Action="Invest.asp?sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>#<%="vira" & i8 & i9%>">
<%i9=1%>
<%Do Until mdbor.EOF%>
<trBBDDFF>
<td class="inv" class="inv">
<%=j%>
</td>
<td class="inv" onmouseover='window.status="Expands/Minimizes Project Group.";'onmouseout='window.status="";'>
<%a="bpm" & mdbor("ProjCode") & mdbor("Yearr")%>
<%b="vpm" & mdbor("ProjCode") & mdbor("Yearr")%>
<%c="vpc" & mdbor("ProjCode") & mdbor("Yearr")%>
<%If request.Form(a)="" and request.Form(b)="" and request.Form(c)="" then%>
<input type="Submit" value="+" name="<%=a%>" class="plu">
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="-" name="<%=b%>">
<%Else%>
<%'=request.Form(a)%>
<%If request.Form(a)="" AND request.Form(b)<>"" then%>
<%'=request.Form(b)%>
<%If Request.Form(b)="+" Then%>
<%z="-"%>
<%End If%>
<%If Request.Form(b)="-" Then%>
<%z="+"%>
<%End If%>
<%'=z%>
<%If z="+" Then%>
<Input type="Submit" value="<%=z%>" name="<%=a%>" class="plu">
<%Else%>
<Input type="Submit" value="<%=z%>" name="<%=a%>"  class="min" >
<%End If%>
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="<%=Request.Form(b)%>" name="<%=b%>">
<%Else%>
<%If request.Form(a)="+" then%>
<a name="vira"></a>
<Input type="Submit" value="-" name="<%=a%>"  class="min" >
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="+" name="<%=b%>">
<%Else%>
<%If request.Form(a)="-" then%>
<a name="vira"></a>
<Input type="Submit" value="+" name="<%=a%>" class="plu">
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="-" name="<%=b%>">
<%Else%>
<%End If%>
<%End If%>
<%End If%>
<%End If%>
<font style="font-size:8"><%=mdbor("Pid")%></font>
</td>
<td class="inv" colspan="5">
<%=Mid(mdbor("ProjCode"),1,3)%>/<%=mdbor("ProjName")%>
</td>
<td class="inv" colspan="1">

<%If mdborde.EOF<>True Then%>
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("SP")%>
<%sup=sup+CDbl(mdborde("SP"))%>
<%Else%>
0
<%End If%>
<%Else%>
0
<%End If%>
</td>
<td class="inv" colspan="1">

<%If mdborde.EOF<>True Then%>
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("SC")%>
<%suc=suc+Cdbl(mdborde("SC"))%>
<%Else%>
0
<%End If%>
<%Else%>
0
<%End If%>
</td>
</td>

<td class="inv" colspan="2">

------------
</td>
</tr>
<%c="vpc" & mdbor("ProjCode") &  mdbor("Yearr")%>

<%If Request.Form(a)="+" or (Request.Form(a)="" and Request.Form(b)="+") then%>

<%p7=mid(Request.Form(c),1,8)%>
<%p8=mid(Request.Form(c),1,2)%>
<%mdbo3.CommandText="SELECT DISTINCT Pid,ProjCode,ProjName,Yearr from inpl1 WHERE "& ia & " AND Identifier='C' AND ProjCode LIKE '" & p8 & ".__.00' AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='7')"%>
<%mdbor3.Open mdbo3%>
<%mdbode3.CommandText="SELECT DISTINCT SUM(SummaPlan) AS SP, SUM(SummaContract) AS SC, be, mi FROM dbo.Delta WHERE " & y &" AND (mi <> '00') AND enn<>'00' AND be='" & p8 & "' GROUP BY be, mi"%>
<%mdborde3.Open mdbode3%>
<%i8=1%>
<%Do Until mdbor3.EOF%>
<%j=j+1%>

<trBBDDFF>
<td class="inv" class="inv">


<%=j%>
</td>
<%a2=a & "bpm2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%b2=b & "vpm2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%c2=c & "vpc2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%m=m+1%>
<td class="inv" onmouseover='window.status="Expands/Minimizes Project SubGroup.";'onmouseout='window.status="";'>
<%If request.Form(a2)="" and request.Form(b2)="" and request.Form(c2)="" then%>
<input type="Submit"  value="+" name="<%=a2%>" class="plu" >
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="-" name="<%=b2%>">
<%Else%>
<%'=request.Form(a2)%>
<%If request.Form(a2)="" AND request.Form(b2)<>"" then%>
<%'=request.Form(b2)%>
<%If Request.Form(b2)="+" Then%>
<%z="-"%>
<%End If%>
<%If Request.Form(b2)="-" Then%>
<%z="+"%>
<%End If%>
<%'=z%>
<%If z="+" Then%>
<Input type="Submit" value="<%=z%>" name="<%=a2%>"  class="plu" >
<%Else%>
<Input type="Submit" value="<%=z%>" name="<%=a2%>"  class="min" >
<%End If%>
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="<%=Request.Form(b2)%>" name="<%=b2%>">
<%Else%>
<%If request.Form(a2)="+" then%>
<a name="vira"></a>
<Input type="Submit" value="-" name="<%=a2%>"  class="min" >
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="+" name="<%=b2%>">
<%Else%>
<%If request.Form(a2)="-" then%>
<a name="vira"></a>
<Input type="Submit" value="+" name="<%=a2%>"  class="plu" >
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="-" name="<%=b2%>">
<%Else%>
<%End If%>
<%End If%>
<%End If%>
<%End If%>
<font style="font-size:8"><%=mdbor3("Pid")%></font>
</td>
<td class="inv" colspan="5">

&nbsp<img border="0" src="icons/lev.png">&nbsp<%=Mid(mdbor3("ProjCode"),1,6)%>/<%=mdbor3("ProjName")%>
</td>
<td class="inv" colspan="1">

<%If mdborde3.EOF="False" Then%>
<%If mdborde3("be") & "." & mdborde3("mi")=Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("SP")%>
<%Else%>
0
<%End If%>
<%Else%>
0
<%End If%>
</td>
<td class="inv" colspan="1">

<%If mdborde3.EOF="False" Then%>
<%If mdborde3("be") & "." & mdborde3("mi")=Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("SC")%>
<%Else%>
0
<%End If%>
<%Else%>
0
<%End If%>
</td>
<td class="inv" colspan="2">
------------
</td>
</tr>

<%c2=c & "vpc2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%If Request.Form(a2)="+" or (Request.Form(a2)="" and Request.Form(b2)="+") then%>
<%p7=mid(Request.Form(c2),1,8)%>
<%p8=mid(Request.Form(c2),1,5)%>
<%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
<%If y="" Then%>
<%If srt="" then%>
<%mdbo4.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='C' AND(ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor4.Open mdbo4%>
<%mdbo4P.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor4P.Open mdbo4P%>
<%Else%>
<%mdbo4.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor4.Open mdbo4%>
<%mdbo4P.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor4P.Open mdbo4P%>
<%End If%>
<%Else%>
<%If srt="" then%>
<%mdbo4.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor4.Open mdbo4%>
<%mdbo4P.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor4P.Open mdbo4P%>
<%Else%>
<%mdbo4.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor4.Open mdbo4%>
<%mdbo4P.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".___' OR ProjCode LIKE '" & p8 & ".__' OR ProjCode LIKE '" & p8 & ".__.00') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor4P.Open mdbo4P%>
<%End If%>
<%End If%>

<%Do Until mdbor4.EOF%>
<%j=j+1%>
<trBBDDFF>
<td class="inv" class="inv">

<%=j%>
</td>
<td class="inv" class="inv">
<%m=m+1%>
<%a3=a2 & "bpm" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("StatusID") & mdbor4("EmployeeID") & mdbor4("Enterprise")%>
<%b3=b2 & "vpm" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("StatusID") & mdbor4("EmployeeID") & mdbor4("Enterprise")%>
<%c3=c2 & "vpc" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("StatusID") & mdbor4("EmployeeID") & mdbor4("Enterprise")%>

<%If request.Form(a3)="" and request.Form(b3)="" and request.Form(c3)="" and Len(mdbor4("ProjCode"))>9 then%>
<input type="Submit" value="+" name="<%=a3%>"   class="plu" >
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="-" name="<%=b3%>">
<%Else%>
<%'=request.Form(a3)%>
<%If request.Form(a3)="" AND request.Form(b3)<>"" and Len(mdbor4("ProjCode"))>9 then%>
<%'=request.Form(b3)%>
<%If Request.Form(b3)="+" Then%>
<%z="-"%>
<%End If%>
<%If Request.Form(b3)="-" Then%>
<%z="+"%>
<%End If%>
<%'=z%>
<%If z="+" Then%>
<Input type="Submit" value="<%=z%>" name="<%=a3%>"  class="plu" >
<%Else%>
<Input type="Submit" value="<%=z%>" name="<%=a3%>"  class="min" >
<%End If%>
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="<%=Request.Form(b3)%>" name="<%=b3%>">
<%Else%>
<%If request.Form(a3)="+" and Len(mdbor4("ProjCode"))>9 then%>
<a name="vira"></a>
<Input type="Submit" value="-" name="<%=a3%>" class="min" >
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="+" name="<%=b3%>">
<%Else%>
<%If request.Form(a3)="-" and Len(mdbor4("ProjCode"))>9 then%>
<a name="vira"></a>
<Input type="Submit" value="+" name="<%=a3%>"   class="plu" >
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="-" name="<%=b3%>">
<%Else%>
<%End If%>
<%End If%>
<%End If%>
<%End If%>
<%'If Len(mdbor4("ProjCode"))<9 Then%>
<%=mdbor4("PID")%>
<%'End If%>
</td>
<td class="inv" class="inv">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<img border="0" src="icons/lev.png">
</td>
<td class="inv" class='altActive' title='<%=mdbor4("RusName")%>'>

<%If Len(mdbor4("ProjCode"))>9 and MID(mdbor4("ProjCode"),10,2)="00" Then%>
<%mdbobi.CommandText="SELECT SUM(Ikvartal) as s1,SUM(IIkvartal) as s2, SUM(IIIkvartal) as s3,SUM(IVkvartal) as s4 FROM Main WHERE ProjCode LIKE '" & MID(mdbor4("ProjCode"),1,9) & "__' AND ProjCode <> '" & mdbor4("ProjCode") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='F'"%>
<%mdborbi.Open mdbobi%>
<%If Mdborbi("s1") & "e" <> "e" then%>
<%sf1=Mdborbi("s1")%>
<%else%>
<%sf1=0%>
<%End If%>
<%If Mdborbi("s2")<>"" then%>
<%sf2=Mdborbi("s2")%>
<%Else%>
<%sf2=0%>
<%End If%>
<%If Mdborbi("s3")<>"" then%>
<%sf3=Mdborbi("s3")%>
<%Else%>
<%sf3=0%>
<%End If%>
<%If Mdborbi("s4")<>"" then%>
<%sf4=Mdborbi("s4")%>
<%Else%>
<%sf4=0%>
<%End If%>

<%mdbou.CommandText="UPDATE Main SET Ikvartal=" & sf1 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IIkvartal=" & sf2 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IIIkvartal=" & sf3 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IVkvartal=" & sf4 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='F'"%>
<%mdboru.Open mdbou%>
<%mdborbi.Close%>
<%mdbobi.CommandText="SELECT SUM(Ikvartal) as s1,SUM(IIkvartal) as s2, SUM(IIIkvartal) as s3,SUM(IVkvartal) as s4 FROM Main WHERE ProjCode LIKE '" & MID(mdbor4("ProjCode"),1,9) & "__' AND ProjCode <> '" & mdbor4("ProjCode") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='C'"%>
<%mdborbi.Open mdbobi%>
<%If Mdborbi("s1")<>"" then%>
<%sc1=Mdborbi("s1")%>
<%else%>
<%sc1=0%>
<%End If%>
<%If Mdborbi("s2")<>"" then%>
<%sc2=Mdborbi("s2")%>
<%Else%>
<%sc2=0%>
<%End If%>
<%If Mdborbi("s3")<>"" then%>
<%sc3=Mdborbi("s3")%>
<%Else%>
<%sc3=0%>
<%End If%>
<%If Mdborbi("s4")<>"" then%>
<%sc4=Mdborbi("s4")%>
<%Else%>
<%sc4=0%>
<%End If%>
<%mdbou.CommandText="UPDATE Main SET Ikvartal=" & sc1 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='C'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IIkvartal=" & sc2 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='C'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IIIkvartal=" & sc3 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='C'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IVkvartal=" & sc4 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='C'"%>
<%mdboru.Open mdbou%>
<%mdborbi.Close%>
<%mdbobi.CommandText="SELECT SUM(Ikvartal) as s1,SUM(IIkvartal) as s2, SUM(IIIkvartal) as s3,SUM(IVkvartal) as s4 FROM Main WHERE ProjCode LIKE '" & MID(mdbor4("ProjCode"),1,9) & "__' AND ProjCode <> '" & mdbor4("ProjCode") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='P'"%>
<%mdborbi.Open mdbobi%>
<%If Mdborbi("s1")<>"" then%>
<%sp1=Mdborbi("s1")%>
<%else%>
<%sp1=0%>
<%End If%>
<%If Mdborbi("s2")<>"" then%>
<%sp2=Mdborbi("s2")%>
<%Else%>
<%sp2=0%>
<%End If%>
<%If Mdborbi("s3")<>"" then%>
<%sp3=Mdborbi("s3")%>
<%Else%>
<%sp3=0%>
<%End If%>
<%If Mdborbi("s4")<>"" then%>
<%sp4=Mdborbi("s4")%>
<%Else%>
<%sp4=0%>
<%End If%>
<%mdbou.CommandText="UPDATE Main SET Ikvartal=" & sp1 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='P'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IIkvartal=" & sp2 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='P'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IIIkvartal=" & sp3 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='P'"%>
<%mdboru.Open mdbou%>
<%mdbou.CommandText="UPDATE Main SET IVkvartal=" & sp4 & " WHERE Pid='" & mdbor4("Pid") & "' AND Enterprise='" & mdbor4("Enterprise") & "' AND Yearr='" & mdbor4("Yearr") & "' AND Identifier='P'"%>
<%mdboru.Open mdbou%>
<%mdborbi.Close%>
<%End If%>
<font style="font-size:9">
<%If len(mdbor4("ProjCode"))>9  Then%>
<%=MId(mdbor4("ProjCode"),1,8)%>./<%=mdbor4("ProjName")%>
<%Else%>
<%ref="newWindow('ProjCard.asp?pc=" &mdbor4("Yearr") & mdbor4("Enterprise") & mdbor4("Pid") & "&sr=" & srt & dd & "&no=" & n & "&s=" & co & "&em=" & pb & "&so="& Request.QueryString("so") & "','','800','600','')"%>
<a style="font-size:9" href="#null" onClick="<%=ref%>"><%=mdbor4("ProjCode")%>./<%=mdbor4("ProjName")%></a>
<%End If%>
<br>
<font style="color:#000000;font-size:9;">
<%=mdbor4("RusName")%>
</font>
</font>
</td>
<td class="inv" class="inv">
<%=mdbor4("OracleCode")%>
</td>
<td class="inv" class="inv">
<%=mdbor4("Yearr")%>
</td>
<td class="inv" class="inv">
<a href="invest.asp?sr=<%=srt & dd%>&e3=<%="Enterprise='" & mdbor4("Enterprise") & "'"%>&no=<%="3"%>&s=<%=co%>&em=<%=pb%>&y=<%=zo%>&so=<%=Request.QueryString("So")%>"><%=mdbor4("Edescr")%></a>
</td>
<td class="inv" class="inv">

<%If Len(mdbor4("ProjCode"))>9 and MID(mdbor4("ProjCode"),10,2)="00" Then%>
<%=Cdbl(sp1)+Cdbl(sp2)+Cdbl(sp3)+Cdbl(sp4)%>
<%Else%>
<%=mdbor4P("SummYe")%>
<%End If%>
<%If mdbor4P("SummYe")<>"" Then%>
<%h5=mdbor4P("ProjCode") & mdbor4P("Enterprise")%>
<%If h5<>h6 Then%>
<%If Len(mdbor4("ProjCode"))>9 and MID(mdbor4("ProjCode"),10,2)="00" Then%>
<%sup2=sup2+Cdbl(sp1)+Cdbl(sp2)+Cdbl(sp3)+Cdbl(sp4)%>
<%Else%>
<%sup2=sup2+Cdbl(mdbor4P("SummYe"))%>
<%End If%>
<%End If%>
<%h6=mdbor4P("ProjCode") & mdbor4P("Enterprise")%>
<%End If%>
<%g1=g1+1%>
</td>
<td class="inv" class="inv">

<%If Len(mdbor4("ProjCode"))>9 and MID(mdbor4("ProjCode"),10,2)="00" Then%>
<%=Cdbl(sc1)+Cdbl(sc2)+Cdbl(sc3)+Cdbl(sc4)%>
<%Else%>
<%=mdbor4("SummYe")%>
<%End If%>
<%If mdbor4("SummYe")<>"" Then%>
<%f5=mdbor4("ProjCode") & mdbor4("Enterprise")%>
<%If f5<>f6 Then%>
<%If Len(mdbor4("ProjCode"))>9 and MID(mdbor4("ProjCode"),10,2)="00" Then%>
<%suc2=suc2+Cdbl(sc1)+Cdbl(sc2)+Cdbl(sc3)+Cdbl(sc4)%>
<%Else%>
<%suc2=suc2+Cdbl(mdbor4("SummYe"))%>
<%End If%>
<%End If%>
<%f6=mdbor4("ProjCode") & mdbor4("Enterprise")%>
<%End If%>
</td>
<td class="inv" class="inv">


<%If mdbor4("StatusId")="6" or mdbor4("StatusId")="7" Then%>
<img border="0" src="icons/stoi.png">
<%End If%>
<%a32="?" & mdbor4("StatusId")%>
<%If a32 = "?" Then%>
<img border="0" src="icons/que.png">
<%End If%>

<a href="invest.asp?sr=<%=srt & dd%>&s=<%="StatusID='" & mdbor4("StatusID") & "'"%>&no=<%="3"%>&y=<%=zo%>&em=<%=pb%>&e3=<%=np%>&so=<%=Request.QueryString("So")%>"><%=mdbor4("StatusName")%></a>
</td>
<td class="inv" class="inv">

<%a33="?" & mdbor4("EmployeeID")%>
<%If a33 = "?" Then%>
<a href="invest.asp?sr=<%=srt & dd%>&em=<%="EmployeeID IS NULL "%>&no=<%="3"%>&y=<%=zo%>&s=<%=co%>&e3=<%=np%>&so=<%=Request.QueryString("So")%>"><%=mdbor4("EmplFName")%>&nbsp<%=mdbor4("EmplName")%></a>
<%Else%>
<a href="invest.asp?sr=<%=srt & dd%>&em=<%="EmployeeID='" & mdbor4("EmployeeID") & "'"%>&no=<%="3"%>&y=<%=zo%>&s=<%=co%>&e3=<%=np%>&so=<%=Request.QueryString("So")%>"><%=mdbor4("EmplFName")%>&nbsp<%=mdbor4("EmplName")%></a>
<%End If%>

</td>
</tr>

<%c3=c2 & "vpc" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("StatusID") & mdbor4("EmployeeID") & mdbor4("Enterprise")%>
<%If Request.Form(a3)="+" or (Request.Form(a3)="" and Request.Form(b3)="+") then%>
<%p7=mid(Request.Form(c3),1,12)%>
<%p8=mid(Request.Form(c3),1,8)%>

<%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
<%If y="" Then%>
<%If srt="" then%>
<%mdbo5.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor5.Open mdbo5%>
<%mdbo5P.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor5P.Open mdbo5P%>
<%Else%>
<%mdbo5.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor5.Open mdbo5%>
<%mdbo5P.CommandText="SELECT * from inpl1 WHERE " & ia & " AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor5P.Open mdbo5P%>
<%End If%>
<%Else%>
<%If srt="" then%>
<%mdbo5.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor5.Open mdbo5%>
<%mdbo5P.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1')"%>
<%mdbor5P.Open mdbo5P%>
<%Else%>
<%mdbo5.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='C' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor5.Open mdbo5%>
<%mdbo5P.CommandText="SELECT * from inpl1 WHERE " & y & "AND Identifier='P' AND (ProjCode LIKE '" & p8 & ".__') AND ProjCode <> '" & p7 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='1') ORDER BY " & srt%>
<%mdbor5P.Open mdbo5P%>
<%End If%>
<%End If%>

<%Do Until mdbor5.EOF%>
<%j=j+1%>
<trBBDDFF>
<td class="inv" class="inv">

<%=j%>

</td>
<td class="inv" class="inv">
<%=mdbor5("PID")%>
</td>
<%m=m+1%>
<td class="inv" class="inv">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<img border="0" src="icons/lev.png"></td>
<td class="inv" class='altActive' title="<%=mdbor5("RusName")%>">

<font style="font-size:9">
<%ref="newWindow('ProjCard.asp?pc=" & mdbor5("Yearr") & mdbor5("Enterprise") & mdbor5("Pid") & "&sr=" & srt & dd & "&no=" & n & "&s=" & co & "&em=" & pb & "&e3=" & np & "&so="& Request.QueryString("so") & "','','800','600','')"%>
<a style="font-size:9"  href="#null" onClick="<%=ref%>"><%=mdbor5("ProjCode")%>./<%=mdbor5("ProjName")%></a>
<br>
<font style="color:#000000">
<%=mdbor5("RusName")%>
</font>
</font>

</td>
<td class="inv" class="inv">

<%=mdbor5("OracleCode")%>

</td>
<td class="inv" class="inv">

<%=mdbor5("Yearr")%>

</td>
<td class="inv" class="inv">
<a href="invest.asp?sr=<%=srt & dd%>&e3=<%="Enterprise='" & mdbor5("Enterprise") & "'"%>&no=<%="3"%>&s=<%=co%>&em=<%=pb%>&y=<%=zo%>&so=<%=Request.QueryString("So")%>"><%=mdbor5("Edescr")%></a>
</td>
<td class="inv" class="inv">

<%=mdbor5P("SummYe")%>
<%If mdbor5P("SummYe")<>"" Then%>
<%h5=mdbor5P("ProjCode") & mdbor5P("Enterprise")%>
<%If h5<>h6 Then%>
<%sup2=sup2+CDbl(mdbor5P("SummYe"))%>
<%End If%>
<%h6=mdbor5P("ProjCode") & mdbor5P("Enterprise")%>
<%End If%>
<%g1=g1+1%>
</td>
<td class="inv" class="inv">

<%=mdbor5("SummYe")%>
<%If mdbor5("SummYe")<>"" Then%>
<%f5=mdbor5("ProjCode") & mdbor5("Enterprise")%>
<%If f5<>f6 Then%>
<%suc2=suc2+Cdbl(mdbor5("SummYe"))%>
<%End If%>
<%f6=mdbor5("ProjCode") & mdbor5("Enterprise")%>
<%End If%>
</td>
<td class="inv" class="inv">


<%If mdbor5("StatusId")="6" or mdbor5("StatusId")="7" Then%>
<img border="0" src="icons/stoi.png">
<%End If%>
<%a32="?" & mdbor5("StatusId")%>
<%If a32 = "?" Then%>
<img border="0" src="icons/que.png">
<%End If%>

<a href="invest.asp?sr=<%=srt & dd%>&s=<%="StatusID='" & mdbor5("StatusID") & "'"%>&no=<%="3"%>&y=<%=zo%>&em=<%=pb%>&e3=<%=np%>&so=<%=Request.QueryString("So")%>"><%=mdbor5("StatusName")%></a>
</td>
<td class="inv" class="inv">

<%a33="?" & mdbor5("EmployeeID")%>
<%If a33 = "?" Then%>
<a href="invest.asp?sr=<%=srt & dd%>&em=<%="EmployeeID IS NULL "%>&no=<%="3"%>&y=<%=zo%>&s=<%=co%>&e3=<%=np%>&so=<%=Request.QueryString("So")%>"><%=mdbor5("EmplFName")%>&nbsp<%=mdbor5("EmplName")%></a>
<%Else%>
<a href="invest.asp?sr=<%=srt & dd%>&em=<%="EmployeeID='" & mdbor5("EmployeeID") & "'"%>&no=<%="3"%>&y=<%=zo%>&s=<%=co%>&e3=<%=np%>&so=<%=Request.QueryString("So")%>"><%=mdbor5("EmplFName")%>&nbsp<%=mdbor5("EmplName")%></a>
<%End If%>

</td>
</tr>
<%mdbor5.MoveNext%>
<%mdbor5P.MoveNext%>
<%loop%>
<%mdbor5.Close%>
<%mdbor5P.Close%>
<%End If%>
<%mdbor4.MoveNext%>
<%mdbor4P.MoveNext%>
<%loop%>
<%mdbor4.Close%>
<%mdbor4P.Close%>
<%End If%>
<%If mdborde3.BOF="True" Then%>
<%Else%>
<%If mdborde3.EOF="True" Then%>
<%mdborde3.MoveFirst%>
<%End If%>
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%mdborde3.MoveNext%>
<%End If%>
<%End If%>
<%mdbor3.MoveNext%>
<%i8=i8+1%>
<%loop%>
<%mdborde3.Close%>
<%mdbor3.Close%>
<%End If%>

<%j=j+1%>
<%If mdborde.EOF<>"True" Then%>
<%If (mdborde("be")=Mid(mdbor("ProjCode"),1,2)) Then%>
<%mdborde.MoveNext%>
<%End If%>
<%End If%>
<%i9=i9+1%>
<%mdbor.MoveNext%>
<%loop%>
<%End If%>
<tr class="sin" >

<td class="sin" >
<%If Request.QueryString("so")="all" Then%>
<a href="Invest.asp?so=group&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><font size="2" Color="000000"><Img border="0" src="icons/summ.png"></a>
<%Else%>
<a href="Invest.asp?so=all&sr=<%=srt & dd%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>"><font size="2" Color="000000"><Img border="0" src="icons/list.png"></a>
<%End If%>
</td>
<td class="sin" >
<a href="Control.asp?sr=<%=srt & dd%>&y=<%=zo%>&e3=<%=np%>">X</a>
</td>
<%If request.QueryString("so")="all" Then%><td colspan="2" class="sin" ><%Else%>
<td colspan="3" class="sin"><%End If%>
K&otilde;ikidele Projekti Gruppidele ette n&auml;htud summa
</td>
<td class="sin" >
-
</td>
<td class="sin" >
<a href="Invest.asp?sr=<%=srt & dd%>&no=<%=""%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=""%>&so=<%=so%>">K&otildeik</a></td>
<td class="sin" >
<%=sup%></td>
<td class="sin" >
<%=suc%></td>
<td class="sin" >
<a href="Invest.asp?sr=<%=srt & dd%>&no=<%=""%>&y=<%=zo%>&s=<%=""%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>">K&otildeik</a></td>
<td class="sin" >
<a href="Invest.asp?sr=<%=srt & dd%>&no=<%=""%>&y=<%=zo%>&s=<%=co%>&em=<%=""%>&e3=<%=np%>&so=<%=so%>">K&otildeik</a></td>
</tr>
</Form>
</table>
</body>
</html>