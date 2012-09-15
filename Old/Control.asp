
<html>
<Head><meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>
Invest-IT!on: SUMMADE JA K&Otilde;RVALEKALDUMISTE KONTROLL
</title></Head>
<body bgcolor="CCCCCC">

<img border="0" src="icons/SINEWAVE.ICO" Style=float:Left><p align="center"><p align="center"><a href="Main.asp"><font face="Verdana" Size="5" color="000099"><b><u>SUMMADE JA K&Otilde;RVALEKALDUMISTE KONTROLL</font></u></b></a></p><p>
<hr>
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<%b= Server.MapPath("\inv")%>
<%set mdbo =  Server.CreateObject("ADODB.Connection")%>
<%set servcfg=Server.CreateObject("Scripting.FileSystemObject")
  set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close%>
<%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Database=invest;Trusted_Connection=yes;"%>
<%mdbo.Open ConnectionString%>

<%set mdbo2 = Server.CreateObject("ADODB.Command")%>
<%set mdbor = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2.ActiveConnection = mdbo%>
<%set mdbo2P = Server.CreateObject("ADODB.Command")%>
<%set mdbor2P = Server.CreateObject("ADODB.Recordset")%>
<%mdbo2P.ActiveConnection = mdbo%>
<%set mdbode = Server.CreateObject("ADODB.Command")%>
<%set mdborde = Server.CreateObject("ADODB.Recordset")%>
<%mdbode.ActiveConnection = mdbo%>
<%set mdbode3 = Server.CreateObject("ADODB.Command")%>
<%set mdborde3 = Server.CreateObject("ADODB.Recordset")%>
<%mdbode3.ActiveConnection = mdbo%>

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

<%srt=Request.QueryString("sr")%>
<%If Request.Form("btn") = "OK" Then%>
<%zo="Yearr='" & Request.Form("ye") & "'"%>
<%Else%>
<%zo=Request.QueryString("y")%>

<%If zo="" Then%>

<%ya=Year(Date())%>
<%mo=Month(Date())%>
<%da=Day(Date())%>
<%zz=mo-04%>

<%If zz>=0 Then%>
<%ya=Year(Date())%>
<%Else%>
<%ya=ya-1%>
<%End If%>

<%zo="yearr='" & ya & "'"%>
<%End If%>
<%End If%>
<%np=Request.QueryString("e3")%>


<table bordercolor="0F0F0F" border="1"  style="border-collapse: collapse">
<tr  bgcolor="666666">
<td>
<Font Color="FFFFFF" Face="Verdana" size="1">
No
</Font>
</td>
<td>
<a href="control.asp?so=<%=Request.QueryString("so")%>&y=<%=zo%>"><Font Color="FFFFFF" Face="Verdana" size="1">*</Font></a>
</td>
<%so=Request.QueryString("so")%>
<%If srt="" then%>
  <td>
 <a href="control.asp?sr=<%="ProjCode,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Projekti kood ja nimetus</Font></a>
  <.td>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Projekti kood ja nimetus <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Projekti kood ja nimetus <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
   <td>
   <a href="control.asp?sr=<%=bz & "ProjCode,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Projekti kood ja nimetus</Font></a>
   </td>
  <%End If%>
 <%End If%>
<%End If%>


 <td><Font Color="FFFFFF" Face="Verdana"  size="1">Aasta</Font></td>
<%set mdboe = Server.CreateObject("ADODB.Command")%>
<%set mdbore = Server.CreateObject("ADODB.Recordset")%>
<%mdboe.ActiveConnection = mdbo%>
<%mdboe.CommandText="SELECT Enterprise,EDescr from Enterprise"%>
<%mdbore.Open mdboe%>


<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="Enterprise,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Ettev&otilde;te</Font></a>
<select size="1" name="entt"style="font-family: Verdana; color: #FFFFFF; font-size:small; background-color: #777777; border-width:0">
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
<input type="submit" value="   " name="filt" size="5" style="font-family: Verdana; color: #FFFFFF; font-size:medium; background-color: #777777; border-width:0; background-image:url('icons/filter.png'); background-repeat: no-repeat; background-position: center;" >
 </td>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Ettev&otilde;te <img border="0" src="icons/down.png"></Font></a>
<select size="1" name="entt" style="font-family: Verdana; color: #FFFFFF; font-size:small; background-color: #777777; border-width:0" >
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
<input type="submit" value="   " name="filt" size="5" style="font-family: Verdana; color: #FFFFFF; font-size:medium; background-color: #777777; border-width:0; background-image:url('icons/filter.png'); background-repeat: no-repeat; background-position: center;" >
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
 <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Ettev&otilde;te <img border="0" src="icons/up.png"></Font></a>
<select size="1" name="entt" style="font-family: Verdana; color: #FFFFFF; font-size:small; background-color: #777777; border-width:0" >
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
<input type="submit" value="   " name="filt" size="5" style="font-family: Verdana; color: #FFFFFF; font-size:medium; background-color: #777777; border-width:0; background-image:url('icons/filter.png'); background-repeat: no-repeat; background-position: center;">  

   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "Enterprise,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Ettev&otilde;te</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>


<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="summaPlan,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Plaaniline summa</Font></a>
 </td>
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
   <%If a2="summaPlan" Then%>
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
    <%If a2="summaPlan" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Plaaniline summa <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Plaaniline summa <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "summaPlan,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Plaaniline summa </Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>


<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="summaFact,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Tegelik summa</Font></a>
 </td>
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
   <%If a2="summaFact" Then%>
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
    <%If a2="summaFact" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Tegelik summa <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Tegelik summa <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "summaFact,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Tegelik summa</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>

<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="summaContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Lepingup&otilde;hine summa </Font></a>
 </td>
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
   <%If a2="summaContract" Then%>
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
    <%If a2="summaContract" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Lepingup&otilde;hine summa  <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td> 
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana"  size="1">Lepingup&otilde;hine summa  <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "summaContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana"  size="1" >Lepingup&otilde;hine summa </Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>



<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="TotsummaPlan,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud Plaaniline summa</Font></a>
 </td>
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
   <%If a2="TotsummaPlan" Then%>
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
    <%If a2="TotsummaPlan" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud Plaaniline summa <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud Plaaniline summa <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "TotsummaPlan,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud Plaaniline summa</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>


<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="TotsummaFact,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud tegelik summa</Font></a>
 </td>
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
   <%If a2="TotsummaFact" Then%>
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
    <%If a2="TotsummaFact" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud tegelik summa <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud tegelik summa <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "TotsummaFact,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud tegelik summa</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>

<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="TotsummaContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud lepingup&otilde;hine summa </Font></a>
 </td>
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
   <%If a2="TotsummaContract" Then%>
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
    <%If a2="TotsummaContract" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud lepingup&otilde;hine summa  <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud lepingup&otilde;hine summa  <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "TotsummaContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1">Kogunenud lepingup&otilde;hine summa </Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>




<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="PlanFact,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - teostus</Font></a>
 </td>
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
   <%If a2="PlanFact" Then%>
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
    <%If a2="PlanFact" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - teostus <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - teostus <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "PlanFact,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - teostus</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>

<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="PlanContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - leping</Font></a>
 </td>
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
   <%If a2="PlanContract" Then%>
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
    <%If a2="PlanContract" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - leping <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - leping <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "PlanContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Kava - leping</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>

<%If srt="" then%>
 <td>
 <a href="control.asp?sr=<%="FactContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Teostus - leping</Font></a>
 </td>
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
   <%If a2="FactContract" Then%>
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
    <%If a2="FactContract" Then%>
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
  <td>
  <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Teostus - leping <img border="0" src="icons/down.png"></Font></a>
  </td> 
 <%Else%>
  <%If f=4 Then%>
   <td>
   <a href="control.asp?sr=<%=bz%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Teostus - leping <img border="0" src="icons/up.png"></Font></a>
   </td>  
  <%Else%> 
  <td>
  <a href="control.asp?sr=<%=bz & "FactContract,"%>&no=<%=n%>&y=<%=zo%>&s=<%=co%>&em=<%=pb%>&e3=<%=np%>&so=<%=so%>"><Font Color="FFFFFF" Face="Verdana" size="1"><img src="icons/delta.png" border="0">Teostus - leping</Font></a>
  </td>
  <%End If%>
 <%End If%>
<%End If%>

<%If srt<>"" Then%>
<%l2=len(srt)%>
<%srt=Mid(srt,1,l2-1)%>
<%End If%>

<%j=1%>


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
<%mdbo2.CommandText="SELECT DISTINCT Pid,ProjCode,ProjName,Yearr from inpl WHERE " & zo & " AND ProjCode LIKE '__.00.00' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='7')"%>
<%mdbor.Open mdbo2%>
<%mdbode.CommandText="SELECT SUM(TotsummaContract) AS TSC, SUM(TotsummaFact) AS TSF, SUM(TotsummaPlan) AS TSP,SUM(FactContract) AS FC,SUM(PlanContract) AS PC, SUM(PlanFact) AS PF, SUM(summaFact) AS SF, SUM(summaPlan) AS SP, SUM(summaContract) AS SC, be FROM dbo.Delta WHERE " & yr & " GROUP BY be"%>
<%mdborde.Open mdbode%>
<%If srt="" Then%>
<%dd=""%>
<%Else%>
<%dd=","%>
<%End If%>

<Form Method ="POST" Action="Control.asp?sr=<%=srt & dd%>&y=<%=zo%>&e3=<%=np%>#vira">
<%Do Until mdbor.EOF%>

<tr bgcolor="FFFFAA">
<td>
<Font Color="0000FF" Face="Verdana">
<%=j%>
</Font>
</td>
<td>
<%a="bpm" & mdbor("ProjCode") & mdbor("Yearr")%>
<%b="vpm" & mdbor("ProjCode") & mdbor("Yearr")%>
<%c="vpc" & mdbor("ProjCode") & mdbor("Yearr")%>


<%If request.Form(a)="" and request.Form(b)="" and request.Form(c)="" then%>
<input type="submit" value="+" name="<%=a%>">
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="-" name="<%=b%>">
<%Else%>
<%If request.Form(a)="" AND request.Form(b)<>"" then%>
<%If Request.Form(b)="+" Then%>
<%z="-"%>
<%End If%>
<%If Request.Form(b)="-" Then%>
<%z="+"%>
<%End If%>
<%'=z%>
<Input type="submit" value="<%=z%>" name="<%=a%>">
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="<%=Request.Form(b)%>" name="<%=b%>">
<%Else%>
<%If request.Form(a)="+" then%>
<a name="vira"></a>
<Input type="submit" value="-" name="<%=a%>">
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="+" name="<%=b%>">
<%Else%>
<%If request.Form(a)="-" then%>
<a name="vira"></a>
<Input type="submit" value="+" name="<%=a%>">
<input type="hidden" value="<%=mdbor("ProjCode")%>" name="<%=c%>">
<input type="hidden" value="-" name="<%=b%>">
<%Else%>
<%End If%>
<%End If%>
<%End If%>
<%End If%>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=Mid(mdbor("ProjCode"),1,3)%> / <%=mdbor("ProjName")%>
</Font>
</td>
<td colspan="2">
<Font Color="0000FF" Face="Verdana" size="1"><%=mdbor("Yearr")%></Font>
</Font>
</td>

<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("SP")%>
<%sup=sup+CDBl(mdborde("SP"))%>
<%End If%>
<%g=g+1%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("SF")%>
<%suf=suf+Cdbl(mdborde("SF"))%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("SC")%>
<%suc=suc+Cdbl(mdborde("SC"))%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("TSP")%>
<%tsup=tsup+CDbl(mdborde("TSP"))%>
<%End If%>
<%g=g+1%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("TSF")%>
<%tsuf=tsuf+Cdbl(mdborde("TSF"))%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("TSC")%>
<%tsuc=tsuc+CDbl(mdborde("TSC"))%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("PF")%>
<%pf=pf+Cdbl(mdborde("PF"))%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("PC")%>
<%pc=pc+Cdbl(mdborde("PC"))%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%=mdborde("FC")%>
<%fc=fc+CDbl(mdborde("FC"))%>
<%End If%>
</Font>
</td>
</tr>
<%c="vpc" & mdbor("ProjCode") & mdbor("Yearr")%>

<%If Request.Form(a)="+" or (Request.Form(a)="" and Request.Form(b)="+") then%>
<%p8=mid(Request.Form(c),1,2)%>
<%p9=mid(Request.Form(c),1,8)%>

<%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
<%mdbo3.CommandText="SELECT DISTINCT ProjCode,ProjName,Yearr from inpl WHERE "& zo & " AND Identifier='C' AND ProjCode LIKE '" & p8 & ".__.00' AND ProjCode <> '" & p9 & "' AND ((DateBegin<='" & d & "' AND DateEnd>='" & d & "') OR StatusID IS NULL OR StatusID='6' OR StatusID='7')"%>
<%mdbor3.Open mdbo3%>
<%mdbode3.CommandText="SELECT SUM(TotsummaContract) AS TSC, SUM(TotsummaFact) AS TSF, SUM(TotsummaPlan) AS TSP,SUM(FactContract) AS FC,SUM(PlanContract) AS PC, SUM(PlanFact) AS PF, SUM(summaFact) AS SF, SUM(summaPlan) AS SP, SUM(summaContract) AS SC, be, mi FROM dbo.Delta WHERE " & yr &" AND (mi <> '00') AND be='" & p8 & "' GROUP BY be, mi"%>
<%mdborde3.Open mdbode3%>

<%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>

<%Do Until mdbor3.EOF%>
<%j=j+1%>
<tr bgcolor="FFFFAA">
<td>
<Font Color="0000FF" Face="Verdana">
<%=j%>
</Font>
</td>
<td>

<%a2=a & "bpm2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%b2=b & "vpm2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%c2=c & "vpc2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%If request.Form(a2)="" and request.Form(b2)="" and request.Form(c2)="" then%>
<input type="submit" value="+" name="<%=a2%>">
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
<Input type="submit" value="<%=z%>" name="<%=a2%>">
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="<%=Request.Form(b2)%>" name="<%=b2%>">
<%Else%>
<%If request.Form(a2)="+" then%>
<a name="vira"></a>
<Input type="submit" value="-" name="<%=a2%>">
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="+" name="<%=b2%>">
<%Else%>
<%If request.Form(a2)="-" then%>
<a name="vira"></a>
<Input type="submit" value="+" name="<%=a2%>">
<input type="hidden" value="<%=mdbor3("ProjCode")%>" name="<%=c2%>">
<input type="hidden" value="-" name="<%=b2%>">
<%Else%>
<%End If%>
<%End If%>
<%End If%>
<%End If%>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<Font face="Arial">&nbsp<img border="0" src="icons/lev.png">&nbsp</font><%=Mid(mdbor3("ProjCode"),1,6)%>/<%=mdbor3("ProjName")%>
</Font>
</td>
<td colspan="2"> 
<Font Color="0000FF" Face="Verdana" size="1"><%=mdbor3("Yearr")%></Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("SP")%>
<%End If%>
<%g2=g2+1%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("SF")%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("SC")%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("TSP")%>
<%End If%>
<%g=g+1%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("TSF")%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("TSC")%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("PF")%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("PC")%>
<%End If%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%If mdborde3("be") & "." & mdborde3("mi") = Mid(mdbor3("ProjCode"),1,5) Then%>
<%=mdborde3("FC")%>
<%End If%>
</Font>
</td>
</tr>

<%c2=c & "vpc2" & mdbor3("ProjCode") & mdbor3("Yearr")%>
<%'=Request.Form(c2)%>
<%If Request.Form(a2)="+" or (Request.Form(a2)="" and Request.Form(b2)="+") then%>
<%p8=mid(Request.Form(c2),1,5)%>
<%p9=mid(Request.Form(c2),1,8)%>

<%d=Month(Date()) & "." & Day(Date()) & "." & Year(Date())%>
<%If srt="" then%>
<%mdbo4.CommandText="SELECT * from Delta WHERE " & yr & " AND ProjCode LIKE '" & p8 & ".__' AND ProjCode <> '" & p9 & "'"%>
<%mdbor4.Open mdbo4%>
<%Else%>
<%mdbo4.CommandText="SELECT * from Delta WHERE " & yr & " AND ProjCode LIKE '" & p8 & ".__' AND ProjCode <>  '" & p9 & "' ORDER BY " & srt%>
<%mdbor4.Open mdbo4%>
<%End If%>

<%Do Until mdbor4.EOF%>
<%j=j+1%>
<tr bgcolor="FFFFAA">
<td>
<Font Color="0000FF" Face="Verdana">
<%=j%>
</Font>
</td>
<td>

<%a3=c2 & "bpm" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("Enterprise")%>
<%b3=b2 & "vpm" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("Enterprise")%>
<%c3=c2 & "vpc" & mdbor4("ProjCode") & mdbor4("Yearr") & mdbor4("Enterprise")%>
<%If request.Form(a3)="" and request.Form(b3)="" and request.Form(c3)="" then%>
<input type="submit" value="+" name="<%=a3%>">
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="-" name="<%=b3%>">
<%Else%>
<%'=request.Form(a3)%>
<%If request.Form(a3)="" AND request.Form(b3)<>"" then%>
<%'=request.Form(b3)%>
<%If Request.Form(b3)="+" Then%>
<%z="-"%>
<%End If%>
<%If Request.Form(b3)="-" Then%>
<%z="+"%>
<%End If%>
<%'=z%>
<Input type="submit" value="<%=z%>" name="<%=a3%>">
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="<%=Request.Form(b3)%>" name="<%=b3%>">
<%Else%>
<%If request.Form(a3)="+" then%>
<a name="vira"></a>
<Input type="submit" value="-" name="<%=a3%>">
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="+" name="<%=b3%>">
<%Else%>
<%If request.Form(a3)="-" then%>
<a name="vira"></a>
<Input type="submit" value="+" name="<%=a3%>">
<input type="hidden" value="<%=mdbor4("ProjCode")%>" name="<%=c3%>">
<input type="hidden" value="-" name="<%=b3%>">
<%Else%>
<%End If%>
<%End If%>
<%End If%>
<%End If%>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<Font face="Arial">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<img border="0" src="icons/lev.png">&nbsp</font><%=mdbor4("ProjCode")%>/<%=mdbor4("ProjName")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1"><%=mdbor4("Yearr")%></Font>
</td>
<td>
<a href="Control.asp?sr=<%=srt & dd%>&y=<%=zo%>&e3=<%="Enterprise=" & mdbor4("Enterprise")%>"><Font Color="0000FF" Face="Verdana" size="1"><%=mdbor4("EDescr")%></Font>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("summaPlan")%>
<%g1=g1+1%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("summaFact")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("summaContract")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("TotsummaPlan")%>
<%g=g+1%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("TotsummaFact")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("TotsummaContract")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("PlanFact")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("PlanContract")%>
</Font>
</td>
<td>
<Font Color="0000FF" Face="Verdana" size="1">
<%=mdbor4("FactContract")%>
</Font>
</td>
</tr>


<%mdbor4.MoveNext%>
<%loop%>
<%mdbor4.Close%>
<%End If%>
<%If mdborde3("be") & "." & mdborde3("mi")=Mid(mdbor3("ProjCode"),1,5) Then%>
<%mdborde3.MoveNext%>
<%End if%>

<%mdbor3.MoveNext%>
<%loop%>
<%mdborde3.Close%>
<%mdbor3.Close%>
<%End If%>

<%j=j+1%>

<%If mdborde("be")=Mid(mdbor("ProjCode"),1,2) Then%>
<%mdborde.MoveNext%>
<%End if%>
<%mdbor.MoveNext%>
<%loop%>
<tr bgcolor="CCEEFF">
<td>
<Font Color="000099" Face="Verdana">
<a href="chart.asp?pg=<%="Control"%>"><Img border="0" src="icons/chart.png"></a>
</Font>
</td>
<td>
<a href="Invest.asp?sr=<%=srt & dd%>&y=<%=zo%>&e3=<%=np%>"><Font Color="000000" Face="Verdana">X</Font></a>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
Kokku
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
--
</Font>
</td>
<td>

<a href="Control.asp?sr=<%=srt & dd%>&y=<%=zo%>&e3="><Font Color="000099" Face="Verdana" size="1">K&otilde;ik</Font></a>

</td>

<td>
<Font Color="000099" Face="Verdana" size="1">
<%=sup%>
<%sup=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=suf%>
<%suf=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=suc%>
<%suc=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=tsup%>
<%tsup=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=tsuf%>
<%tsuf=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=tsuc%>
<%tsuc=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=pf%>
<%pf=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=pc%>
<%pc=0%>
</Font>
</td>
<td>
<Font Color="000099" Face="Verdana" size="1">
<%=fc%>
<%fc=0%>
</Font>
</td>

</tr>
</Form>
</table>
<body>
<html>