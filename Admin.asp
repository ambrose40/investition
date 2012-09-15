<html>

<head>
<!--Серверный скрипт выбора стиля оформления страницы на основе выбора пользователя, путём чтения объекта Cookie-->
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

<!--Выбор русскоязычной кодировки-->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<SCRIPT LANGUAGE="JavaScript">
var win = null;
function newWindow(mypage,myname,w,h,features) {
  var winl = (screen.width-w)/8;
  var wint = (screen.height-h)/8;
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
</SCRIPT>

<title>
Invest-IT!on: ADMINISTREERIMINE
</title>
</head>
<body background="icons/back.gif" class="Main">

<!--Выбор русскоязычной кодировки-->
<img border="0" src="icons/sinewave.ico" Style=float:Left><p align="center"><a href="Main.asp" class="headlinka"><b>INVEST-IT!ON ADMINISTREERIMINE</b></a></p>
<br>
<!--Выбор русскоязычной кодировки-->
<%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<Hr>
<li class="Admin"><a class="Admin" href="<%=Nol.GetNthURL("Links.cfg", 5)%>"><%=Nol.GetNthDescription("Links.cfg",5)%></a></li>
<li class="Admin"><a class="Admin" href="<%=Nol.GetNthURL("Links.cfg", 9)%>"><%=Nol.GetNthDescription("Links.cfg",9)%></a></li>
<li class="Admin"><a class="Admin" href="<%=Nol.GetNthURL("Links.cfg", 17)%>"><%=Nol.GetNthDescription("Links.cfg",17)%></a></li>
<li class="Admin"><a class="Admin" href="http://sql-2/projectserver/Views/ProjectReport.asp?_projectID=419&_viewID=103&noBanter=0"><%=Nol.GetNthDescription("Links.cfg",18)%></a></li>
<li class="Admin"><a class="Admin" href="delete.asp">Andmete Kustutamine</a></li>
<li class="Admin"><a class="Admin" href="#" onClick="newWindow('stylec.asp','','300','200','')">N&auml;gemuse valimine</a></li>
</body>
</html>