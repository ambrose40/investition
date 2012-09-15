<html>
 <head>
  <title>Invest-IT!on: GRAAFIKUD </title>
  <link rel="shortcut icon" href="favicon.ico">
<!--Серверный скрипт выбора стиля оформления страницы на основе выбора пользователя, путём чтения объекта Cookie-->
  <%b= Server.MapPath("\")%>
  <%If request.Cookies("StyleInv")="" Then%>
   <%set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%Else%>
   <%s=request.Cookies("StyleInv")%>
   <link rel="stylesheet" href="<%=s%>" type="text/css">
  <%End if%>
 </head>
 <body class="chart">
  <p align="center"><Font color="#ff0000" face="VERDANA" size="5">Susteemi INVEST-IT!ON graafikud</Font></p>
  <hr class="chart">
  Excelis kasutatav investeeringute <a href="chart.xls" target="_blank" class="chart">graafik</a> on loodud ettevotte-siseste (office) dokumentide ja aruannete koostamiseks Microsoft Graph abil.
  <hr class="chart">
  Ulalnimetatud viite abil koostatud graafiku <a href="page.htm" target="_blank" class="chart">kuvamine</a> tekstivormingu HTML-lehekuljena.
  <hr class="chart">
  Uleminek <a href="chart.asp" target="_blank" class="chart">lehekuljele</a>, mis loob universaalsed investeeringute graafikud tekstivormingus.
  <hr class="chart">
  OLAP tarkvaral pohinev investeeringute <a href="olap.xls" target="_blank" class="chart">graafik</a> ja aruanne on loodud andmete alusel, mis on struktureeritud analuusi tegemiseks andmetootlusvahendi SQL Analysis Services abil.
  <hr class="chart">
 </body>
</html>
