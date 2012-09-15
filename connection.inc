<!--Подключение к серверу баз данных. Создаём объект Подключение-->
  <%Set mdbo =  Server.CreateObject("ADODB.Connection")%>
<!--В переменную b закидываем физический путь к каталогу интранета inv-->
<!--Имя сервера вынимается из конфигурационного файла server.cfg и подставляется к строке подключения-->
<%
  Set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")
  s=servFileStream.ReadLine
  i=servFileStream.ReadLine
  p=servFileStream.ReadLine
  servFileStream.Close
%>
<!--Инициализация строки подключения-->
  <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Trusted_Connection=yes;Database=invest;"
  mdbo.Open ConnectionString%>