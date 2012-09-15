<!--Создается серверный объект для операций с фаловой системой, необходим для считывания конфигурационного файла (имя SQL сервера)-->
 <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
<!--Здается кодировка страницы и добавляется иконка для отображения в строке адреса-->
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Link Rel="SHORTCUT ICON" Href="https://intranet/inv/invfavicon.ico">
<!--Серверный скрипт выбора стиля оформления страницы на основе выбора пользователя, путём чтения объекта Cookie-->
  <%b= Server.MapPath("\inv")%>
<!--В случае если объект Cookie  пустой, то стиль оформления читается из файла конфигурации Style.cfg-->
  <%If Request.Cookies("StyleInv")="" Then
   Set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")
   s=servFileStream.ReadLine
   servFileStream.Close%>

   <Link Rel="stylesheet" Href="<%=s%>" type="text/css">
  <%Else
   s=Request.Cookies("StyleInv")%>
   <Link Rel="stylesheet" Href="<%=s%>" type="text/css">
  <%End If%>