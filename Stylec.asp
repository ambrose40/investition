<Html>
 <Head>
<!--Скрипт для подтверждения закрытия окна изменения стиля отображения пользователем-->
  <Script Type="text/javaScript">
   function confirmClose() 
    {
    if (confirm("Kas tahate panema see aken kinni?")) 
     {
     parent.close();
     }
    }
  </Script>
<!--Записываем физический путь виртуального пути Inv-->
  <%b= Server.MapPath("\inv")%>
<!--Запрос если значение кука StyleInv пустое, то загружаем стандартную гамму записанную в конфигуационном файле на сервере-->
  <%If Request.Cookies("StyleInv")="" Then%>
   <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <Link Rel="stylesheet" href="<%=s%>" Type="text/css">
  <%Else%>
<!--Запрос если значение кука StyleInv не пустое, то загружаем гамму указанную в этом значении-->
   <%s=Request.Cookies("StyleInv")%>
   <Link Rel="stylesheet" href="<%=s%>" Type="text/css">
  <%End If%>
<!--Русскоязычная кодировка-->
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Title>Invest-IT!on: N&Auml;GEMUSE VALIMINE</Title>
 </Head>
 <Body Class="Main">
<!--При шелчке мышью на заголовок вызывается скрипт подтверждения закрытия и можно закрыть окошко настройки стиля-->
  <p align="center"><a href="Main.asp"  target="_top" Class="HeadLink" onClick="confirmClose()">N&Auml;GEMUSE VALIMINE</a></p><br><br><br>
<!--Если была нажата кнопка MUUTA то записывается новое значение кука StyleInv со сроком истчения и путём-->
  <%If Request.Form("btn")="MUUTA" Then%>
   <%s=Request.Form("styl")%>
   <%response.Cookies("StyleInv")=s%> 
   <%Response.Cookies("StyleInv").path="/"%>
   <%Response.Cookies("StyleInv").expires="01/01/2010"%>
  <%End If%>
<!--Рендируем форму для выбора заголовленных стилей с кнопочко подтверждающей изменения-->
  <Form Action="stylec.asp" Method="POST">
   <Select  Name="styl"  Class="Main">
    <Option Value="STYLE.CSS">Kollane N&auml;gemus</Option>
    <Option Value="STYLE2.CSS">Sinine N&auml;gemus</Option>
    <Option Value="STYLE3.CSS">Roheline N&auml;gemus</Option>
    <Option Value="STYLE4.CSS">Halline N&auml;gemus</Option>
    <Input Type="Submit" Name="btn" Value="MUUTA" Class="Main">
   </select>
  </Form>
 </Body>
</Html>