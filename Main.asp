<Html>
 <%Set servcfg=Server.CreateObject("Scripting.FileSystemObject")%>
 <Head>
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Link Rel="SHORTCUT ICON" Href="https://intranet/inv/invfavicon.ico">
<!--Серверный скрипт выбора стиля оформления страницы на основе выбора пользователя, путём чтения объекта Cookie-->
  <%b= Server.MapPath("\")%>
  <%If Request.Cookies("StyleInv")="" Then%>
   <%Set servFileStream=servcfg.OpenTextFile(b & "\style.cfg")%>
   <%s=servFileStream.ReadLine%>
   <%servFileStream.Close%>
   <Link Rel="stylesheet" Href="<%=s%>" type="text/css">
  <%Else%>
   <%s=Request.Cookies("StyleInv")%>
   <Link Rel="stylesheet" Href="<%=s%>" type="text/css">
  <%End If%>
<!--Установке досовской кодировки-->
  <Meta http-equiv="Content-Type" Content="text/Html; Charset=windows-1251">
  <Title>InFormatsiooniSusteem Investeerimise Kava Teostamise Kontrollimiseks. Invest-IT!on</Title>
<!--Здесь идут некоторые клиентские скрипты помогающие создать более приемлимый и дружественный интерфейс. -->
  <Script Language="JavaScript">

   function onKeyPress () {
    var keycode;
    if (window.event) keycode = window.event.keyCode;
    else if (e) keycode = e.which;
    else return true;
    if (keycode == 13) {
     newWindow ('Help.mht','','','','');
     return false
    }
    return true 
   }
   document.onkeypress = onKeyPress;
<!--Скрипт который открывает проектную карточку в окне без полос прокрутки и других излишностей-->
   function change(a,b,c)
   {
   newWindow ('Projcard.asp?ye=' + a + '&entt=' + b + '&pco=' + c,'','800','600','scrollbars');
   }
<!--Скрипт инициализации некой страницы в новом окне-->
   var win = null;
function newWindow(mypage,myname,w,h,features) {
  var winl = 0;
  var wint = 02;
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
  </Script>
  <Script Language="VBScript"> 
<!--Скрипты подтверждения формы для окошек выбора года за который будет формироваться отчёт-->
<!--
   Sub submi
    Dim TheForm
    Set TheForm = Document.Forms("Forma")
    TheForm.Submit
   End Sub
   Sub submic
    Dim TheForm
    Set TheForm = Document.Forms("Formac")
    TheForm.Submit
   End Sub
   Sub submip
    Dim TheForm
    Set TheForm = Document.Forms("Formap")
    TheForm.Submit
   End Sub
   Sub submi0
    Dim TheForm
    Set TheForm = Document.Forms("Forma0")
    TheForm.Submit
   End Sub
   Sub submiz
    Dim TheForm
    Set TheForm = Document.Forms("Formaz")
    TheForm.Submit
   End Sub
   Sub submi1
    Dim TheForm
    Set TheForm = Document.Forms("Forma1")
    TheForm.Submit
   End Sub
   Sub submi3
    Dim TheForm
    Set TheForm = Document.Forms("Forma3")
    TheForm.Submit
   End Sub
   Sub submi4
    Dim TheForm
    Set TheForm = Document.Forms("Forma4")
    TheForm.Submit
   End Sub
   Sub submi6
    Dim TheForm
    Set TheForm = Document.Forms("Forma6")
    TheForm.Submit
   End Sub
   Sub submi9
    Dim TheForm
    Set TheForm = Document.Forms("Forma9")
    TheForm.Submit
   End Sub
   Sub submi10
    Dim TheForm
    Set TheForm = Document.Forms("Forma10")
    TheForm.Submit
   End Sub
   Sub submi12
    Dim TheForm
    Set TheForm = Document.Forms("Forma12")
    TheForm.Submit
   End Sub
 -->
  </Script>
 </Head>
 <Body Class="Main">
<!--Создаём объект для листания ссылок-->
  <%Set Nol=Server.CreateObject("MSWC.NextLink")%>
<!--Подключение к серверу баз данных. Создаём объект Подключение-->
  <%Set mdbo =  Server.CreateObject("ADODB.Connection")%>
<!--В переменную b закидываем физический путь к каталогу интранета inv-->
<!--Имя сервера вынимается из конфигурационного файла server.cfg и подставляется к строке подключения-->
  <%Set servFileStream=servcfg.OpenTextFile(b & "\server.cfg")%>
  <%s=servFileStream.ReadLine%>
  <%i=servFileStream.ReadLine%>
  <%p=servFileStream.ReadLine%>
  <%servFileStream.Close%>
<!--Инициализация строки подключения-->
  <%mdbo.ConnectionString="Driver={SQL Server};Server=" & s & ";Trusted_Connection=yes;Database=invest;"%>
  <%mdbo.Open ConnectionString%>
<!--Прорисовка заголовка-->
  <Img Border="0" Src="icons/delta.ico" Style=float:Left><p Align="center"></Img><A Class="HeadLink" Href="#">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<b>TERE TULEMAST INVEST-IT!ON &copy; INFOS&Uuml;STEEMI LEHEK&Uuml;LJELE!</a></b></p>
  <hr Class="main">
<!--Расчёт текущего финансового года исходя из текущей календарной даты-->
  <%ya=Year(Date())%>
  <%mo=Month(Date())%>
  <%da=Day(Date())%>
  <%zz=mo-04%>
  <%If zz>=0 Then%>
  <%ya=Year(Date())%>
  <%Else%>
  <%ya=ya-1%>
  <%End If%>
<!--Создание объектов Команды и Массива записей для работы с базой данных-->
  <%Set mdbop = Server.CreateObject("ADODB.Command")%>
  <%Set mdborp = Server.CreateObject("ADODB.RecordSet")%>
  <%Set mdboe = Server.CreateObject("ADODB.Command")%>
  <%Set mdbore = Server.CreateObject("ADODB.RecordSet")%>
  <%Set mdboy = Server.CreateObject("ADODB.Command")%>
  <%Set mdbory = Server.CreateObject("ADODB.RecordSet")%>
<!--Связь объекта Команда и объекта Подключение-->
  <%mdbop.ActiveConnection = mdbo%>
  <%mdboe.ActiveConnection = mdbo%>
  <%mdboy.ActiveConnection = mdbo%>
<!--Задание текста запроса и подключение к объекта Команда к объекту Массива записей-->
<!--Запрашиваем базу данных для вывода списка всех проектов, которые активны начиная с текущего финансового года. 
Используется представление kaart, где справочник проектов объединяется со главной инвестиционной таблицей-->
  <%mdbop.CommandText="Select DISTINCT Pid,ProjCode,ProjName from kaart WHERE yearr>='" & ya & "' ORDER BY Pid"%>
  <%mdborp.Open mdbop%>
<!--Запрашиваем список всех предприятий-->
  <%mdboe.CommandText="Select Enterprise,EDescr from Enterprise"%>
  <%mdbore.Open mdboe%>
<!--Запрашиваем список годов которые по которым существуют записи в инвестиционном плане-->
  <%mdboy.CommandText="Select DISTINCT yearr from Main ORDER BY Yearr DESC"%>
  <%mdbory.Open mdboy%>

<!--Заполнение таблицы с меню системы-->
  <Table Class="main" Align="Center"> 
   <tr Class="main">
    <td Valign=top>
     <li Class="main">Aruannete koostamine</li>
     <ul Style="list-style-type:circle">
<!--Создание Web-формы и присвоение ей уникального идентификатора-->
      <Form Method="GET" Action="Report3.asp" Class="main" id="Forma3">
<!--Ссылка и её описание вынимаються специальным серверным объектов из файла списка ссылок Links.cfg-->
       <li Class="main"><Img Src="icons/3kuu.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main" Href="<%=Nol.GetNthURL("Links.cfg", 13)%>"><%=Nol.GetNthDeScription("Links.cfg",13)%></a></li>&nbsp &nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi3()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";'  style="margin:0; padding:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report6.asp" Class="main" id="Forma6">
       <li Class="main"><Img Src="icons/6kuu.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 14)%>"><%=Nol.GetNthDeScription("Links.cfg",14)%></a></li>&nbsp;&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi6()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0; padding:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report9.asp" Class="main" id="Forma9">
       <li Class="main"><Img Src="icons/9kuu.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 10)%>"><%=Nol.GetNthDeScription("Links.cfg",10)%></a></li>&nbsp &nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi9()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0;">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report12.asp" Class="main" id="Forma12">
       <li Class="main"><Img Src="icons/12kuu.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 15)%>"><%=Nol.GetNthDeScription("Links.cfg",15)%></a></li>
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi12()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report4.asp" Class="main" id="Forma4">
       <li Class="main"><Img Src="icons/neli.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 11)%>"><%=Nol.GetNthDeScription("Links.cfg",11)%></a></li>&nbsp&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi4()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report1.asp" Class="main" id="Forma1">
       <li Class="main"><Img Src="icons/igak.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 12)%>"><%=Nol.GetNthDeScription("Links.cfg",12)%></a></li>&nbsp&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye"  onchange="submi1()"  onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report10.asp" Class="main" id="Forma10">
       <li clss="main"><Img Src="icons/teny.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 16)%>"><%=Nol.GetNthDeScription("Links.cfg",16)%></a></li>&nbsp&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi10()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report_r.asp" Class="main" id="Forma0">
       <li Class="main"><Img Src="icons/sise.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 20)%>"><%=Nol.GetNthDeScription("Links.cfg",20)%></a></li>&nbsp&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye"  onchange="submi0()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
      <Form Method="GET" Action="Report_0.asp" Class="main" id="Formaz">
       <li Class="main"><Img Src="icons/repo.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 23)%>"><%=Nol.GetNthDeScription("Links.cfg",23)%></a></li>&nbsp&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye"  onchange="submiz()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>
      </Form>
     </ul>
    </td>
    <td Valign=top>
     <li Class="main">Peamine osa</li>
     <ul Style="list-style-type:circle">
      <Form Method="POST" Action="Invest.asp?sr=ProjCode,&no=&y=&s=&em=&e3=&so=" Class="main" id="Forma">
       <li Class="main"><Img Src="icons/kava.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 2)%>"><%=Nol.GetNthDeScription("Links.cfg",2)%></a></li>&nbsp&nbsp
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <Select Class="main" size="1" name="ye" onchange="submi()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";'style="margin:0">
        <Option value="">Tegelik m.a.</Option>
        <%Do Until mdbory.EOF%>
         <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
         <%mdbory.MoveNext%>
        <%Loop%>
       </Select>  
      </Form>
<!--Временно отключенный устаревший элемент системы
      <Form Method="GET" Action="Control.asp" Class="main" id="Formac">
       <li Class="main"><a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 3)%>"><%=Nol.GetNthDeScription("Links.cfg",3)%></a>&nbsp l&otilde;ikes&nbsp
        <%If mdbory.BOF=False Then%>
         <%mdbory.MoveFirst%>
        <%End If%>
        <Select Class="main" size="1" name="ye" onchange="submic()" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";' style="margin:0">
         <%Do Until mdbory.EOF%>
          <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
          <%mdbory.MoveNext%>
         <%Loop%>
        </Select>
       </li>
      </Form>
-->
      <%ref="newWindow('" & Nol.GetNthURL("Links.cfg", 6) & "','','950','530','')"%>
      <li Class="main"><Img Src="icons/ansi.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="#null" onClick="<%=ref%>"><%=Nol.GetNthDeScription("Links.cfg",6)%></a></li>
      <%ref="newWindow('" & Nol.GetNthURL("Links.cfg", 7) & "','','420','400','')"%>
      <li Class="main"><Img Src="icons/chrt.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="#null" onClick="<%=ref%>"><%=Nol.GetNthDeScription("Links.cfg",7)%></a></li>

<!--Временно отключенный устаревший элемент системы
      <li Class="main"><a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 5)%>"><%=Nol.GetNthDeScription("Links.cfg",5)%></a></li>
-->
      <li Class="main"><Img Src="icons/home.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 1)%>" onClick="this.style.behavior='url(#default#homepage)'; this.SetHomePage('https://intranet/inv/');">Tee koduleheks!</a></li>
      <%ref="newWindow('" & Nol.GetNthURL("Links.cfg", 8) & "','','720','700','scrollbars')"%>
      <li Class="main"><Img Src="icons/help.gif" Border="0" Valign="middle"></Img>&nbsp&nbsp<a Class="main"  Href="#null" onClick="<%=ref%>"><%=Nol.GetNthDeScription("Links.cfg",8)%></a></li>

<!--Временно отключенный устаревший элемент системы
      <li Class="main"><a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 9)%>"><%=Nol.GetNthDeScription("Links.cfg",9)%></a></li>
      <li Class="main"><a Class="main"  Href="<%=Nol.GetNthURL("Links.cfg", 17)%>"><%=Nol.GetNthDeScription("Links.cfg",17)%></a></li>
      <li Class="main"><a Class="main"  Href="http://sql-2/projectserver/Views/ProjectReport.asp?_projectID=419&_viewID=103&noBanter=0"><%=Nol.GetNthDeScription("Links.cfg",18)%></a></li>
-->

      <li Class="main"><Img Src="icons/coin.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main" Href="http://sql-2/sites/projectserver_126/default.aspx"><%=Nol.GetNthDeScription("Links.cfg",19)%></a></li>
      <li Class="main"><Img Src="icons/kiir.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main" Href="http://intranet/invrep/">Kiired Aruanned</a></li>
      <li Class="main"><Img Src="icons/comp.gif" Border="0" Valign="center"></Img>&nbsp&nbsp<a Class="main" Href="admin.asp">Administreerimine</a></li>
     </ul>
    </td>
   </tr>
  </Table>

  <Table Align="Center">
   <Form Method="GET" Action="#" Class="main" id="Formap">
    <tr Class="main">
     <td Colspan="2">
      <%yaa=Year(Date())%>
      <%mo=Month(Date())%>
      <%da=Day(Date())%>
      <%zz=mo-04%>
      <%If zz>=0 Then%>
      <%yaa=Year(Date())%>
      <%Else%>
      <%yaa=yaa-1%>
      <%End If%>
      <%ondc="change(yea.options[yea.selectedIndex].value,entt.options[entt.selectedIndex].value,pco.options[pco.selectedIndex].value)"%>
      <input Class="main" type="button" value="Ava" name="btn" onClick="<%=ondc%>" onmouseover='window.status="Avaneb Proektide Kaart Valitud  ja Projeti Numbri j&auml;rgi";'onmouseout='window.status="";'> Kaart 
      <Select Class="main" size="1" id="yea" onmouseover='window.status="Vali  siin.";'onmouseout='window.status="";'>
       <%If mdbory.BOF=False Then%>
        <%mdbory.MoveFirst%>
       <%End If%>
       <%Do Until mdbory.EOF%>
        <%If mdbory("Yearr")-yaa=0 Then%>
         <Option value="<%=mdbory("Yearr")%>" Selected="true"><%=mdbory("Yearr")%> m.a.</Option>
        <%End If%>
        <Option value="<%=mdbory("Yearr")%>"><%=mdbory("Yearr")%> m.a.</Option>
        <%mdbory.MoveNext%>
       <%Loop%>
      </Select>
       
      <Select Class="main" size="1" name="entt"  onmouseover='window.status="Vali ettev&otilde;te siin.";'onmouseout='window.status="";'>
       <%If mdbore.BOF=False Then%>
        <%mdbore.MoveFirst%>
       <%End If%>
       <%Do Until mdbore.EOF%>
        <Option value="<%=mdbore("Enterprise")%>"><%=mdbore("EDescr")%></Option>
        <%mdbore.MoveNext%>
       <%Loop%>
      </Select> ettev&otilde;tte
      <br>
     </td>
    </tr>
    <tr Class="main">
     <td Colspan="2">
      <Select Class="main" size="10" name="pco" style="Width:700;font-family:Lucida Console" ondblclick="<%=ondc%>" onmouseover='window.status="Vali Proekti kood ja number.";'onmouseout='window.status="";'>
       <%If mdborp.BOF=False Then%>
        <%mdborp.MoveFirst%>
       <%End If%>
       <%Do Until mdborp.EOF%>
        <%pp="" & mdborp("Pid")%>
        <%jp=4-len(pp)%>
        <%For i=1 to jp%>
         <%pp="&nbsp" & pp%>
        <%Next%>
        <%pc="" & mdborp("ProjCode")%>
        <%If len(pc)<9 Then%>
         <%pc=pc & "&nbsp&nbsp&nbsp"%>
        <%End If%>
        <Option value="<%=mdborp("Pid")%>"><%=pp%>&nbsp<%=chr(179)%>&nbsp<%=pc%>&nbsp<%=chr(179)%>&nbsp<%=mdborp("ProjName")%></Option>
        <%mdborp.MoveNext%>
       <%Loop%>
      </Select><br>projektide kohta
      <%mdbore.Close%><%mdborp.Close%>
     </Form>
    </td>
   </tr>
  </Table>
<!--Прорисовка баннеров-->
  <hr Class="main">
  <Table Align="Center">
   <tr Class="main">
    <td Class="main">
     <a Class="main"  Href="http://intranet/"><Img Src="Img/intranet.jpg" Border="2"></a>
    </td>
    <td Class="main">
     <a Class="main"  Href="http://www.eesise/dbout/index.php"><Img Src="Img/logo_energia.gif" Border="3"></a>
    </td>
    <td Class="main">
     <%If Request.QueryString("pic")="nice" Then%>
      <a Class="main" Href="https:\\intranet\inv\main.asp?pic=tech">Tehnoloogia</a>
      <%Set objRotate=Server.CreateObject("MSWC.AdRotator")%>
      <%objRotate.Border=3%>
      <%objRotate.Clickable=False%>
      <%RotateHtml=objRotate.GetAdvertisement("Rotaton.cfg")%>
      <a Class="main" Href="http://intranet/air_t/default.php" Target="_blank"><%=RotateHtml%></a>
     <%Else%>
      <%Set objRotate=Server.CreateObject("MSWC.AdRotator")%>
      <%objRotate.Border=3%>
      <%objRotate.Clickable=False%>
      <%RotateHtml=objRotate.GetAdvertisement("Rotator.cfg")%>
      <a Class="main" Href="http://www.powerplant.ee" Target="_blank"><%=RotateHtml%></a>
     <%End If%>
    </td>
    <td Class="main">
     <a Class="main" Href="http://intranet/bar"><Img Src="img/barcode.png" Border="2" Width=160 Height=40></a>
    </td>
    <td Class="main">
     <a Class="main" Href="http://intranet/invrep"><Img Src="img/quick.jpg" Border="2" Width=160 Height=40></a>
    </td>
   </tr>
  </Table>
  <hr Class="main">
  <p Align="Center">
   S&uuml;steemi projekteerija ja arendaja: <a Class="main"  Href="mailto:Boris.Lariushin@nj.energia.ee">Boris Lariushin</a> Tel. 66368<br>
   S&uuml;steemi arendusjuht: <a Class="main"  Href="mailto:Maksim.Starostin@nj.energia.ee">Maksim Starostin</a> Tel. 66518<br>
   IT koordineerija: <a Class="main"  Href="mailto:Andrei.Gorohhov@nj.energia.ee">Andrei Gorohhov</a> Tel. 66091
  </p>
<%mdbo.Close%>
 </Body>
</Html>