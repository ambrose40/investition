<SCRIPT LANGUAGE="JavaScript">
  function change(name,name2,name3,name4)
  {
    document.links[1].href=name;
    document.links[2].href=name2;
    document.links[3].href=name3;
    document.links[4].href=name4;
  }
<!--Скрипт инициализации некой страницы в новом окне-->
  var win = null;
  function newWindow(mypage,myname,w,h,features) 
  {
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
  function confirmClose() {
    if (confirm("Kas tahate panema see aken kinni?")) {
      parent.close();
    }

  }
<!--Скрипт вызова файла помощи при нажатии клавищи Enter-->
   function onKeyPress () 
   {
     var keycode;
     if (window.event) keycode = window.event.keyCode;
     else if (e) keycode = e.which;
     else return true;
     if (keycode == 13) 
     {
       newWindow ('Help.mht','','','','');
       return false
     }
     return true 
   }
   document.onkeypress = onKeyPress;
<!--Скрипт который открывает проектную карточку в окне без полос прокрутки и других излишностей-->
   function changem(a,b,c)
   {
     newWindow ('Projcard.asp?ye=' + a + '&entt=' + b + '&pco=' + c,'','800','600','scrollbars');
   }
  </SCRIPT>

  <SCRIPT LANGUAGE="VBScript" FOR="btnk"> 

  Sub btnk_OnClick
  Dim TheForm
  Set TheForm = Document.forms("ValidForm2")
     MyVar=MsgBox("Kas4 tahate kustuta see projekt???",VbYesNo)
	If myVar=6 then     
        TheForm.Submit
        End if
  End Sub
  Sub btnk2_OnClick
  Dim CEForm
  Set CEForm = Document.forms("ValidForm3")
     MyVar2=MsgBox("Kas3 tahate kustuta see projekt???",VbYesNo)
	If myVar2=6 then     
        CEForm.Submit
        End if
  End Sub
  </SCRIPT>
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