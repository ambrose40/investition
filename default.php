<style type="text/css">
<!--
h3 {font-family: Times New Roman; color:black; font-size=18px;}
h2 {font-family: Times New Roman; color:black; font-size=16px;}
p {font-family: Tahoma; color:black; font-size=13px;}

TD {font-family: Arial; color:black; border-color:black; border-style:solid; border-width:1; font-size=12px;border:solid windowtext .5pt; text-align:center;}
TH {font-family: Arial; color:black; border-color:black; border-style:solid; border-width:1; font-size=12px;border:solid windowtext .5pt; text-align:center;}
TABLE {border-collapse:collapse; border:0; font-family: Arial; color:black; font-size=12px;}

.toolbar {
	padding-top: 0px;
	padding-bottom: 0px; padding-left: 0px;
	font-family: verdana;
	font-size: 11px; 
	color:black;
	background-color:white;
	border: 1px solid black;
}
-->
</style>

<script language=Javascript>
function b() {
	document.formata.text.focus()
	sel_1=document.selection.createRange();
	sel_1.text="<b>"+sel_1.text+"</b>"
}
function i() {
	document.formata.text.focus()
	sel_2=document.selection.createRange();
	sel_2.text="<i>"+sel_2.text+"</i>"
}
function u() {
	document.formata.text.focus()
	sel_2=document.selection.createRange();
	sel_2.text="<u>"+sel_2.text+"</u>"
}

</script>




<html><head>

<script language="vbscript">
	Sub submit
	Dim TheForm
	Set TheForm = Document.Forms("formata")
	if document.formata.ShortText.value="" or document.formata.text.value="" then
	else
	TheForm.Submit
	end if
	End Sub
</script>

<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=windows-1251'><title>Объявление для пользователей НЭС</title></head><body bgcolor=linen><center>
<h3 align=center>Объявление для пользователей НЭС</h3>

<?php

include("connect.php");
$kasutajanimi = (substr($REMOTE_USER,12));
$dostup = dostup($kasutajanimi);
$con = con();

if ($dostup==1) {
	if(!isset($action)) {$action=1;}
	//else {$action=2;}

	// $action = 1 - Создать новое сообщение
	// $action = 2 - Редактировать сообщение
	// $action = 3 - Деактивировать сообщение (удалить с главной страницы Интранета и поместить в архив)


	//-------------------------------------Вывод текущего сообщения----------------------------------------------

	$q = "select * from ITkuulutus where status=1";
	$qres = mssql_query($q, $con);
	$k = mssql_num_rows($qres);

	if ($k!=0) {
		print "<p align=center><b>Текущие сообщения на страничке Интранета</b></p>";
		print "<table width=750 border=1 cellpadding=5>";
		print "<tr bgcolor=#FFE0C9>";
		print "<th>Сообщение</th>";
		print "<th nowrap width=150>Изменить / Удалить</th>";
		print "</tr>";
		
		for ($i=1;$i<=$k;$i++){
			$a = mssql_fetch_row($qres);
			
			print "<tr bgcolor=white>";
			print "<td style='border-style:solid; border-width:1px;'>";

			print "<p align=left>";
			if ($a[4]!='') {print "<img src='".$a[4]."' border=0 align=left vspace=0 hspace=10>";}
			print "".nl2br($a[5])."";
			//print "<br><br><br><p align=justify>".nl2br($a[6])."";
			print "<p align=left><a href='teade.php?id=$a[0]'>См.подробнее</a>";

			print "</td>";
			print "<th><a href='default.php?action=2&id=$a[0]'>Изменить</a> / <a href='add.php?action=3&id=$a[0]'>Удалить</a></th>";
			print "</tr>";
		}
		print "</table>";
	}
	else {
		print "<p>Текущих сообщений нет!</p><hr size=1 width=60%>";
	}

	print "<p><a href='arhiiv.php'>Архив всех сообщений</a>";
	print "<br><br><br>";

	//-------------------------------------Ввод нового сообщения-------------------------------------------------
	?>
	<form action="default.php#preview" method="post" enctype="multipart/form-data" name="formata">
	<?php
	print "<input type=hidden name=actionT value='$action'>";

	$img_src1 = 'http://intranet/it/it_kuulutus/img1.gif';
	$img_src2 = 'http://intranet/it/it_kuulutus/img2.gif';
	$img_src3 = 'http://intranet/it/it_kuulutus/img4.gif';


	if (!$change & isset($id)) {
		$q = mssql_fetch_row(mssql_query("select * from ITkuulutus where id=$id", $con));
		$pic = $q[4];
		$Mname = $q[5];
		$Mtext = $q[6];
	}

	if ($change) {
		$Mname = stripslashes(stripslashes($Mname));
		//$Mname = str_replace('\'', '\'\'', $Mname);
		$Mtext = stripslashes(stripslashes($Mtext));
		//$Mtext = str_replace('\'', '\'\'', $Mtext);
	}

	if (isset($id)) {print "<input type=hidden name=row_id value='$id'>";}

	if ($action==2) {print "<p><b>РЕДАКТИРОВАНИЕ СООБЩЕНИЯ</b></p>";}
	else {print "<p><b>НОВОЕ СООБЩЕНИЕ</b></p>";}


	print "<p>Добавить значок к сообщению:</p>";

	if (isset($pic) & $pic==$img_src1) {
		print "<INPUT name='image' type='radio' value='$img_src1' checked> <img src='img1.gif' border=0 alt='Значок - Не работает'>";
	}
	else {
		print "<INPUT name='image' type='radio' value='$img_src1'> <img src='img1.gif' border=0 alt='Значок - Не работает'>";
	}
	print "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
	if (isset($pic) & $pic==$img_src2) {
		print "<INPUT name='image' type='radio' value='$img_src2' checked> <img src='img2.gif' border=0 alt='Значок - Внимание!'>";
	}
	else {
		print "<INPUT name='image' type='radio' value='$img_src2'> <img src='img2.gif' border=0 alt='Значок - Внимание!'>";
	}
	print "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
	if (isset($pic) & $pic==$img_src3) {
		print "<INPUT name='image' type='radio' value='$img_src3' checked> <img src='img4.gif' border=0 alt='Значок - Информация'>";
	}
	else {
		print "<INPUT name='image' type='radio' value='$img_src3'> <img src='img4.gif' border=0 alt='Значок - Информация'>";
	}

	print "<p>Краткое описание сообщения:<br>";
	if (isset($Mname)) {?> <textarea name='ShortText' rows=3 style='width:750'><?php echo $Mname; } 
	else { ?> <textarea name='ShortText' rows=3 style='width:750'> <?php } ?>
	</textarea>

	<p>Текст сообщения:<br>
	

	<p><input type='button' value=' B ' style='font-weight:600' class='toolbar' onclick='b()'> 
	<input type='button' value=' I ' style='font-style:italic;font-weight:600' class='toolbar' onclick='i()'> 
	<input type='button' value=' U&nbsp; ' style='text-decoration:underline;font-weight:600' class='toolbar' onclick='u()'> 

	| <a href="picture.php" onclick="window.open('picture.php','_blank','width=500,height=50,'+'location=no,toolbar=no,menubar=no,status=yes,scroll=yes');return false;" onMouseOut="window.status=''; return true;" title="закачать картинку">Вставить картинку</a> | 

	<a href="link.php" onclick="window.open('link.php','_blank','width=500,height=200,'+'location=no,toolbar=no,menubar=no,status=yes,scroll=yes');return false;" onMouseOut="window.status=''; return true;" title="создать ссылку">Создать ссылку</a></p>


	<?php

	if (isset($Mtext)) {print "<textarea name='text' rows=10 style='width:750'>".$Mtext."";}
	else {print "<textarea name='text' rows=10 style='width:750'>";}
	print "</textarea>";

	//print "<p align=center><input type=submit name=button1 value='Предпросмотр'> </p>";
	?>
	<p align=center><input type=button onclick="submit" name=button1 value='Предпросмотр'> </p>
	</form>
	<?php

	//-------------------------------------Создание предварительного просмотра сообщения-------------------------
	if ($button1) {
		print "<br><hr size=1 width=80%><br>";
		$message_name = $ShortText;
		$message_text = $text;
		$err = 0;

		if ($message_name=='') {
			print "<p align=center><b><font color=red>Не введено описание сообщения!</font></b></p>";
			$err = $err + 1;
		}
		if ($message_text=='') {
			print "<p align=center><b><font color=red>Не введен текст сообщения!</font></b></p>";
			$err = $err + 1;
		}

		if ($err==0) {
			print "<FORM ACTION='add.php' METHOD='post'>";
			?>

			<input type="hidden" name="message_text" value="<?= htmlspecialchars($message_text,ENT_QUOTES)?>">
			<input type="hidden" name="message_name" value="<?= htmlspecialchars($message_name,ENT_QUOTES)?>">
			
			<?php

			$ShortText = nl2br($ShortText);
			$ShortText = stripslashes($ShortText);
			$ShortText = str_replace('\'', '\'\'', $ShortText);
			$text = nl2br($text);
			$text = stripslashes($text);
			$text = str_replace('\'', '\'\'', $text);
			
			print "<a name=preview>";
			print "<p align=center><b>Сообщение для пользователей Интранета</b></p>";
			print "<table width=750 border=1><tr bgcolor=white><td style='border-style:solid; border-width:1px;'>";
			print "<p align=justify>";
			if (isset($image)) {print "<img src='".$image."' border=0 align=left vspace=0 hspace=10>";}
			print "$ShortText";
			print "<br><br><br><p align=justify>$text</p>";
			print "</td></tr></table>";
			print "</a>";
			
			print "<input type=hidden name=image value=$image>";
			print "<input type=hidden name=action value='$actionT'>";
			if ($actionT==2) {print "<input type=hidden name=id value='$row_id'>";}

			print "<p align=center><table border=0><tr><td style='border:none windowtext .5pt;'><input type=submit name=button2 value='Подтверждаю добавление объявления'></p></form></td>";

			print "<td style='border:none windowtext .5pt;'>";
				print "<form action=".$PHP_SELF." method=post enctype=multipart/form-data>";
				print "<input type=hidden name=pic value=$image>";
				?>
				
				<input type="hidden" name="Mname" value="<?= htmlspecialchars($message_name,ENT_QUOTES)?>">
				<input type="hidden" name="Mtext" value="<?= htmlspecialchars($message_text,ENT_QUOTES)?>">
				
				<?php
				print "<input type=hidden name=action value='$actionT'>";
				if ($actionT==2) {print "<input type=hidden name=id value='$row_id'>";}
				print "<input type=submit name=change value='   Изменить   '>";
				print "</form>";
			print "</td>";

			print "</tr>";
			print "</table>";
		}
	}
}
else {print "<p><b><font color=red>У вас нет доступа.</b>";}
?>

<p align=center><a href='http://intranet/' TARGET='_top'><img src='../../image/intra_m.gif' alt='На главную страницу' border='0'></a></p>
</td></tr></table>

</body>
</html>