<%
'Корректрировка эстонских символов
function correct_est(a)
l=len(a)
sl=""
For i=1 To l
  c=Mid(a,i,1)
  v=asc(c)
  SELECT CASE v
    Case 245
      sl=sl & "&otilde;"
    Case 228
      sl=sl & "&auml;"
    Case 246
      sl=sl & "&ouml;"
    Case 252
      sl=sl & "&uuml;"
    Case 213
      sl=sl & "&Otilde;"
    Case 196
      sl=sl & "&Auml;"
    Case 214
      sl=sl & "&Ouml;"
    Case 220
      sl=sl & "&Uuml;"
    Case Else
      sl=sl & c
  END SELECT
Next
correct_est=sl
End function
 
'Вычисление финансового года на основе текущей даты
Function fin_year()
ya=Year(Date())
mo=Month(Date())
da=Day(Date())
zz=mo-04
If zz>0 Then
  ya=Year(Date())
Else
  ya=ya-1
End If
fin_year=ya
End Function

'Обновление сумм в проектах, при сохранении изменений в карточке
Sub Update_main(i, c)
    ay="yr" & i
    a0="ps" & i

    mdboen.CommandText="UPDATE Main SET PastSum=" & Request.Form(a0) & " WHERE Pid='" & Mid(proc,6,5) & "' AND IDentifier='" & c & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(Proc,5,1) & "'"
    mdboren.Open mdboen
    a0="1k" & i
    mdboen.CommandText="UPDATE Main SET Ikvartal=" & Request.Form(a0) & " WHERE Pid='" & Mid(proc,6,5) & "' AND IDentifier='" & c & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(Proc,5,1) & "'"
    mdboren.Open mdboen
    a0="2k" & i
    mdboen.CommandText="UPDATE Main SET IIkvartal=" & Request.Form(a0) & " WHERE Pid='" & Mid(proc,6,5) & "' AND IDentifier='" & c & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(Proc,5,1) & "'"
    mdboren.Open mdboen
    a0="3k" & i
    mdboen.CommandText="UPDATE Main SET IIIkvartal=" & Request.Form(a0) & " WHERE Pid='" & Mid(proc,6,5) & "' AND IDentifier='" & c & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(Proc,5,1) & "'"
    mdboren.Open mdboen
    a0="4k" & i
    mdboen.CommandText="UPDATE Main SET IVkvartal=" & Request.Form(a0) & " WHERE Pid='" & Mid(proc,6,5) & "' AND IDentifier='" & c & "' AND Yearr='" & Mid(proc,1,4) & "' AND Enterprise='" & Mid(Proc,5,1) & "'"
    mdboren.Open mdboen
End Sub

'Обновление сумм в композитных проектах
Sub Update_main_composite(c)
  mdboem.CommandText="SELECT ISNULL(SUM(PastSum),0) as s0, ISNULL(SUM(Ikvartal),0) as s1,ISNULL(SUM(IIkvartal),0) as s2, ISNULL(SUM(IIIkvartal),0) as s3,ISNULL(SUM(IVkvartal),0) as s4 FROM Main WHERE ProjCode LIKE '" & MID(mdbor("ProjCode"),1,9) & "__' AND ProjCode <> '" & mdbor("ProjCode") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
  mdboren.Open mdboen
  mdboco.CommandText="UPDATE Main SET PastSum=" & Mdboren("s0") & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
  mdborco.Open mdboco
  mdboco.CommandText="UPDATE Main SET Ikvartal=" & Mdboren("s1") & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
  mdborco.Open mdboco
  mdboco.CommandText="UPDATE Main SET IIkvartal=" & Mdboren("s2") & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
  mdborco.Open mdboco
  mdboco.CommandText="UPDATE Main SET IIIkvartal=" & Mdboren("s3") & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
  mdborco.Open mdboco
  mdboco.CommandText="UPDATE Main SET IVkvartal=" & Mdboren("s4") & " WHERE Pid='" & mdbor("Pid") & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='F'"
  mdborco.Open mdboco
  mdboren.Close
End Sub

Function Select_sum(col, mes, i)
  mdboen.CommandText="SELECT ROUND(ISNULL(SUM(ISNULL(" & col & ",0)),0)/1000,0) AS summi,PROJECT FROM glav_project WHERE PROJECT='" & Mdbor("OracleCode") & "' AND (KONTO NOT BETWEEN 18410 AND 18430) AND DEBET IS NOT NULL AND LEFT(MES,1)=" & Mid(proc,4,1)+i & " AND RIGHT(MES,2) BETWEEN " & mes & " GROUP BY PROJECT ORDER BY PROJECT DESC"
  mdboren.Open mdboen
  If mdboren.EOF=True Then
    va1=0
  Else
    va1=mdboren("summi")
  End If
  mdboren.close
  Select_sum=va1
End Function

Sub Contract_sum_update(kv, mes1, mes2, y)
  mdboen.CommandText="SELECT SUM(SummOfContr) AS Summi FROM Contra WHERE Pid='" & mid(proc,6,5) & "' AND MONTH(DateOfConcl)>='" & mes1 & "' AND MONTH(DateOfConcl)<='" & mes2 & "' AND YEAR(DateOfConcl)='" & CDbl(Mid(proc,1,4))+y & "'"
  mdboren.Open mdboen
  If mdboren.EOF=True Or mdboren("summi") & "e" = "e" Then
    vu1=0
  Else
    vu1=mdboren("summi")
  End If
  mdboren.Close
  mdboen.CommandText="UPDATE Main SET " & kv & "kvartal=" & vu1 & " WHERE Pid='" & MID(proc,6,5) & "' AND Enterprise='" & mdbor("Enterprise") & "' AND Yearr='" & Mid(proc,1,4) & "' AND Identifier='C'"
  mdboren.Open mdboen
End Sub
%>