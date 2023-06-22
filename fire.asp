<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>林嘉英1410903049</title>
</head>
<body>
<p>竊盜險(火災保險附加險)查詢結果</p>
<%                   'asp程式由小於符號及百分比符號為開始百分比符號及大於符號為終止
                     '在單引號後面的內容在asp程式中不會執行僅是說明功能
   rate1=request("rate1")
   rate2=request("rate2")
   house_price=request("house_price")                     'asp程式執行時必須先定義變數名稱(本例為sql)sql=request("sql")然後才可以執行變數名稱sql
sql=request("sql")   '"等於"符號後面的sql表示在car.htm 中HTML之名稱(N)
DbPath = SERVER.MapPath("fire.mdb")                                    '設定資料庫路徑(本例為autobrand.mdb)（使用Access資料庫時必須更改資料庫名稱）
Set conn = Server.CreateObject("ADODB.Connection")                    '設定連結(此行敘述為固定不用更改)           
conn.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DbPath    '開啟資料庫(使用Access資料庫時，此行敘述為固定不用更改)
sql="select * from ns where house_price=" & rate1 & "and times=' & rate2 & '"
response.write "竊盜險(火災保險附加險)的SQL查詢語法：<font color=red>" & sql &"</font><p>"
set rs=conn.Execute(sql)  '：執行指定的查詢，並將查詢結果放入rs中
%>
	<table width="893" border="1">
  <tbody>
    <tr>
       <td width="122">保險公司</td>
		<td width="122">房價</td>
      <td width="45">時間</td>
      <td width="76">保險費</td>
      <td width="150">通路一附加險保險費</td>
		 <td width="150">通路二附加險保險費</td>
      
    </tr>
    
  
<p>&nbsp;</p>
<p>
  
<%	 
  while not rs.eof  
   response.write "<tr>"
response.write "<td>" & rs("company") & "</td>" & "<td>" & rs("house_price") & "</td>" & "<td>" & rs("time") & "</td>" & "<td>" & rs("house_premium") & "</td>" & "<td>" & rs("rate1_premium") & "</td>"& "<td>" & rs("rate2_premium") & "</td>"
response.write "</tr>"
   rs.movenext          
wend                                                                 %>   
	  </tbody>
	</table> 
  
<%
  conn.Close '關閉資料庫連接
%>
</body>
</html>                                                                    '說明  While...Wend  迴圈：               
																	  ' 如果執行指定的查詢條件式為True，則所有的陳述式都會執行，一直執行到 Wend
																	  ' 陳述式。然後再回到 While 陳述式，並再一次檢查條件式，如果條件式還是為
																	  ' True，則重複此步驟，如果條件式不為 True，則程式會從 Wend
																	  ' 陳述式之後的指令行繼續執行。
%>