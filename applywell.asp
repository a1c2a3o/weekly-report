<%Response.Charset="GB2312"%>
<!--#include file="constant.asp"-->
<% person=Request.Form("person") %>
<% password=Request.Form("password") %>
<% title=person & " �������ȵ�¼��" %>
<!--#include file="head.asp"-->
<!--#include file="function.asp"-->

<p align=center>
�װ��� <font color=green><b><%=person%></b></font>�������˺���ע��ɹ������μ����롣
<br>
����󽫻ص���ҳ��
<p>


<%

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "DBQ=" & Server.MapPath(memberDb) & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25;FIL=MS Access;"
sql = "insert into MIR (chineseName, birthday,active) values('"
sql = sql & Request("person") & "', '"
sql = sql & Request("password") & "', "
sql = sql & "True)"
Conn.Execute(sql)

%>




<meta http-equiv="refresh" content="5;url=default.asp">