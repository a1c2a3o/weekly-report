<%Response.Charset="GB2312"%>

<!-- (ע��)���ļ��� GB2UTF8 1.2 Beta ������ 2009-3-31 12:46:59��������ߣ�����(fxy_2002@163.com) pc-soft.cn -->
<!--#include file="constant.asp"-->
<%title="MIR ʵ���ҹ������ȵ�¼��ҳ"%>
<!--#include file="head.asp"-->
<!--#include file="function.asp"-->

<h3><img src="redball.gif">
��д����</h3>

<% students = getMirStudentName %>

<center>
<form method=post action="form.asp">ѡ�����������
<select name="person">

<% For i = 0 to UBound(students)
	Response.Write("<option>" & students(i))
Next %>
</select>
<br><br>
<align=center>����: <input type=password size=10 name="password">
<input type=submit value="��¼">
</form>
</center>


<center>�����˺ţ�����<a target=_parent href="apply.asp">���˴�ע�᡹</a></center>

<ol>
<li>�������ÿ�������������ǰ��д��ϡ�������������ҹ��ϵͳ�Զ�������һ�ܣ����޷�����д���ܵĽ����ˡ�
<li>�����ÿһ����Ҫ��д�������ǡ�����Ԥ����������һ��Ҫ������صġ�Ԥ�����ʱ�䡹��
</ol>
<hr>

<h3><img src="redball.gif">
�����б�</h3>
<ul>
<li>ÿ����д֮���ϣ�
	<a target=_blank href="listEachWeek.asp?weekDiff=0">����</a>��
	<a target=_blank href="listEachWeek.asp?weekDiff=-1">ǰһ��</a>��
	<a target=_blank href="listEachWeek.asp?weekDiff=-2">ǰ����</a>��
	<a target=_blank href="listEachWeek.asp?weekDiff=-3">ǰ����</a>��
	<a target=_blank href="listEachWeek.asp?weekDiff=-4">ǰ����</a>
<li>ÿ���˵���ʷ���ϣ�
<% For i = 0 to UBound(students)
	If i=0 Then
		Response.Write("<a target=_blank href=listEachPerson.asp?person=" & students(i) & ">" & students(i) & "</a>")
	Else
		Response.Write("��<a target=_blank href=listEachPerson.asp?person=" & students(i) & ">" & students(i) & "</a>")
	End If
Next %>
<li><a target=_blank href="listAllPersonLastRecord.asp">ÿ���˵�����һ������</a>
</ul>
<hr>



<!!--#include file="foot.asp"-->