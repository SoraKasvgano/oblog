<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
'Oblog4 ������֤��ƥ�����
Dim rs,Sql,sObCode,creatuserid,username
sObcode=Request("sn")
username=Trim(request("user"))
If CheckObCode(sObcode)=false Then
	Response.Write "���������֤�����!"
	Response.End
End If
If Not IsObject(conn) Then link_database
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "select * From oblog_obcodes Where obcode='" & LCase(sObCode) & "'",conn,1,3
If rs.Eof Then
	rs.Close
	Set rs=Nothing
	Set oblog=Nothing
	Response.Write "���������֤�벻����!"
	Response.End
End If
creatuserid=rs("creatuser")
If rs("istate")=1 Then
	rs.Close
	Set rs=Nothing
	Set oblog=Nothing
	Response.Write "���������֤���Ѿ���ʹ��!"
	Response.End
End If

'�������ͽ��д���
select Case rs("itype")
	Case 1
		'�ʺ���֤
		oblog.Execute "Update oblog_user Set EmailValid=2,user_level=7 Where userid=" & creatuserid
		rs("istate")=1
		rs("useuser")=creatuserid
		rs("usetime")=Now
		rs("useip")=oblog.UserIp
		rs.Update
		
		Session ("CheckUserLogined_"&username) = ""
		Oblog.CheckUserLogined()
		Response.Write "<a href=""login.asp"">���ʺ��Ѿ�����֤ͨ��,���¼ϵͳ������ز���<a>"
	Case 2
		'�ʼ���֤
		oblog.Execute "Update oblog_user Set EmailValid=2,user_level=7 Where userid=" &  creatuserid
		rs("istate")=1
		rs("useuser")= creatuserid
		rs("usetime")=Now
		rs("useip")=oblog.UserIp
		rs.Update
		rs.Update
		Response.Write "�����ʼ���ȷ��Ϊ��Ч"
	Case 3
		'������֤
		Session("userid")= creatuserid
		Session("usercode")=sObCode
		Response.Redirect ""
End select

Function CheckObCode(sCode)
	Dim i,iAsc
	sCode=UCase(Trim(sCode))
	CheckObCode=false
	If Len(sCode)<>32 Then Exit Function
 	For i = 1 To Len(sCode)
    	iAsc = Asc(Mid(sCode, i, 1))
        '48~57,65~90
        If iAsc < 48 Or (iAsc > 57 And iAsc < 65) Or iAsc > 90 Then Exit Function
     Next
     CheckObCode=true
End Function


%>