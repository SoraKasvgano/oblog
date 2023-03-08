<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
'Oblog4 所有验证码匹配操作
Dim rs,Sql,sObCode,creatuserid,username
sObcode=Request("sn")
username=Trim(request("user"))
If CheckObCode(sObcode)=false Then
	Response.Write "您输入的验证码错误!"
	Response.End
End If
If Not IsObject(conn) Then link_database
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "select * From oblog_obcodes Where obcode='" & LCase(sObCode) & "'",conn,1,3
If rs.Eof Then
	rs.Close
	Set rs=Nothing
	Set oblog=Nothing
	Response.Write "您输入的验证码不存在!"
	Response.End
End If
creatuserid=rs("creatuser")
If rs("istate")=1 Then
	rs.Close
	Set rs=Nothing
	Set oblog=Nothing
	Response.Write "您输入的验证码已经被使用!"
	Response.End
End If

'根据类型进行处理
select Case rs("itype")
	Case 1
		'帐号验证
		oblog.Execute "Update oblog_user Set EmailValid=2,user_level=7 Where userid=" & creatuserid
		rs("istate")=1
		rs("useuser")=creatuserid
		rs("usetime")=Now
		rs("useip")=oblog.UserIp
		rs.Update
		
		Session ("CheckUserLogined_"&username) = ""
		Oblog.CheckUserLogined()
		Response.Write "<a href=""login.asp"">您帐号已经被验证通过,请登录系统进行相关操作<a>"
	Case 2
		'邮件验证
		oblog.Execute "Update oblog_user Set EmailValid=2,user_level=7 Where userid=" &  creatuserid
		rs("istate")=1
		rs("useuser")= creatuserid
		rs("usetime")=Now
		rs("useip")=oblog.UserIp
		rs.Update
		rs.Update
		Response.Write "您的邮件被确认为有效"
	Case 3
		'密码验证
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