<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="Inc/Class_TrackBack.asp" -->
<%
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
ON Error Resume Next
If Not lcase(Request.ServerVariables("REQUEST_METHOD"))="post" Then Response.End
If oblog.CacheConfig(54) = "1" Then Response.write("ϵͳ���ý�ֹ����ͨ�湦��!"):Response.End
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(3)="1" Then	Response.write("ϵͳ��ʱ��ֹʹ������ͨ�湦��!"):Response.End
End if
Dim objTrackback
Dim LogId,IP,url,title,BlogName,Excerpt,rst,rstCache
'�������ݵĲ���������XML����ֵ
'IP���
oblog.chk_commenttime
'tb.asp?id=53&TBcode=200703210942u85Csg1RRm6O&url=http://lj/oblog41/go.asp&blog_name=atai&title=������&excerpt=����
LogId=Request("id")
logId=CLng(LogId)
IP=GetIP
Url=Trim(Request("url"))
title=Trim(Request("title"))
BlogName=Trim(Request("blog_name"))
Excerpt=Trim(Request("excerpt"))
If url=blogdir&"tb.asp" Then
'���urlΪ����ֹͣ��Ӧ
	Response.End
End if
'���ݼ��
if oblog.chk_badword(url)>0 then oblog.adderrstr("��ַ�к���ϵͳ��������ַ���")
if oblog.chk_badword(title)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
if oblog.chk_badword(BlogName)>0 then oblog.adderrstr("BLOG�����к���ϵͳ��������ַ���")
if oblog.chk_badword(Excerpt)>0 then oblog.adderrstr("ժҪ�к���ϵͳ��������ַ���")
'ר���ؼ����ж�
if oblog.errstr<>"" Then oblog.showerr(): Response.End()
Call Link_Database
'Ƶ�ȼ��,���ͬһIP�ڵ�λʱ���ڷ�����ͨ������ﵽһ���޶���Զ���IP
'If oblog.ChkWhiteIP(IP) = False Then
	Set rst=oblog.Execute("select count(id) From Oblog_trackback Where ip='" & IP & "' And datediff("&G_Sql_mi&",addtime,"&G_Sql_Now&")<="&oblog.CacheConfig(66))
	If rst(0)> Int(oblog.CacheConfig(65)) Then
		'���������
		oblog.KillIp(IP)
'		oblog.ShowMsg "��Ϊ����һЩ������ϵͳ�����˸��ţ����IP�����������",""
		Response.End
	End If
	rst.Close
	Set rst=Nothing
'End if

'���н��ջ��ڵĴ���
Set objTrackback = New Class_TrackBack
objTrackback.LOGID=LogId
objTrackback.IP=IP
objTrackback.URL=Url
'objTrackback.TBUSER=Trim(Request.QueryString("tbuser"))
objTrackback.TITLE=title
objTrackback.BLOG_NAME=BlogName
objTrackback.EXCERPT=Excerpt
Response.Cookies(cookies_name)("LastComment") = oblog.ServerDate(Now())
If objTrackback.CheckTB (LCase(Trim(Request("TBcode")))) Then Call objTrackback.Receive()
Set objTrackback=Nothing
conn.Close
Set conn=Nothing
%>