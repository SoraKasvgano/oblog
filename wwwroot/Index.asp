<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%'����Ա��¼�������Ĭ����ҳΪ����ģʽ�����ɾ�̬ҳ.�������Ա����ģ��.
Dim oFSO,Is_debug_index
'�Ƿ����ù���Ա��¼״̬�µĵ���ģʽ.
Const  IsOpenAdminIndexDebug=false 



'�����������ҳ��̬�ļ�
Const index_html_name = "index.html"


Is_debug_index=False 


If session("adminname")<>"" And IsOpenAdminIndexDebug Then Is_debug_index=True 
If Is_debug_index  Then DO_index(1)
'�ж��Ƿ���Ҫ����������־��
If IsNull(oblog.Setup(11,0)) Or IsNull(oblog.Setup(10,0)) Then
	YesterDay_Log_Count
Else
	If DateDiff("d",CDate(oblog.Setup(11,0)),Date())>1 Then
		YesterDay_Log_Count
	End if
End If
Set oFSO = Server.CreateObject(oblog.CacheCompont(1))
If Application(oblog.cache_name&"_index_update")=False And oFSO.FileExists(Server.mappath(index_html_name)) and DateDiff("s",Application(oblog.cache_name&"_index_updatetime"),Now())<Int(oblog.CacheConfig(33)) Then
	Set oFSO = Nothing
	RedirectBy301(index_html_name)
Else
	'�ȴ����棬��ֹ���ɵ�ʱ��Ч�ʹ��͵��²�������̲���������ҳ�����
	Application.Lock
	Application(oblog.cache_name&"_index_update") = False
	Application(oblog.cache_name&"_index_updatetime") = Now()
	Application(oblog.cache_name&"_list_update") = True
	Application(oblog.cache_name&"_class_update") = True
	Application.unLock

	Set oFSO = Nothing
	DO_index(0)

	If Request("re")<>"0" Then
		RedirectBy301(index_html_name)
	End If
End If
'ͳ��������־��
Sub YesterDay_Log_Count()
	On Error Resume Next
	Dim rs,rst,sql
	If Not IsObject(conn) Then link_database
	Set rs = Server.CreateObject("adodb.recordset")
	rs.open "select log_count_Yesterday,log_Yesterday FROM oblog_setup",conn,1,3
	sql = "select COUNT(logid) FROM oblog_log WHERE isdel=0 "
	If Is_Sqldata = 0 Then
		sql = sql & " AND DATEDIFF("&G_Sql_d&",truetime,"&G_Sql_Now&")=1"
	Else
		sql = sql & " AND truetime>=CONVERT(CHAR(10),GETDATE()-1,120) AND truetime < CONVERT(CHAR(10),GETDATE(),120)"
	End if
	Set rsT = oblog.Execute(SQL)
	rs(0) = rsT(0)
	rs(1) = DateAdd("d",-1,Date())
	rs.Update
	rs.close
	Set rsT= Nothing
	Set rs = Nothing
	oblog.ReloadSetup
End Sub
Sub DO_index(is_debug)
	Dim rstmp,sContent,sStyle
	Set rstmp=oblog.execute("select skinmain from oblog_sysskin where isdefault=1")
	sContent=rstmp(0)
	sStyle=OB_PickUpCss(sContent)
	'G_P_Show="" 
	G_P_Show=Replace(G_P_Show,"{OB_STYLE}",sStyle)
	G_P_Show=G_P_Show&sContent
	'G_P_Show=G_P_Show&"<script src=""index.asp""></script>"
	Set rstmp=Nothing
	Call indexshow()
	G_P_Show=G_P_Show&oblog.site_bottom

	If Cbool(is_debug) And Request("re")<>"0" Then 
		response.write G_P_Show
		response.End
		Exit Sub 
	Else 
	Call oblog.BuildFile(Server.mappath(index_html_name),G_P_Show)
	End If 
End Sub  
%>