<!--#include file="inc/inc_syssite.asp"-->
<%Dim Rs , Sql 
'------------------------------------------------
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------

Const Log_Num = 10    '����������Ա�������ʵ���Ҫ�����ٵ�����������,Ϊ0Ϊ�����ơ�

Const Com_Num = 10  '����������Ա�������ʵ���Ҫ�����ٵ���������������,Ϊ0Ϊ�����ơ�

Const Is_Only_Best = False 'True False �Ƿ�ֻ��������Ƽ��û�


Sql="SELECT TOP 1 UserName FROM Oblog_User WHERE lockuser = 0 AND isdel=0 AND (is_log_default_hidden=0 or is_log_default_hidden is null)"
	If Log_Num > 0  Then Sql = Sql & " AND log_count > "&Log_Num&" "
	If Com_Num > 0  Then Sql = Sql & " AND comment_count > "&Com_Num&" "
	If Is_Only_Best Then Sql = Sql & " AND user_isbest = 1 "
	If CBool(Is_Sqldata)   Then 
		Sql = Sql &" ORDER BY NEWID()"
	Else
		Randomize
		Sql = Sql &" ORDER BY Rnd(-(UserId+"&Rnd()&"))"
	End If 
Set Rs=oblog.Execute(Sql)
	If Not (Rs.eof Or Rs.bof) Then Response.Redirect(BlogDir&"blog.asp?name="&Rs(0))
Set Rs=Nothing 
%>