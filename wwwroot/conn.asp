<!--#include file="config.asp"-->
<!--#include file="ver.asp"-->
<!--#include file="Inc/Inc_Functions.asp"-->
<!--#include file="API/Api_Config.asp"-->
<%
'-----------------------------------------
'conn.asp
'���ݿ��������
'-----------------------------------------

'���ݿ�����:0-Access,1-Sql Server
Const Is_Sqldata	=	0

'ʹ���ⲿ���ݿ⣺0-��ʹ�ã�1-ʹ��
Const Is_ot_User	=	0

'���ݿ����Ӳ�����������
Dim G_Sql_DelChar,G_Sql_Now,G_Sql_d_Char
Dim G_Sql_y,G_Sql_m,G_Sql_d,G_Sql_h,G_Sql_mi,G_Sql_s
Dim connstr,conn,db

'�ⲿ���ݿ������������
Dim ot_connstr,ot_conn,ot_usertable,ot_username,ot_password
Dim ot_regurl,ot_lostpasswordurl,ot_modIfypass1,ot_modIfypass2

'����ϵͳ״̬
Call SystemState

'�������ݿ�
Sub link_database()
	If Is_Sqldata = 0 Then
		'Access���ݿ����Ӳ���
		'�˴�����Ϊ�Ը�Ŀ¼��ʼ,��ǰ�����Ϊ/��
		'����û����ΰ�װ����޸�DATAĿ¼�����ݿ�����
		db =    "/data/oblog4.60.mdb"

		'-----------------------------------------------------------------------------------------------------
		'���²��������޸ģ�������ܵ���ϵͳ�޷�����
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
		G_Sql_d_Char	="#"
		G_Sql_y			="'yyyy'"
		G_Sql_m			="'m'"
		G_Sql_d			="'d'"
		G_Sql_h			="'h'"
		G_Sql_mi		="'n'"
		G_Sql_s			="'s'"
		G_Sql_Now		="Now()"
		'-----------------------------------------------------------------------------------------------------
	Else
		'Sql Server���ݿ����Ӳ���
		Dim Sql_DBServer,Sql_DBName,Sql_User,Sql_Password
		Sql_DBServer	= "(local)"'"ZLOGCN\ZLOG"		'������(������(local),�����IP�磺127.0.0.1)
		Sql_DBName		= "oblog46"		'���ݿ���
		Sql_User		= "sa"			'�������ݵ��û���
		Sql_Password	= "000000"		'�������ݵ�����

		'-----------------------------------------------------------------------------------------------------
		'���²��������޸ģ�������ܵ���ϵͳ�޷�����
		ConnStr			= "Provider = Sqloledb; User ID = " & Sql_User & "; Password = " & Sql_Password & "; Initial Catalog = " & Sql_DBName & "; Data Source = " & Sql_DBServer & ";"
		G_Sql_d_Char	="'"
		G_Sql_y			="Year"
		G_Sql_m			="Month"
		G_Sql_d			="Day"
		G_Sql_h			="Hour"
		G_Sql_mi		="Minute"
		G_Sql_s			="Second"
		G_Sql_Now		="GetDate()"
		'-----------------------------------------------------------------------------------------------------
	End If
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr

	'��ʹ���ⲿ���ݿ���������޸�����ı���ֵ
	If Is_ot_User=1 And InStr(LCase(Request.ServerVariables("HTTP_REFERER")),"admin_")=0 Then

		'access�ⲿ���ݿ������ַ���(Ĭ�����ӷ�ʽ)
		ot_connstr= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("/bbs/data/dvbbs7.mdb")

		'sql�ⲿ���ݿ������ַ���(sql server���ݿ���ע���ϱ�ACCESS���ݿ�������ַ�����ȡ��ע��SQL Server���ݿ������ַ���)
		'ot_connstr = "Provider = Sqloledb; User ID = bbs; Password = bbs; Initial Catalog = bbs; Data Source = (local);"

		'�����ⲿ���ݿ�����
		Set ot_conn = Server.CreateObject("ADODB.Connection")
		ot_conn.open ot_connStr								'�ⲿ���ݿ�����
		ot_usertable		=	"dv_user"					'�ⲿ���ݿ��û�����
		ot_username			=	"username"					'�ⲿ���ݿ��û����ֶ�
		ot_password			=	"userpassword"				'�ⲿ���ݿ������ֶ�
		ot_regurl			=	"../bbs/reg.asp"			'�ⲿ���ݿ�ע���û�����
		ot_modIfypass1		=	"../bbs/modIfyadd.asp?t=1"	'�ⲿ���ݿ��޸���������
		ot_modIfypass2		=	"../bbs/modIfyadd.asp?t=1"	'�ⲿ���ݿ��޸�������ʾ��������
		ot_lostpasswordurl	=	"../bbs/lostpass.asp"		'�ⲿ���ݿ��һ���������
	End If

	If Err Then
		'Err.clear
		Set conn = Nothing
	ECHO_ERR "Connection Database","<B>�������ݿ����</B><br/>������û����ȷ����  conn.asp  �����ݿ����ӡ�<br/>�������Sql Server ���ݿ�Ļ���Ҳ���������ݿ��������ݿ��û������������ݿ����벻��ȷ�����п��������ݿ����û�п�����",1
	End If
End Sub

%>