<!--#include file="config.asp"-->
<!--#include file="ver.asp"-->
<!--#include file="Inc/Inc_Functions.asp"-->
<!--#include file="API/Api_Config.asp"-->
<%
'-----------------------------------------
'conn.asp
'数据库参数设置
'-----------------------------------------

'数据库类型:0-Access,1-Sql Server
Const Is_Sqldata	=	0

'使用外部数据库：0-不使用，1-使用
Const Is_ot_User	=	0

'数据库连接参数变量定义
Dim G_Sql_DelChar,G_Sql_Now,G_Sql_d_Char
Dim G_Sql_y,G_Sql_m,G_Sql_d,G_Sql_h,G_Sql_mi,G_Sql_s
Dim connstr,conn,db

'外部数据库参数变量定义
Dim ot_connstr,ot_conn,ot_usertable,ot_username,ot_password
Dim ot_regurl,ot_lostpasswordurl,ot_modIfypass1,ot_modIfypass2

'检验系统状态
Call SystemState

'连接数据库
Sub link_database()
	If Is_Sqldata = 0 Then
		'Access数据库连接参数
		'此处必须为以根目录开始,最前面必须为/号
		'免费用户初次安装务必修改DATA目录的数据库名称
		db =    "/data/oblog4.60.mdb"

		'-----------------------------------------------------------------------------------------------------
		'以下参数请勿修改，否则可能导致系统无法运行
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
		'Sql Server数据库连接参数
		Dim Sql_DBServer,Sql_DBName,Sql_User,Sql_Password
		Sql_DBServer	= "(local)"'"ZLOGCN\ZLOG"		'连接名(本地用(local),外地用IP如：127.0.0.1)
		Sql_DBName		= "oblog46"		'数据库名
		Sql_User		= "sa"			'访问数据的用户名
		Sql_Password	= "000000"		'访问数据的密码

		'-----------------------------------------------------------------------------------------------------
		'以下参数请勿修改，否则可能导致系统无法运行
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

	'若使用外部数据库表请自行修改下面的变量值
	If Is_ot_User=1 And InStr(LCase(Request.ServerVariables("HTTP_REFERER")),"admin_")=0 Then

		'access外部数据库连接字符串(默认连接方式)
		ot_connstr= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("/bbs/data/dvbbs7.mdb")

		'sql外部数据库连接字符串(sql server数据库请注释上边ACCESS数据库的连接字符串，取消注释SQL Server数据库连接字符串)
		'ot_connstr = "Provider = Sqloledb; User ID = bbs; Password = bbs; Initial Catalog = bbs; Data Source = (local);"

		'创建外部数据库连接
		Set ot_conn = Server.CreateObject("ADODB.Connection")
		ot_conn.open ot_connStr								'外部数据库连接
		ot_usertable		=	"dv_user"					'外部数据库用户表名
		ot_username			=	"username"					'外部数据库用户名字段
		ot_password			=	"userpassword"				'外部数据库密码字段
		ot_regurl			=	"../bbs/reg.asp"			'外部数据库注册用户链接
		ot_modIfypass1		=	"../bbs/modIfyadd.asp?t=1"	'外部数据库修改密码连接
		ot_modIfypass2		=	"../bbs/modIfyadd.asp?t=1"	'外部数据库修改密码提示问题连接
		ot_lostpasswordurl	=	"../bbs/lostpass.asp"		'外部数据库找回密码链接
	End If

	If Err Then
		'Err.clear
		Set conn = Nothing
	ECHO_ERR "Connection Database","<B>连接数据库出错</B><br/>您可能没有正确设置  conn.asp  的数据库连接。<br/>如果您是Sql Server 数据库的话，也可能是数据库名、数据库用户名、或者数据库密码不正确，还有可能是数据库服务没有开启。",1
	End If
End Sub

%>