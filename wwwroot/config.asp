<%
Option Explicit
Response.Buffer = True
'Response.CodePage = 936
'WIN2000服务器不支持此属性
Response.CharSet = "GB2312"

'全局常量或变量、参数定义程序段

'是否开启调试模式,True为开启，False为关闭
Const Is_Debug	=	true
'-----------------------------------------------------------------------------------------------------
'以下参数非常重要，必须认真填写，运行系统后请勿任意更改，防止运行出错

'blog程序所在目录,如为根目录请改为"/"
Const blogdir			=	"/"
'生成的日志静态文件后缀,可以为htm,html,shtml,asp四种格式
Const f_ext				=	"html"
'cookies名,一般无须修改
Const cookies_name		=	"oblog45"
'cookeies域名根,一般留空
Const cookies_domain	=	""
'系统缓存名前缀，一般无须修改
Const cache_name_user	=	"oblog45"
'是否启用真实二级域名,需OblogDns组件支持
Const true_domain	=	0
Dim blogurl, str_domain
If true_domain		=	1 Then
	'真实域名必填，设置blog程序绝对路径，即访问blog首页的绝对URL
	'非同域名根请将下边的C_Editor_Type参数改成1，否则会使得编辑器不正常
	'同域名将blogurl的值设置成blogdir则不受以上条件限制
	blogurl			=	"http://5.4sk.cn/"
'	blogurl			=	 blogdir
	str_domain		=	",custom_domain"
Else
	blogurl			=	blogdir
End If
'http://support.microsoft.com/kb/269238/
Const MsxmlVersion = ".3.0"
'设置自定义验证码生成文件路径,相对博客站点根目录,一般改文件名即可.
Const IncCodePath = "inc/code.asp" 

'是否打开lightbox图片特效,修改此设置用户相册即时生效不用刷新文件.
Const Islightbox=0
'非常重要 这里的密钥组合在您正式使用的时候必须更改为您自己的 .
'站点 唯一的加密解密key,请尽量选择较小的key比如都是四位的key.您的Key请不要在公共场合公开.获取key请访问生成工具的地址
'/非常重要的密钥生成地址  http://www.oblog.cn/RsaKeyPair.asp
Const RsaKeyCode="9307,4243,7831"
'-----------------------------------------------------------------------------------------------------
'分页公用变量
Dim G_P_PerMax,G_P_AllRecords,G_P_AllPages,G_P_This,G_P_FileName,G_P_Guide

Const P_BLOG_UPDATEPAUSE= 50 		'全站更新时，每更新多少篇日志暂停一次(100以内)
Const P_TAGS_SPLIT		= " "		'TAG Split
Const EN_NameIsNum		= 0			'是否允许全数字的用户名,1为允许,0为不允许
Const SYSFOLDER_ADMIN	= "admin"	'该目录名称将被作为系统禁止注册的用户名使用
Const SYSFOLDER_MANAGER = "manager" '该目录名称将被作为系统禁止注册的用户名使用
Const En_OutRss			= 1			'是否允许订阅站外rss源
Const En_Recycle		= 1			'是否启用回收站功能(只对日志及用户数据启用,回复、评论、文件等不启用)
Const Str_HtmlFilt		= ""		'自定义过滤的html代码，必须以|结束，格式如aaa|bbb|(对评论和留言有效)
Const C_UserIcon_Width	="48"
Const C_UserIcon_Height	="48"

'预留参数
Const C_Vote_Action1	="鲜花"
Const C_Vote_Action2	="板砖"

'编辑器路径
Dim C_Editor,C_Editor_Type,C_Editor_LoadIcon,C_Editor_UBB

' 1 为4.x编辑器 2为3.x编辑器
C_Editor_Type = 1

'以下参数请勿更改！！
Select Case C_Editor_Type
	Case 1
		C_Editor=blogdir&"editor"
		C_Editor_LoadIcon="yes"
	Case 2
		C_Editor=blogdir&"editor2"
		C_Editor_LoadIcon="none"
End Select
C_Editor_UBB=blogdir&"editor"
'-----------------------------------------------------------------------------------------------------
%>