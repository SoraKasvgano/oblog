<%
Option Explicit
Response.Buffer = True
'Response.CodePage = 936
'WIN2000��������֧�ִ�����
Response.CharSet = "GB2312"

'ȫ�ֳ����������������������

'�Ƿ�������ģʽ,TrueΪ������FalseΪ�ر�
Const Is_Debug	=	true
'-----------------------------------------------------------------------------------------------------
'���²����ǳ���Ҫ������������д������ϵͳ������������ģ���ֹ���г���

'blog��������Ŀ¼,��Ϊ��Ŀ¼���Ϊ"/"
Const blogdir			=	"/"
'���ɵ���־��̬�ļ���׺,����Ϊhtm,html,shtml,asp���ָ�ʽ
Const f_ext				=	"html"
'cookies��,һ�������޸�
Const cookies_name		=	"oblog45"
'cookeies������,һ������
Const cookies_domain	=	""
'ϵͳ������ǰ׺��һ�������޸�
Const cache_name_user	=	"oblog45"
'�Ƿ�������ʵ��������,��OblogDns���֧��
Const true_domain	=	0
Dim blogurl, str_domain
If true_domain		=	1 Then
	'��ʵ�����������blog�������·����������blog��ҳ�ľ���URL
	'��ͬ�������뽫�±ߵ�C_Editor_Type�����ĳ�1�������ʹ�ñ༭��������
	'ͬ������blogurl��ֵ���ó�blogdir����������������
	blogurl			=	"http://5.4sk.cn/"
'	blogurl			=	 blogdir
	str_domain		=	",custom_domain"
Else
	blogurl			=	blogdir
End If
'http://support.microsoft.com/kb/269238/
Const MsxmlVersion = ".3.0"
'�����Զ�����֤�������ļ�·��,��Բ���վ���Ŀ¼,һ����ļ�������.
Const IncCodePath = "inc/code.asp" 

'�Ƿ��lightboxͼƬ��Ч,�޸Ĵ������û���ἴʱ��Ч����ˢ���ļ�.
Const Islightbox=0
'�ǳ���Ҫ �������Կ���������ʽʹ�õ�ʱ��������Ϊ���Լ��� .
'վ�� Ψһ�ļ��ܽ���key,�뾡��ѡ���С��key���綼����λ��key.����Key�벻Ҫ�ڹ������Ϲ���.��ȡkey��������ɹ��ߵĵ�ַ
'/�ǳ���Ҫ����Կ���ɵ�ַ  http://www.oblog.cn/RsaKeyPair.asp
Const RsaKeyCode="9307,4243,7831"
'-----------------------------------------------------------------------------------------------------
'��ҳ���ñ���
Dim G_P_PerMax,G_P_AllRecords,G_P_AllPages,G_P_This,G_P_FileName,G_P_Guide

Const P_BLOG_UPDATEPAUSE= 50 		'ȫվ����ʱ��ÿ���¶���ƪ��־��ͣһ��(100����)
Const P_TAGS_SPLIT		= " "		'TAG Split
Const EN_NameIsNum		= 0			'�Ƿ�����ȫ���ֵ��û���,1Ϊ����,0Ϊ������
Const SYSFOLDER_ADMIN	= "admin"	'��Ŀ¼���ƽ�����Ϊϵͳ��ֹע����û���ʹ��
Const SYSFOLDER_MANAGER = "manager" '��Ŀ¼���ƽ�����Ϊϵͳ��ֹע����û���ʹ��
Const En_OutRss			= 1			'�Ƿ�������վ��rssԴ
Const En_Recycle		= 1			'�Ƿ����û���վ����(ֻ����־���û���������,�ظ������ۡ��ļ��Ȳ�����)
Const Str_HtmlFilt		= ""		'�Զ�����˵�html���룬������|��������ʽ��aaa|bbb|(�����ۺ�������Ч)
Const C_UserIcon_Width	="48"
Const C_UserIcon_Height	="48"

'Ԥ������
Const C_Vote_Action1	="�ʻ�"
Const C_Vote_Action2	="��ש"

'�༭��·��
Dim C_Editor,C_Editor_Type,C_Editor_LoadIcon,C_Editor_UBB

' 1 Ϊ4.x�༭�� 2Ϊ3.x�༭��
C_Editor_Type = 1

'���²���������ģ���
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