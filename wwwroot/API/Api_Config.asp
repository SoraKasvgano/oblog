<%
'*********************************************************
'File:			Api_Config.asp
'Description:	DPO_API Config File For oBlog4.0
'Author:		�о�
'HomePage:		http://www.oblog.cn
'BBS			http://bbs.oblog.cn
'Copyright (C)	2004-2005 oblog.cn All rights reserved.
'LastUpdate:	20060815

'˵����
'��conn.asp��<!--#include file="config.asp"-->�ĺ�����ϣ�
'<!--#include file="API/Api_Config.asp"-->
'API_Enable��ͨ��֤�ӿڿ��أ�TrueΪ���ýӿڣ�FalseΪ������
'oblog_Key ��oblog����վ��ȫ�룬����ϵͳ�Ĵ�������һ��
'strTargetUrls������ϵͳ�ӿ������ļ���url��������϶��������"|"���ŷֿ�

'*********************************************************

'����ͨ�ýӿڲ���
Const API_Enable = false 	'�Ƿ�����,�����������ΪTrue,����ΪFalse��
Const oblog_Key = "YIUSDYIUOIDSOIDUI"	'��վkey�����������϶˵�keyһ�¡�
Const strTargetUrls = "http://bbs.4sk.cn/dv_dpo.asp"      'Ҫ���ϵĳ��������URL���ԡ�http://����ͷ���Խӿ��ļ����ļ�����β��������ж��ϵͳҪ���ϣ�ÿ��URL���á�|���ָ�
Dim aUrls
aUrls=Split(strTargetUrls,"|")
%>