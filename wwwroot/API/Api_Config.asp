<%
'*********************************************************
'File:			Api_Config.asp
'Description:	DPO_API Config File For oBlog4.0
'Author:		感觉
'HomePage:		http://www.oblog.cn
'BBS			http://bbs.oblog.cn
'Copyright (C)	2004-2005 oblog.cn All rights reserved.
'LastUpdate:	20060815

'说明：
'打开conn.asp在<!--#include file="config.asp"-->的后面加上：
'<!--#include file="API/Api_Config.asp"-->
'API_Enable：通行证接口开关，True为启用接口，False为不启用
'oblog_Key ：oblog端网站安全码，所有系统的此量必须一致
'strTargetUrls：整合系统接口数据文件的url，如果整合多个，请用"|"符号分开

'*********************************************************

'整合通用接口参数
Const API_Enable = false 	'是否整合,如果整合请设为True,否则为False。
Const oblog_Key = "YIUSDYIUOIDSOIDUI"	'网站key，必须与整合端的key一致。
Const strTargetUrls = "http://bbs.4sk.cn/dv_dpo.asp"      '要整合的程序的完整URL（以“http://”开头，以接口文件的文件名结尾），如果有多个系统要整合，每个URL间用“|”分隔
Dim aUrls
aUrls=Split(strTargetUrls,"|")
%>