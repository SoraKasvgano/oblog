<!--#include file="inc/inc_sys.asp"-->
<%
'20070704:������Ŀ87��(85Ϊ��֤ģ���ж�,86Ϊ�Ƿ���Ӿ���ͨ���������뷢�Ͷ���֪ͨ.87Ϊ�Զ����������վ����)
Const C_Items=90
Dim Action
Action = Trim(Request("action"))
Select Case action
	Case "saveconfig" 
		Call Saveconfig
	Case "updateuserdomain"
		Call updateuserdomain()
	Case Else 
		Call Showconfig
End Select 

Sub updateuserdomain()
					Dim user_domainroot,Arr_domainroot,TEMP_domainroot
					TEMP_domainroot=Trim(oblog.CacheConfig(4))
					If InStr(TEMP_domainroot,"|")>0 Then
						Arr_domainroot=Split(TEMP_domainroot,"|")
						user_domainroot=Arr_domainroot(0)
					Else
						user_domainroot=TEMP_domainroot
					End If
oblog.execute("update oblog_user set user_domain=userid,user_domainroot='"&user_domainroot&"' where user_domain='' or user_domain is null")
oblog.ShowMsg "���³ɹ�","close"
End Sub 

Sub Showconfig()
Dim rs,ac,sConfig,i
Set rs = oblog.execute("select ob_Value From oblog_config Where Id=1")
sConfig=rs(0)
ac=Split(sConfig,"$$")

'������������,�����Ҫ���ĳ���C_Items
'��Ŵ�1��ʼ
If UBound(ac)<C_Items Then
	For i=1 To (C_Items-UBound(ac))
		sConfig=sConfig & "$$0"
	Next
	'���·ָ�
	ac=Split(sConfig,"$$")
End If
Set oblog=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>վ������</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">��վ����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr>
      <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="#SiteInfo">��վ��Ϣ����</a> | <a href="#sysInfo">ϵͳ����</a> |   <a href="#SiteOption">����ģ��</a> | <a href="#sys">ϵͳ����ģ��</a> |  <a href="#spam">��������ģ��</a> | <a href="#code">��֤ģ��</a>  |  <a href="#biz">��ҵ�û�����ģ��</a> | <a href="#reg">ע��ѡ��</a>  <br />| <a href="#log">��־ѡ��</a> | <a href="#cmt">��������</a> | <a href="#group">Ȧ��ѡ��</a> </td>
    </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<br />
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">��վ����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="admin_setup.asp" id="form1" name="form1" onsubmit="return CheckRadio();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr >
      <td height="22" class="topbg" ><a name="SiteInfo"></a><strong>��վ��Ϣ����</strong></a></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��վ����<b>(������,��֧��Html)</b>��</td>
      <td  width="409" height="25"><% Call EchoInput("a1",40,50,ac(1))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��վ����<b>(������,��֧��Html)</b>��</td>
      <td><% Call EchoInput("a2",40,50,ac(2))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��վ��ַ��<br>
        ��Ҫ������д����URL��ַ,��http://www.oblog.com.cn/,<font color="#FF0000">����ʡ������/��</font>,�����ý�Ӱ�쵽rss��trackback���������С�</td>
      <td><% Call EchoInput("a3",40,255,ac(3))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > ������������<br />�밴��oblog.cn��������ʽ��д�����ж����������������&quot;|&quot;������<font color="#FF0000">��رն���������������</font>��</td>
      <td><% Call EchoInput("a4",40,255,ac(4))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > Ⱥ�������������<br />�밴��qq.oblog.cn��������ʽ��д�����ж����������������&quot;|&quot;���������ܺͶ����������ظ�.<font color="#FF0000">��رն���������������</font>��</td>
      <td><% Call EchoInput("a75",40,255,ac(75))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">�Ƿ������������û����ӣ�<br /><font color="#FF0000">��رն�����������ѡ���</font></td>
      <td><% Call EchoRadio("a5","","",ac(5))%>&nbsp;<font color="#FF0000">(��رջ�֧�ֶ�����������ѡ���!�������ǰ��δ���ù������������<A HREF="admin_setup.asp?action=updateuserdomain" target="_blank">��ʼ���û�����ѡ��</A>��)</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>վ��ؼ��֣�<br />�������ױ����������ҵ�,&quot;,&quot;�Ÿ�����</td>
      <td><% Call EchoInput("a9",50,100,ac(9))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td height="25">վ���Ȩ��Ϣ��<br />����ʾ��ϵͳҳ��ײ�����</td>
      <td><textarea name="a10" id='a10' cols="55" rows="5"><%=ac(10)%></textarea>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >վ�����䣺</td>
      <td> <% Call EchoInput("a11",50,100,ac(11))%></td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="sysInfo" id="user"></a><strong>ϵͳ����<font color="red">������Ƶ�����ģ�</font></strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >վ������·��������<b>��Ĭ�����·����</b></td>
      <td><% Call EchoRadio("a55","����·��","���·��",ac(55))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�ϴ�Ŀ¼��<b>��Ĭ�ϣ�UploadFiles��</b></td>
      <td><% Call EchoInput("a56",12,12,ob_iif(ac(56),"UploadFiles"))%><font color=red>����ָ������Ŀ¼�����ֹ�����Ŀ¼,���Ƽ����û�Ŀ¼���ϴ�Ŀ¼�����ڰ�ȫ���ã�</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >���Ŀ¼��<b>��Ĭ�ϣ�GG��</b></td>
      <td><% Call EchoInput("a80",12,12,ob_iif(ac(80),"GG"))%><font color=red>�����ֹ�����Ŀ¼����ȷ�϶Դ�Ŀ¼���޸ĵ�Ȩ�ޣ�</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��־����Ŀ¼��<b>��Ĭ��ARCHIVESĿ¼��</b></td>
      <td><% Call EchoRadio("a57","�û���Ŀ¼","ARCHIVESĿ¼",ac(57))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�û�Ŀ¼������ʽ��<b>��Ĭ���û�����Ŀ¼��</b></td>
      <td><% Call EchoRadio("a58","�û�ID","�û���",ac(58))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >����������ϵͳ����<b>��Ĭ�ϼ������ģ�</b>��</td>
      <td><% Call EchoRadio("a24","��������","����",ac(24))%><font color=red>�����������Ϊ�������Ĳ�Ҫѡ���������ή������Ч�ʣ�</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >����������ʱ��<b>��Ĭ��GMT+8.00��</b>��</td>
      <td>
	  <select name="a68" id="a68">
		<option value="">������������ʱ��</option>
		<option value="-12">(GMT-12.00)�������ڱ������</option>
		<option value="-11">(GMT-11.00)��;������Ħ��Ⱥ��</option>
		<option value="-10">(GMT-10.00)������</option>
		<option value="-9">(GMT-9.00)����˹��</option>
		<option value="-8">(GMT-8.00)̫ƽ��ʱ�䣨�����ͼ��ô󣩣��ٻ���</option>
		<option value="-7.a">(GMT-7.00)�����ߣ�����˹����������</option>
		<option value="-7.b">(GMT-7.00)ɽ��ʱ�䣨�����ͼ��ô�</option>
		<option value="-7.c">(GMT-7.00)����ɣ��</option>
		<option value="-6.a">(GMT-6.00)�ϴ���������ī����ǣ�������</option>
		<option value="-6.b">(GMT-6.00)��˹������</option>
		<option value="-6.c">(GMT-6.00)�в�ʱ�䣨�����ͼ��ô�</option>
		<option value="-6.d">(GMT-6.00)������</option>
		<option value="-5.a">(GMT-5.00)�������������</option>
		<option value="-5.b">(GMT-5.00)����ʱ�䣨�����ͼ��ô�</option>
		<option value="-5.c">(GMT-5.00)ӡ�ڰ����ݣ�������</option>
		<option value="-4.a">(GMT-4.00)������ʱ�䣨���ô�</option>
		<option value="-4.b">(GMT-4.00)������˹������˹</option>
		<option value="-4.c">(GMT-4.00)ʥ���Ǹ�</option>
		<option value="-3.a">(GMT-3.00)Ŧ����</option>
		<option value="-3.b">(GMT-3.00)��������</option>
		<option value="-3.c">(GMT-3.00)����ŵ˹����˹�����ζ�</option>
		<option value="-3.d">(GMT-3.00)������</option>
		<option value="-2">(GMT-2.00)�д�����</option>
		<option value="-1.a">(GMT-1.00)��ý�Ⱥ��</option>
		<option value="-1.b">(GMT-1.00)���ٶ�Ⱥ��</option>
		<option value="0">(GMT)�������α�׼ʱ�䣬�����֣����������׶أ���˹��</option>
		<option value="0.a">(GMT)����������������ά��</option>
		<option value="1.b">(GMT+1.00)��ķ˹�ص������֣������ᣬ����˹�¸��Ħ��άҲ��</option>
		<option value="1.c">(GMT+1.00)���������£�������˹������������˹��¬�������ǣ�������</option>
		<option value="1.d">(GMT+1.00)��³�������籾��������������</option>
		<option value="1.e">(GMT+1.00)�������ѣ�˹�������ɳ�������ղ�</option>
		<option value="1.f">(GMT+1.00)�з�����</option>
		<option value="2.a">(GMT+2.00)������˹��</option>
		<option value="2.b">(GMT+2.00)�����ף�����������</option>
		<option value="2.c">(GMT+2.00)�ն���������������ӣ������ǣ����֣�ά��Ŧ˹</option>
		<option value="2.d">(GMT+2.00)����</option>
		<option value="2.e">(GMT+2.00)�ŵ䣬��³�أ���˹̹��������˹��</option>
		<option value="2.f">(GMT+2.00)Ү·����</option>
		<option value="3.a">(GMT+3.00)�͸��</option>
		<option value="3.b">(GMT+3.00)�����أ����ŵ�</option>
		<option value="3.c">(GMT+3.00)Ī˹�ƣ�ʥ�˵ñ��������Ӹ���</option>
		<option value="3.d">(GMT+3.00)���ޱ�</option>
		<option value="3.e">(GMT+3.00)�º���</option>
		<option value="4.a">(GMT+4.00)�������ȣ���˹����</option>
		<option value="4.b">(GMT+4.00)�Ϳ⣬�ڱ���˹��������</option>
		<option value="4.5">(GMT+4.30)������</option>
		<option value="5.a">(GMT+5.00)Ҷ�����ձ�</option>
		<option value="5.b">(GMT+5.00)��˹�����������棬��ʲ��</option>
		<option value="5.5">(GMT+5.30)�����˹���Ӷ����������µ���</option>
		<option value="5.75">(GMT+5.45)�ӵ�����</option>
		<option value="6.a">(GMT+6.00)����ľͼ������������</option>
		<option value="6.b">(GMT+6.00)��˹���ɣ��￨</option>
		<option value="6.c">(GMT+6.00)˹����ǻ���������</option>
		<option value="6.d">(GMT+6.30)����</option>
		<option value="7.a">(GMT+7.00)����˹ŵ�Ƕ�˹��</option>
		<option value="7.b">(GMT+7.00)���ȣ����ڣ��żӴ�</option>
		<option value="8.a">(GMT+8.00)���������죬����ر�����������³ľ��</option>
		<option value="8.b">(GMT+8.00)��¡�£��¼���</option>
		<option value="8.c">(GMT+8.00)��˹</option>
		<option value="8.d">(GMT+8.00)̨��</option>
		<option value="8.e">(GMT+8.00)������Ŀˣ�������ͼ</option>
		<option value="9.a">(GMT+9.00)���࣬����������</option>
		<option value="9.b">(GMT+9.00)����</option>
		<option value="9.c">(GMT+9.00)�ſ�Ŀ�</option>
		<option value="9.501">(GMT+9.30)��������</option>
		<option value="9.502">(GMT+9.30)�����</option>
		<option value="10.a">(GMT+10.00)����˹��</option>
		<option value="10.b">(GMT+10.00)��������˹�пˣ������ˣ�</option>
		<option value="10.c">(GMT+10.00)�ص���Ī���ȱȸ�</option>
		<option value="10.d">(GMT+10.00)������</option>
		<option value="10.e">(GMT+10.00)��������ī������Ϥ��</option>
		<option value="11">(GMT+11.00)��ӵ���������Ⱥ�����¿��������</option>
		<option value="12.a">(GMT+12.00)�¿����������</option>
		<option value="12.b">(GMT+12.00)쳼ã�����Ӱ뵺�����ܶ�Ⱥ��</option>
		<option value="13">(GMT+13.00)Ŭ�Ⱒ�巨</option>
	</select>
	</td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="SiteOption" id="user"></a><strong>����ģ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>

    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ������û�ȡ�����룺</td>
      <td> <% Call EchoRadio("a84","","",ac(84))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ�������</td>
      <td> <% Call EchoRadio("a12","","",ac(12))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ������ֺУ�</td>
      <td> <% Call EchoRadio("a81","","",ac(81))%><font color="#FF0000">�������ȿ�������</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ�����᣺</td>
      <td> <% Call EchoRadio("a76","","",ac(76))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ��ֹ��������ͨ��<b>�����ý�ֹ��</b>��</td>
      <td> <% Call EchoRadio("a54","","",ac(54))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">�û�ҳ��ͳ�Ʒ�ˢ��ʱ�䣺</td>
      <td><% Call EchoInput("a31",10,10,Ob_IIF(ac(31),"30")) %>
        �� </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">��ҳ��̬�ļ��ĸ���ʱ�䣺</td>
      <td><% Call EchoInput("a33",10,10,Ob_IIF(ac(33),"300"))%>�� <font color="#FF0000">����������300�����ϣ����򽫼��ķѷ�������Դ��</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >ת���û���ҳ�Ƿ�ת��INDEX�ļ���</td>
      <td> <% Call EchoRadio("a46","","",ac(46))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�༭���Ƿ�����XHTML�鿴Դ�룺</td>
      <td> <% Call EchoRadio("a53","","",ac(53))%><font color="#FF0000">�����������ã������ַ�����������������������</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''" style="display:''">
      <td>�Զ���ȡ����Ϣ��ʱ�䣨Ĭ��10���ӣ���</td>
      <td><% Call EchoInput("a8",10,10,Ob_IIF(ac(8),"10"))%>��&nbsp;<font color="#FF0000">����ֵ��Ҫ̫С�����򼫺ķ���Դ��</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�Ƿ�����������</td>
      <td> <% Call EchoRadio("a67","","",ac(67))%>&nbsp;<font color="#FF0000">�������鿪�����ķ���Դ�ߣ�</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�Ƿ������ο����������</td>
      <td> <% Call EchoRadio("a82","","",ac(82))%>&nbsp;<font color="#FF0000">�����鲻����</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�Ƿ������ο�DIGG(�Ƽ���־)��</td>
      <td> <% Call EchoRadio("a83","","",ac(83))%></td>
    </tr>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�Ƿ��Զ����Ͷ���֪ͨ���Ӿ��Ƽ���ͨ��������ˣ�</td>
      <td> <% Call EchoRadio("a86","","",ac(86))%>(��ϵͳ��������Ӱ��)</td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="sys" id="user"></a><strong>ϵͳ����ģ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">ϵͳ��־�б�ÿҳ��ʾ��־������</td>
      <td><% Call EchoInput("a36",10,10,Ob_IIF(ac(36),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25"><p>ϵͳ��־�б������־��������</p></td>
      <td><% Call EchoInput("a37",10,10,Ob_IIF(ac(37),"500"))%>����Ӧlist.asp��</td>
	 </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�����б�ÿҳ��ʾ����������</td>
      <td><% Call EchoInput("a42",10,10,Ob_IIF(ac(42),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�����б���ʾ������������</td>
      <td><% Call EchoInput("a77",10,10,Ob_IIF(ac(77),"20"))%>����ӦListBlogger.asp��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">ϵͳ��Ƭ�б�ÿҳ��ʾ��Ƭ������</td>
      <td><% Call EchoInput("a38",10,10,Ob_IIF(ac(38),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25"><p>ϵͳ��Ƭ�б������Ƭ�ܸ�����</p></td>
      <td><% Call EchoInput("a39",10,10,Ob_IIF(ac(39),"500"))%>����Ӧphoto.asp��</td>
	</tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">ϵͳȺ���б�ÿҳ��ʾȺ�������</td>
      <td><% Call EchoInput("a78",10,10,Ob_IIF(ac(78),"20"))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25"><p>ϵͳȺ���б����Ⱥ���ܸ�����</p></td>
      <td><% Call EchoInput("a79",10,10,Ob_IIF(ac(79),"500"))%>����Ӧgroups.asp��</td>
	</tr>
    <td height="25" class="topbg"><a name="spam" id="user"></a><strong>��������ģ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>����ͨ����Ȩ����¼��ʱ�䣺</p></td>
      <td> <% Call EchoInput("a64",10,10,Ob_IIF(ac(64),"120"))%>���ӣ�30������С,1440�������</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>ͬһIP��λʱ�������������ͨ����Ŀ��<font color="#FF0000"><%=Chr(-23847)%></font></p></td>
      <td> <% Call EchoInput("a65",10,10,Ob_IIF(ac(65),"20"))%>�����������Զ�����IP��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>�������ͨ����Ŀ�ĵ�λʱ�䣺</p></td>
      <td> <% Call EchoInput("a66",10,10,Ob_IIF(ac(66),"120"))%>���ӣ����Ʊ��һ��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>ע����Ȩ����¼��ʱ�䣺</p></td>
      <td> <% Call EchoInput("a60",10,10,Ob_IIF(ac(60),"1440"))%>���ӣ�30������С,1440�������</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>ע���೤ʱ����Է�����־��</p></td>
      <td> <% Call EchoInput("a19",10,10,Ob_IIF(ac(19),"20"))%>����</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>ͬһIP����ע��֮��ļ��ʱ�䣺</p></td>
      <td><% Call EchoInput("a20",10,10,Ob_IIF(ac(20),"300"))%>��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > ͬһIP1Сʱ�ڵ�ע��������Ŀ��</td>
      <td> <% Call EchoInput("a21",10,10,Ob_IIF(ac(21),"20"))%>����0Ϊ�����ƣ��������ڵ�IP���⣩
        </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > ͬһIP24Сʱ�ڵ�ע��������Ŀ��</td>
      <td> <% Call EchoInput("a14",10,10,Ob_IIF(ac(14),"50"))%>����0Ϊ�����ƣ��������ڵ�IP���⣩
        </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>ͬһIP��λʱ������������ۡ�������Ŀ��<font color="#FF0000"><%=Chr(-23846)%></font></p></td>
      <td> <% Call EchoInput("a62",10,10,Ob_IIF(ac(62),"100"))%>��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>��λʱ������������ۡ���������Ŀ��<font color="#FF0000"><%=Chr(-23845)%></font></p></td>
      <td> <% Call EchoInput("a63",10,10,Ob_IIF(ac(63),"100"))%>��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>������ۡ�������Ŀ�ĵ�λʱ�䣺</p></td>
      <td> <% Call EchoInput("a61",10,10,Ob_IIF(ac(61),"60"))%>���ӣ����Ʊ�Ŷ�������</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >�������ٴ�������Ϊ��ϵͳ�Զ������</td>
      <td><% Call EchoInput("a13",10,10,Ob_IIF(ac(13),"5"))%>�Σ�0Ϊ�����ƣ�
    </tr>
	<tr>
      <td height="25" class="topbg"><a name="code" id="user"></a><strong>��֤ģ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
		   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ���ע���ʼ����</td>
      <td> <% Call EchoRadio("a88","","",Ob_IIF(ac(88),"0"))%>����Ҫ�ʼ����֧�֣����������ú���Ӧ���ʼ�������������</td>
    </tr>
			   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ�Ĭ���������û�����־��ʾ��ϵͳ��ҳ��</td>
      <td> <% Call EchoRadio("a89","","",Ob_IIF(ac(89),"0"))%>�����ѡ�ǣ��������ֶ�����û���ǰ̨������ʾ���Ρ���</td>
    </tr>
	  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td>��֤ģ�����������ã�</td>
      <td><input type="radio" name="a85" id="a85" value="0" <%If ac(85)=0 Then %>checked <%End If %> />ֻΪ������֤��<input type="radio" name="a85" id="a85" value="1"  <%If ac(85)=1 Then %>checked <%End If %>  />ֻ���Զ���������֤<input type="radio" name="a85" id="a85" value="2"  <%If ac(85)=2 Then %>checked <%End If %>  />�����֤��ʽ.<br/><font color="red">(ѡ���°�����ͻ����֤��ʽ�Ļ���¼ҳ��Ĭ��Ϊ��ʹ��������֤)</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td>�û�ע���Ƿ���Ҫ������֤ģ�飺</td>
      <td><% Call EchoRadio("a16","","",ac(16))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�û���¼�Ƿ���Ҫ������֤ģ�飺</td>
      <td><% Call EchoRadio("a29","","",ac(29))%>(��¼�ڿ�����֤����֤ģ��Ĭ��Ϊ��֤��)</td>
    </tr>
   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�û��������ۣ������Ƿ���Ҫ������֤ģ�飺</td>
      <td><% Call EchoRadio("a30","","",ac(30))%></td>
    </tr>
      <td height="25" class="topbg"><a name="biz" id="user"></a><strong>��ҵ�û�����ģ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
     <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" ><a href="http://news.oblog.cn/news/20060110192.shtml" target="_blank" title="�鿴����">�Ƿ������ƶ������</a></td>
      <td> <% Call EchoRadio("a51","","",ac(51))%>(ͨ�����ź��ʼ�������־��<a href=" http://www.oblog.cn/gmzn.shtml" target="_blank" title="�鿴����"><font color=red>��ҵ�汾����</font></a>)</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�����ƶ������,���ڽ������ݵĵ����ʼ���ַ(��������ѯOblog�ͷ���Ա)��</td>
       <td><% Call EchoInput("a52",30,50,ac(52))%>(�ռ��㹻���Ҳ�Ҫ����̫����˹��򣬷�ֹ���ղ�������)</td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="reg" id="user"></a><strong>ע��ѡ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�Ƿ��������û���</td>
      <td> <% Call EchoRadio("a15","","",ac(15))%><font color="#FF0000">���رպ󣬽�ֻ�����̨��ӣ�</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >�û�ע����Ƿ��Զ�����Ŀ¼��Ĭ���ǣ���</td>
      <td> <% Call EchoRadio("a59","","",ac(59))%><font color="#FF0000">��ѡ��ɽ�ʡ���̿ռ䣬��������ɲ��õ��û����飩</font></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25" >�Ƿ����������û���<font color="#FF0000"></font>��</td>
      <td><% Call EchoRadio("a6","","",ac(6))%>&nbsp;<font color="#FF0000">����������û�Ŀ¼ǿ��Ϊuserid��</font></td>
    </tr>
     <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>�Ƿ�������������ƣ�<a href="#h2" onClick="hookDiv('hh2','')"><img src="images/ico_help.gif" border=0></a></p></td>
      <td> <% Call EchoRadio("a17","","",ac(17))%></td>
    </tr>
    <tr id="hh2" style="display:none" name="h2">
      <td colspan=2> <p>ʲô��������</p>
        ���л�Աÿ����Ի�ȡһ�������������룬�����Խ����������ֹ����͸��������ڱ�վ��ע��<br/>
        (���ݻ�Ա��Ĳ�ͬ����������Ҳ��ͬ,ÿ��������ֻ��ʹ��һ�Σ��Ҳ����ۻ�)<br/>
        �»�Աע��ʱ����������һ����Ч����������ܽ���ע�ᣬ��������ע�ᡣ<br/>
        ʹ����������ƺ󣬽��鲻Ҫ������ע����ˣ��������Ϊע�Ჽ�跱�������û��������õ�����
        </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>���û�ע���Ƿ���Ҫ����Ա��֤��<br>
          </p></td>
      <td> <% Call EchoRadio("a18","","",ac(18))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>�ʼ���ַΨһ��</p></td>
      <td> <% Call EchoRadio("a22","","",ac(22))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>�ǳƲ������ظ���</p></td>
      <td> <% Call EchoRadio("a47","","",ac(47))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" > <p>�������Ʋ������ظ���</p></td>
      <td> <% Call EchoRadio("a48","","",ac(48))%></td>
    </tr>
      <td height="25" class="topbg"><a name="log"></a><strong>��־ѡ��</strong></td>
      <td class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''" style="display:''">
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">��ƪ��־�������������</td>
      <td><% Call EchoInput("a34",10,10,Ob_IIF(ac(34),"50000"))%>
        ��(Ӣ���ַ�) </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">��ƪ��־�������TAG����</td>
      <td><% Call EchoInput("a73",10,10,Ob_IIF(ac(73),"10"))%>
         </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">��ƪ��־�����������ͨ������</td>
      <td><% Call EchoInput("a74",10,10,Ob_IIF(ac(74),"20"))%>
         </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">������־��ϵͳ��־�����Ƿ�Ϊ���룺</td>
      <td><% Call EchoRadio("a25","","",ac(25))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">�Ƿ�������־ȫ������(����ر�)��</td>
      <td> <% Call EchoRadio("a26","","",ac(26))%></td>
    </tr>
      <td width="348" height="25">��־�Զ�����Ϊ�ݸ��ʱ�䣨Ĭ��2���ӣ���</td>
      <td><% Call EchoInput("a7",10,10,Ob_IIF(ac(7),"2"))%>��&nbsp;<font color="#FF0000">����ֵ��Ҫ̫С�����򼫺ķ���Դ��</font></td>
    </tr>
	 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >ϵͳ�Զ������������ǰ�Ļ���վ��־��<br>
        </td>
      <td> <% Call EchoInput("a87",10,10,Ob_IIF(ac(87),"100"))%>��  �������������������� �� 100 ϵͳĬ����СΪ 60 �죩</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">ȫվ����ÿ����<%=P_BLOG_UPDATEPAUSE%>ƪ��־��ͣ��ʱ�䣺</td>
      <td><% Call EchoInput("a28",10,10,Ob_IIF(ac(28),"5"))%>��&nbsp;��0Ϊ����ͣ�����Ϊ60��һ������Ϊ10���ɣ�</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">ҳ����������ʱ��ʾ�ַ���</td>
      <td><% Call EchoInput("a41",30,50,Ob_IIF(ac(41),"�����С�����"))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">ͼƬ�Զ���С��ȣ�Ϊ�㲻���ţ���</td>
      <td><% Call EchoInput("a43",10,10,Ob_IIF(ac(43),"0"))%>����</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >ͼƬ�Ƿ������������ţ�</td>
      <td> <% Call EchoRadio("a44","","",ac(44))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >������ʾ��־�Ƿ�ʹ��htm���ǿ�����ˣ�<br>
        </td>
      <td> <% Call EchoRadio("a45","","",ac(45))%>����ѡ�������г�ͼƬ����ı�Ƕ��������˵���</td>
    </tr>
	 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25" >��־�ļ���Ϊ��ʱĬ��Ϊ��<br></td>
      <td><% Call EchoRadio("a23","��־ID�Զ����","��־����ʱ��",ac(23))%></td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="cmt"></a><strong>����������</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>

    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ������οͷ������ۼ����ԣ�</td>
      <td> <% Call EchoRadio("a27","","",ac(27))%></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="348" height="25">�Ƿ�����������ÿ�������</td>
      <td> <% Call EchoRadio("a90","","",ac(90))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >������ظ��Ƿ�Ĭ��ͨ����ˣ�</td>
      <td> <% Call EchoRadio("a50","","",ac(50))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">��������ʽ��</td>
      <td><% Call EchoRadio("a40","����","����",ac(40))%>ֻ����ע���δ��������ʽ���û���Ч 
</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">���ۼ������������������</td>
      <td><% Call EchoInput("a35",10,10,Ob_IIF(ac(35),"2000"))%>
        ��(Ӣ���ַ�) </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">���ԣ����۵�ʱ������</td>
      <td><% Call EchoInput("a32",10,10,Ob_IIF(ac(32),"60"))%>�� </td>
    </tr>
    <tr>
      <td height="25" class="topbg"><a name="group"></a><strong>Ȧ��ѡ��</strong></td>
      <td height="22" class="topbg1"><a href="#top"><img src="images/ico_top.gif" border=0></a>&nbsp;<a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a>&nbsp;</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td height="25" >����Ȧ���Ƿ���Ҫ��ˣ�</td>
      <td> <% Call EchoRadio("a49","","",ac(49))%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">Ȧ������</td>
      <td><% Call EchoInput("a69",10,10,Ob_IIF(ac(69),"Ⱥ��"))%>&nbsp;<font color="#FF0000">����������Ϊ�����֣���Ȧ�ӣ�Ⱥ��ȣ�</font> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">Ȧ�ӹ���������</td>
      <td><% Call EchoInput("a70",10,10,Ob_IIF(ac(70),"Ⱥ��"))%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">Ȧ���������ޣ�</td>
      <td><% Call EchoInput("a71",10,10,Ob_IIF(ac(71),"200"))%> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td width="348" height="25">��־ͬʱ������Ȧ����Ŀ���ޣ�</td>
      <td><% Call EchoInput("a72",10,10,Ob_IIF(ac(72),"3"))%></td>
    </tr>
    <tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg"> <a name="formbottom"></a><input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " > </td>
    </tr>
  </table>
</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>
<script language="JavaScript">
//��λʱ����ѡ��
var TimeZoneObject
TimeZoneObject = document.getElementsByTagName('option');
for (var i = 0;i < TimeZoneObject.length ; i ++ ){
	if (TimeZoneObject[i].value =='<%=ac(68)%>'){
		TimeZoneObject[i].selected=true;
	}
}
</script>
<%
Set rs = Nothing
End Sub

Sub Saveconfig()
	If Request.QueryString <>"" Then Exit Sub
	Dim rs, i,sOpt
	Dim arrayList
	ReDim arrayList(C_Items)
  	For i=1 To C_Items
  		sOpt=sOpt & "$$" & Replace(Trim(Request.Form("a"&i)),"$","")
		arrayList(i) = Replace(Trim(Request.Form("a"&i)),"$","")
  	Next
	Dim arrayDir
	arrayDir = Oblog.SysDir
	For i = 0 To UBound(arrayDir)
		if LCase(arrayList(56)) = arrayDir(i) Then
			oblog.ShowMsg "����ѡ��ϵͳĿ¼��Ϊ�ϴ�Ŀ¼",""
		End If
	Next
	On Error Resume Next
	'�ж�Ŀ¼�Ƿ���ڣ�������������Զ�����
	Dim oFso
	Set oFso=Server.CreateObject(oblog.CacheCompont(1))
	If oFso.FolderExists(Server.Mappath(blogdir & LCase(arrayList(80)))) =False Then
		oFso.CreateFolder(Server.Mappath(blogdir & LCase(arrayList(80))))
	End If
	Set oFso=Nothing
	If Err Then
		Err.Clear
		oblog.ShowMsg "���Ŀ¼����ʧ�ܣ����ֹ�����",""
	End if
  	sOpt=Now&sOpt
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open  "select * From oblog_config Where Id=1",conn,1,3
    If rs.Eof Then rs.AddNew
    rs("ob_value")=sOpt
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ReloadCache
	EventLog "�����޸���վ��Ϣ���õĲ���!",""
    Set oblog=Nothing
    Response.Redirect "admin_setup.asp"
End Sub
%>
<script language="javascript">
function CheckRadio()
{
	var obj = document.getElementsByTagName("input");
	for (var i = 0;i<obj.length ;i++ ){
		var e = obj[i];
		if (e.type == 'radio'){
			if (e.value !=1 &&e.value!=0 &&e.value!=2){
				alert('��ȷ��ÿ�Ե�ѡ��ť����ѡ����һ��ѡ��');
				return false;
			}
		}
	}
}
</script>