<!--#include file="inc/inc_sys.asp"-->
<%
Const C_Items=22
Dim Action
Action = Trim(Request("action"))
If Action = "saveconfig" Then
    Call Saveconfig
Else
    Call Showconfig
End If

Sub Showconfig()
dim rs,ac,sConfig,i
set rs=oblog.execute("select ob_Value From oblog_config Where id=3")
sConfig=rs(0)
ac=Split(sConfig,"$$")

If UBound(ac)<C_Items Then
	For i=1 To (C_Items-UBound(ac))
		sConfig=sConfig & "$$0"
	Next
	'���·ָ�
	ac=Split(sConfig,"$$")
End If

set rs=nothing
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
<form method="POST" action="admin_score.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg">
      <td height="22" class="topbg"><a name="SiteInfo"></a><strong>��վ��������</strong></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a></td>
    </tr>
     <tr>
     <td colspan=2>
     	��.���������벻Ҫ̫�ߣ����龡��ʹ�ø�λ��������վ�ڻ��ֻ�ȡ�����ĵ�ƽ��<br/>
     	��.�ϴ��ļ����û������һ�𣬲������Ļ�������
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >ע����Ĭ�ϻ��֣�</td>
      <td><% Call EchoInput("a1",20,5,ac(1))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >����������ƺ�һ����Ч�����ȡ�Ļ��ֽ�����</td>
      <td><% Call EchoInput("a2",20,5,ac(2))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >������־����</td>
      <td><% Call EchoInput("a3",20,5,ac(3))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >��־��ɾ��ʱ�Ķ���ͷ�����(��־ɾ��ʱ����ɾ������־�Ѿ���õ����н�������)</td>
      <td><% Call EchoInput("a4",20,5,ac(4))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >���Ի���(��ָ�����Զ����ȡ���֣���������Ա�ɾ�����û��ֽ����۳�)</td>
      <td><% Call EchoInput("a5",20,5,ac(5))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >���ۻ���(��ָ�����۶����ȡ���֣���������۱�ɾ�����û��ֽ����۳�)</td>
      <td><% Call EchoInput("a6",20,5,ac(6))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >֧���뷴��ʱ�Ļ���(�û���������ʱ�����Լ��ʺ��м�ȥ�û���ֵ)��<br/>�����֧�֣���Ŀ���û������Ӹû���ֵ������Ƿ��ԣ���Ŀ���û������ٸû���ֵ</td>
      <td><% Call EchoInput("a20",20,5,ac(6))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >����־����̬��(֧���뷴��)ʱ�ķ�ֵ����ע���û��ɽ��д˲�������������Ҫ���ĸ÷�ֵ����Ŀ����󽫻���(֧��)�����(����)��Ӧ��ֵ</td>
      <td><% Call EchoInput("a7",20,5,ac(7))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >�Ƽ��Լ�������Ϊ����ʱ����Ҫ���ĵķ�ֵ</td>
      <td><% Call EchoInput("a8",20,5,ac(8))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >ͨ��������˺�Ľ�����ֵ</td>
      <td><% Call EchoInput("a9",20,5,ac(9))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >�Ƽ��Լ�Ϊ�Ƽ�����ʱ����Ҫ���ĵķ�ֵ</td>
      <td><% Call EchoInput("a10",20,5,ac(10))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >����һ��Ⱥ��ʱ��Ҫ���ĵĻ���</td>
      <td><% Call EchoInput("a11",20,5,ac(11))%>
      </td>
    </tr>
   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >Ⱥ�鱻���ͨ�����������Ļ���</td>
      <td><% Call EchoInput("a12",20,5,ac(12))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >��־������Ⱥ��ʱ�Ľ�������</td>
      <td><% Call EchoInput("a13",20,5,ac(13))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >���ظ�������Ҫ�����ĵĻ���</td>
      <td><% Call EchoInput("a21",20,5,ac(21))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >��־���û��Ƽ�(DIGG)һ�Σ����������ӵĻ���</td>
      <td><% Call EchoInput("a22",20,5,ac(22))%>
      </td>
    </tr>
 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >����һ������/�����ʱ���ĵĻ���</td>
      <td><% Call EchoInput("a14",20,5,ac(14))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >����/�ͨ����˺�Ľ�������</td>
      <td><% Call EchoInput("a15",20,5,ac(15))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >������۵Ļ���</td>
      <td><% Call EchoInput("a16",20,5,ac(16))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >����/���������ܽά��</td>
      <td><% Call EchoInput("a17",20,5,ac(17))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >����/�����<br>
		����/������в����мƷֲ��������ڱ���/������󣬱���/������߽��б���/��ܽ���ٽ��л��ּ���
		���õĻ���Ϊ���ܽά��+��������*����,����ֵ����Ϊ0.5~1.5
      	</td>
      <td><% Call EchoInput("a18",20,5,ac(18))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >��־����Ϊվ��ר����Ľ�������</td>
      <td><% Call EchoInput("a19",20,5,ac(19))%>
      </td>
    </tr>
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
<%
Set rs = Nothing
End Sub

Sub Saveconfig()
	If Request.QueryString <>"" Then Exit Sub
	Dim rs, i,sOpt
    'Check
    For i=1 To C_Items
    	sOpt=Request.Form("a"&i)
    	If sOpt="" Or Not IsNumeric(sOpt) Then
    		%>
    		<script language="javascript">
    			alert("<%=i%>������Ŀ������д!")
    			history.back();
    		</script>
    		<%
    		Response.End
    	End If
  	Next
  	sOpt=""
  	For i=1 To C_Items
  		sOpt=sOpt & "$$" & Replace(Trim(Request.Form("a"&i)),"$","")
  	Next
  	sOpt=Now&sOpt
    If Not IsObject(conn) Then link_database
    Set rs = Server.CreateObject("adodb.recordset")
    rs.Open  "select * From oblog_config Where Id=3",conn,1,3
    If rs.Eof Then rs.AddNew
    rs("ob_value")=sOpt
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ReloadCache
	EventLog "�����޸���վ�����ƶȵĲ���",""
    Set oblog=Nothing
    Response.Redirect "admin_score.asp"
End Sub

%>