<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/Cls_XmlDoc.asp"-->
<%
dim action,rs,site_placard
Dim ADPath
Action=Trim(Request("Action"))
Dim xmlDoc
Set xmlDoc = New Cls_XmlDoc
xmlDoc.Unicode = False
ADPath = blogdir & oblog.CacheConfig(80)
If Not xmlDoc.LoadXml (ADPath& "/GG.config") Then
	If xmlDoc.LoadXml (blogdir& "XmlData/GG.config") Then
		xmlDoc.SaveAs (ADPath& "/GG.config")
	Else
		Response.Write (blogdir&"XmlData/GG.config �����ڣ��޷�����������")
		Set XmlDoc = Nothing
		Response.End
	End If
End If
xmlDoc.LoadXml (ADPath& "/GG.config")
if action="saveconfig" Then
	If Request.QueryString <>"" Then Response.End
	dim fso,dirstr,gg_teamtop,gg_teamcomment,gg_teambot,gg_teamlinks
	gg_teamtop=Trim(Request("gg_teamtop"))
	gg_teamcomment=Trim(Request("gg_teamcomment"))
	gg_teambot=Trim(Request("gg_teambot"))
	gg_teamlinks=Trim(Request("gg_teamlinks"))
	set fso=Server.CreateObject(oblog.CacheCompont(1))
	dirstr=Server.MapPath(ADPath)
	if fso.FolderExists(dirstr)=false then fso.CreateFolder(dirstr)
	Call oblog.BuildFile(dirstr&"\gg_teamcomment.htm",gg_teamcomment)
	Call oblog.BuildFile(dirstr&"\gg_teamtop.htm",gg_teamtop)
	Call oblog.BuildFile(dirstr&"\gg_teambot.htm",gg_teambot)
	Call oblog.BuildFile(dirstr&"\gg_teamlinks.htm",gg_teamlinks)
	xmlDoc.UpdateNodeText "gg_teamtop",oblog.htm2js_Script(gg_teamtop,"gg_teamtop"),True
	xmlDoc.UpdateNodeText "gg_teambot",oblog.htm2js_Script(gg_teambot,"gg_teambot"),True
	xmlDoc.UpdateNodeText "gg_teamlinks",oblog.htm2js_Script(gg_teamlinks,"gg_teamlinks"),True
	xmlDoc.UpdateNodeText "gg_teamcomment",oblog.htm2js_Script(gg_teamcomment,"gg_teamcomment"),True
	'���ݾɹ��
	Call oblog.BuildFile(dirstr&"\ad_teamcomment.htm",gg_teamcomment)
	Call oblog.BuildFile(dirstr&"\ad_teamtop.htm",gg_teamtop)
	Call oblog.BuildFile(dirstr&"\ad_teambot.htm",gg_teambot)
	Call oblog.BuildFile(dirstr&"\ad_teamlinks.htm",gg_teamlinks)
	xmlDoc.UpdateNodeText "ad_teamtop",oblog.htm2js_Script(gg_teamtop,"ad_teamtop"),True
	xmlDoc.UpdateNodeText "ad_teambot",oblog.htm2js_Script(gg_teambot,"ad_teambot"),True
	xmlDoc.UpdateNodeText "ad_teamlinks",oblog.htm2js_Script(gg_teamlinks,"ad_teamlinks"),True
	xmlDoc.UpdateNodeText "ad_teamcomment",oblog.htm2js_Script(gg_teamcomment,"ad_teamcomment"),True
	xmlDoc.Save
	set fso=Nothing
	Set XmlDoc = Nothing
	EventLog "�����޸�Ⱥ��ҳ����Ĳ���!",""
	oblog.ShowMsg "�������",""
else

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Ⱥ��ҳ��������</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸�Ⱥ��ҳ����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form name="form1" method="post" action="admin_teamad.asp">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
          <tr>
            <td>
            </td>
          </tr>
    <tr>
      <td height="22" class="topbg"><strong>Ⱥ��ҳ�涥����棨�˶δ�����ʾ������blog�û�ҳ�涥����</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_teamtop" cols="80" rows="8" ><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_teamtop.htm")%></textarea>
          </td></tr>
  </table>
  <br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>Ⱥ��ҳ��ײ���棨�˶δ�����ʾ������Ⱥ��ҳ����ײ���</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_teambot" cols="80" rows="8" ><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_teambot.htm")%></textarea>
          </td></tr>
  </table>
  <br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>���ӻظ��ϲ���棨�˶δ�����ʾ������Ⱥ���û����۴����ϲ���</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_teamcomment" cols="80" rows="8"><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_teamcomment.htm")%></textarea>
          </td></tr>
  </table>
  <br />
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr>
      <td height="22" class="topbg"><strong>�û��������Ӳ��ֹ�棨�˶δ�����ʾ��Ⱥ���������ӱ�ǩ���֣�</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">
                <textarea name="gg_teamlinks" cols="80" rows="8"><%=oblog.readfile("../"&oblog. CacheConfig(80)&"","gg_teamlinks.htm")%></textarea>
</td>
    </tr>
    <tr>
      <td height="40" align="center">
	  <input name="Action" type="hidden" id="Action" value="saveconfig">
                <input type="submit" name="Submit" value="�ύ�޸�">
      </td>
    </tr>
  </table>
      </form>
</body>
</html>
<%
end If
Set oblog=Nothing
%>