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
	'重新分割
	ac=Split(sConfig,"$$")
End If

set rs=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>站点配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">网站配置</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="admin_score.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="tdbg">
      <td height="22" class="topbg"><a name="SiteInfo"></a><strong>网站积分配置</strong></td>
      <td height="22" class="topbg1"><a href="#formbottom"><img src="images/ico_bottom.gif" border=0></a></td>
    </tr>
     <tr>
     <td colspan=2>
     	①.积分设置请不要太高，建议尽量使用个位数，控制站内积分获取与消耗的平衡<br/>
     	①.上传文件与用户组绑定在一起，不再消耗或奖励积分
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >注册后的默认积分：</td>
      <td><% Call EchoInput("a1",20,5,ac(1))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >启用邀请机制后，一个有效邀请获取的积分奖励：</td>
      <td><% Call EchoInput("a2",20,5,ac(2))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >发布日志积分</td>
      <td><% Call EchoInput("a3",20,5,ac(3))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >日志被删除时的额外惩罚积分(日志删除时，将删除该日志已经获得的所有奖励积分)</td>
      <td><% Call EchoInput("a4",20,5,ac(4))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >留言积分(仅指被留言对象获取积分，如果该留言被删除，该积分将被扣除)</td>
      <td><% Call EchoInput("a5",20,5,ac(5))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >评论积分(仅指被评论对象获取积分，如果该评论被删除，该积分将被扣除)</td>
      <td><% Call EchoInput("a6",20,5,ac(6))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >支持与反对时的积分(用户进行评论时将从自己帐号中减去该积分值)，<br/>如果是支持，则目标用户将增加该积分值，如果是反对，则目标用户将减少该积分值</td>
      <td><% Call EchoInput("a20",20,5,ac(6))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >对日志表明态度(支持与反对)时的分值，仅注册用户可进行此操作，操作人需要消耗该分值，而目标对象将会获得(支持)或减少(反对)相应分值</td>
      <td><% Call EchoInput("a7",20,5,ac(7))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >推荐自己的文章为精华时所需要消耗的分值</td>
      <td><% Call EchoInput("a8",20,5,ac(8))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >通过精华审核后的奖励分值</td>
      <td><% Call EchoInput("a9",20,5,ac(9))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >推荐自己为推荐博客时所需要消耗的分值</td>
      <td><% Call EchoInput("a10",20,5,ac(10))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >创建一个群组时需要消耗的积分</td>
      <td><% Call EchoInput("a11",20,5,ac(11))%>
      </td>
    </tr>
   <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >群组被审核通过后所奖励的积分</td>
      <td><% Call EchoInput("a12",20,5,ac(12))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >日志发布到群组时的奖励积分</td>
      <td><% Call EchoInput("a13",20,5,ac(13))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >下载附件所需要的消耗的积分</td>
      <td><% Call EchoInput("a21",20,5,ac(21))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
      <td width="500" height="25" >日志被用户推荐(DIGG)一次，作者所增加的积分</td>
      <td><% Call EchoInput("a22",20,5,ac(22))%>
      </td>
    </tr>
 <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >发起一个辩论/活动申请时消耗的积分</td>
      <td><% Call EchoInput("a14",20,5,ac(14))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >辩论/活动通过审核后的奖励积分</td>
      <td><% Call EchoInput("a15",20,5,ac(15))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >参与辩论的积分</td>
      <td><% Call EchoInput("a16",20,5,ac(16))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >辩论/活动结束后的总结奖励</td>
      <td><% Call EchoInput("a17",20,5,ac(17))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >辩论/活动收益<br>
		辩论/活动过程中不进行计分操作，仅在辩论/活动结束后，辩论/活动发起者进行辩论/活动总结后再进行积分计算
		其获得的积分为：总结奖励+参与人数*收益,收益值建议为0.5~1.5
      	</td>
      <td><% Call EchoInput("a18",20,5,ac(18))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="display:none">
      <td width="500" height="25" >日志被归为站内专辑后的奖励积分</td>
      <td><% Call EchoInput("a19",20,5,ac(19))%>
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="40" colspan="2" align="center" class="tdbg"> <a name="formbottom"></a><input name="Action" type="hidden" id="Action" value="saveconfig">
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " > </td>
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
    			alert("<%=i%>所有项目必须填写!")
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
	EventLog "进行修改网站积分制度的操作",""
    Set oblog=Nothing
    Response.Redirect "admin_score.asp"
End Sub

%>