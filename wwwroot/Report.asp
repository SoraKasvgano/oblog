<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dim logid,report_type
logid = CLng(Request("logid"))
report_type = Request("report_type")
'OB_DEBUG report_type,1
If report_type <>"" Then report_type = CLng(report_type) Else report_type = 0
If LogID = 0 Or UBound(oblog.CacheReport) = 0  Then
	Response.Write "缺少参数"
	Response.End
End if
If Trim(Request("action"))="save" Then Call Save()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>反映问题</title>
<script src="inc/main.js" type="text/javascript"></script>
<style type="text/css" media="screen">
<!--
/* ============================================================================ 全局 == */
html { height: 100%; }
body { margin: 0; height: 100%; background: #fff; font-size: 12px; font-family: tahoma,Arial,Century Gothic,verdana,Helvetica,sans-serif; color: #222; line-height: 150%; text-align: center; word-break: break-all; overflow:hidden; }

div,ul,ol,form { margin: 0; padding: 0; }
img { border: 0; }
li { list-style: none; }
table { font-size: 12px; }
input { font-family: tahoma,Arial,Helvetica,sans-serif; font-size: 12px; }

.red { color: #f30; }
/* ============================================================================ 布局 == */
#TableBody { width: 100%; height: 100%; border-collapse:collapse; border: none; }
		#TableBody thead #thead_tr1 { height: 32px; background: #1C7ABD; }
		#TableBody thead #thead_tr1 td { padding: 0 10px; border-bottom: 1px #000 solid; }
			#TableBody thead #thead_tr1 ul { position: absolute; top: 6px; left: 10px; float: left; }
				#TableBody thead #thead_tr1 ul li { float: left; margin: 0 10px 0 0; padding: 6px 10px 3px 10px; font-size: 14px; color: #fff;  }
				#TableBody thead #thead_tr1 ul li.Selected { background: #fff; border: 1px #000 solid; border-bottom: 1px #fff solid; font-size: 14px; font-weight: 600; color: #000; }

			#TableBody thead #thead_tr1 div { display: none; float: right; }
				#TableBody thead #thead_tr1 div a { display: block; width: 16px; height: 16px; overflow:hidden; background: url(Images/dialog/dialogCloseF.gif) no-repeat left top; text-indent: -9999px; }
				#TableBody thead #thead_tr1 div a:hover { background: url(Images/dialog/dialogClose0.gif) no-repeat left top; }
		#TableBody thead #thead_tr2 { height: 80px; }
			#TableBody thead #thead_tr2 td {  }
				#TableBody thead #thead_tr2 td div { padding: 10px; line-height: 1.8; }
					#TableBody thead #thead_tr2 td div img { float: left; margin: 0 10px 0 0;  }
	#TableBody tbody {  }
			#TableBody tbody tr td { vertical-align: top; }
				#TableBody tbody tr td div { height: 220px!important; height: 90%; overflow-x: hidden; overflow-y: auto;margin: 10px; padding: 10px; border-top: 1px #ddd solid; line-height: 2.5;  }
				*+html #TableBody tbody tr td div { height: 90%!important; }/* IE7 hack */
					#TableBody tbody tr td div strong { display: block; }
							#TableBody tbody tr td div span textarea { background: #fff; }
	#TableBody tfoot { text-align: right; height: 40px; background: #1C7ABD; }
			#TableBody tfoot tr td { padding: 0 10px 0 0; border-top: 1px #000 solid; }
				#TableBody tfoot tr td input { height: 20px; background: #eee;  }

-->
</style>
</head>
<body>

<form name="reportform" method="post" action="report.asp" target="_self">
<table id="TableBody" cellpadding="0">
	<thead>
		<tr id="thead_tr1">
			<td>
				<ul>
					<li class="Selected">反映问题</li>
				</ul>
				<div><a href="#" onclick="self.close();">关闭</a></div>
			</td>
		</tr>
		<tr id="thead_tr2">
			<td>
				<div><img src="Images/dialog/Alert.png" />如果您在<span class="red"><%=oblog.CacheConfig(2)%></span>中发现任何不适当、重复、侵权、垃圾的内容，请提交给我们，我们会及时处理。</div>
			</td>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div>
					<strong>问题类型：</strong>
<%=GetReport%>
				</div>
			</td>
		</tr>
	</tbody>
	<tfoot>
		<tr>
			<td>
				<input type="hidden" id="logid" name="logid" value="<%=logid%>" />
				<input type="hidden" id="action" name="action" value="save" />
				<input type="button" name="submit" onclick="report(document.reportform.logid.value,read_radio('report_type'));" value="向管理员提交问题" >
			</td>
		</tr>
	</tfoot>
</table>
</form>

</body>
</html>
<%
Function GetReport()
	Dim Reports,ii,Report
	Report = oblog.CacheReport
	For ii = 0 To UBound(Report)
		Reports = Reports &"					<label for="""&ii&"""><input type=""radio"" id="""&ii&""" name=""report_type"" id =""report_type"" value="""&ii&""" />"&Report(ii)&"</label><br />"
	Next
	GetReport = Reports
End Function
Sub Save()
	Dim RS,userid,authorid,ajax,username
	set ajax=new AjaxXml
	If oblog.checkuserlogined() Then
		userid = OBLOG.L_uid
		username = oblog.l_uname
	Else
		userid = 0
		username = "游客"
	End if
	Set RS = oblog.Execute ("SELECT authorid FROM oblog_log WHERE logid = "&logid)
	If RS.EOF Then
		ajax.re(Split("日志不存在$$$1$$$"&logid,"$$$"))
		Response.End
	Else
		authorid = RS(0)
		Set RS = Nothing
		if not IsObject(conn) then link_database
		Set RS = Server.CreateObject("ADODB.RecordSet")
		RS.open "SELECT * FROM oblog_digg WHERE 1 = 0",CONN,1,3
		RS.AddNew
		rs("userid") = userid
		rs("diggid") = 0
		rs("addip") = oblog.UserIp
		rs("logid") = logid
		rs("authorid") = authorid
		rs("diggtype") = report_type
		rs("username") = username
		If userid = 0 Then rs("isguest") = 1
		rs.Update
	End If
	ajax.re(Split("感谢您的反馈，我们会及时处理！$$$2$$$"&logid,"$$$"))
	Response.End
End Sub
%>
<script>
function report(logid,report_type){
	if (typeof(report_type) == 'undefined')
	{
		alert('您必须要选择一个项目！');
		return false;
	}
	var Ajax = new oAjax("report.asp?action=save",show_returnsave);
	var arrKey = new Array("logid","report_type");
	var arrValue = new Array(logid,report_type);
	Ajax.Post(arrKey,arrValue);
}
function show_returnsave(arrobj){
	if (arrobj){
		switch (arrobj[1]){
		case '1':
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.set('src',arrobj[1]);
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"");
			break;
		case '2':
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.set('src',arrobj[1]);
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"self.close()");
			break;
		}
		}
	}
</script>