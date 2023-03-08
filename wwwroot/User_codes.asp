<!--#include file="user_top.asp"-->
<script language="javascript">
	function CopyIt(_sTxt){
		try{
			clipboardData.setData('Text',_sTxt);
			alert('邀请码 ' + _sTxt +' 已复制到剪贴板!');
			}
		catch(e)
		{}
		}
</script>
<%
If oblog.CacheConfig(17)=0 Then
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- 没有相关记录 -->
					<div class="msg">系统未启用邀请机制</div>
					<!-- 没有相关记录 end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
Response.End
End If
If oblog.l_Group(8,0)=0 Then
%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- 没有相关记录 -->
					<div class="msg">您所在的分组还不允许邀请他人加入</div>
					<!-- 没有相关记录 end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
Response.End
End If

Dim rs,Sql
Set rs=Server.CreateObject("Adodb.Recordset")
G_P_FileName="user_codes.asp?page="
Call List
Set rs=Nothing
%>
</body>
</html>
<%

Sub List()
Dim bTrue,i,lPage,lAll,lPages,iPage
bTrue=false
	'规则,每天可用的GUID数目
	'首先判断当天是否已经生成过了
	Sql="select lastCode From oblog_user Where userid=" & oblog.l_uid
	Set rs=oblog.Execute(Sql)
	If rs(0)="" Or IsNull(rs(0)) Then
		bTrue=true
	Else
		If datediff("d",rs(0),date)=0 Then
			bTrue=false
		Else
			bTrue=true
		End If
	End If
	rs.Close
	If bTrue Then
		For i=1 To oblog.l_Group(8,0)
			Sql="Insert into oblog_obcodes(obcode,creatuser,createtime,creatip,itype,istate) Values('"
			Sql=Sql & GetGUID & "'," & oblog.l_uid&",'" & Date & "','" & oblog.userip & "',0,0)"
			oblog.execute Sql
		Next
		oblog.Execute "Update oblog_user Set lastcode='" & Date &"' Where userid=" & oblog.l_uid
	End If
	rs.Open "select obcode,createtime From oblog_obcodes  Where creatuser=" & oblog.l_uid & " And iState=0 And iType=0 Order By createtime",conn,1,3
	lAll=Int(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- 没有相关记录 -->
					<div class="msg">您目前无可用邀请码</div>
					<!-- 没有相关记录 end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
    Exit Sub
    End If
    i=0
    iPage=10
	'分页
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'设置缓存大小 = 每页需显示的记录数目
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
	i=0
	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="javascript:void(null);" >您每天可获得 <%=oblog.l_Group(8,0)%> 个可用邀请码</a></li>
					<li><a href="javascript:void(null);" >系统每天自动回收 15 天前创建且未被使用的邀请码</a></li>

					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="CodesTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">可用邀请码</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form name="form1">
					<table id="Codes" class="TableList" cellpadding="0">
<%
do while not rs.eof
i = i + 1
%>
						<tr>
							<td class="t1" title="点击选中">
								<%=Right("00"&i,3)%>
							</td>
							<td class="t2">
								<input type="text" value="<%=rs(0)%>" size=50% name="code<%=i%>"  id="code<%=i%>" onclick="this.focus();this.select();" /><input type="submit" id="copy" onclick="copyclip('<%=rs(0)%>')" value="复制">
							</td>
						</tr>
<%
If i >= iPage Then Exit Do
rs.Movenext
Loop
%>
					</table>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/90.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
<%
rs.Close
Set rs=Nothing
End Sub
%>