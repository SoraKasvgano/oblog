<!--#include file="user_top.asp"-->
<%
dim action,iPage,sGuide
Dim tableName,itype,tableName1,tSQL
Dim teamID
Action=trim(request("Action"))
teamID=trim(request("teamid"))
If teamID<>"" Then teamID = CLng(teamID)
If teamID <> "" Then
	tableName = "[oblog_teamskin]"
	tableName1 = "[oblog_team]"
	tSQL = "teamID="&teamID
	itype = 1
Else
	tableName = "[oblog_userskin]"
	tableName1 = "[oblog_user]"
	tSQL = "userid="&oblog.l_uid
	itype = 0
End if
iPage=12
%>
<%
'����ÿҳģ����ʾ��ʾ��Ŀ
If  teamID<>"" Then
	G_P_FileName="user_template.asp?teamid="&teamID&"&page="
Else
	G_P_FileName="user_template.asp?page="
End if
G_P_PerMax=12
'ģ����ʾ˳��,1:����,���µ�����ǰ��;2:˳��,���ϵ�����ǰ��
Const P_USER_TEMPLATE_ORDERBY=1
%>
<%

select case Action
	case "saveconfig"
		call saveconfig()
	case "modiconfig"
		call modiconfig()
	case "savedefault"
		call savedefault()
	case "modiviceconfig"
		call modiviceconfig()
	case "saveviceconfig"
		call saveviceconfig()
	case "bakskin"
		call bakskin()
	case "good"
		call good()
	case Else
		call showconfig()
End select
%>
</table>
</body>
</html>
<script language=javascript>
function VerIfySubmit()
{
	If (document.oblogform.edit.value == "")
     {
        alert("������ģ������!");
	return false;
     }
	return true;
}
function setbak()
{
	document.bakform.bak.value='bak';
}
function setrestore()
{
	document.bakform.bak.value='restore';
}

function skin_help(action,t){
	if (t == 0)
	{
		if (action==0){
			var str="<div style='height:200px;overflow:auto;z-index:999999;'>��ģ���ǩ<hr />$show_log$ ��Ҫ���˱����ʾ��־���岿�֣��������۵���Ϣ��<br />$show_placard$ �˱����ʾ�û����档 <br />$show_calendar$ �˱����ʾ������ <br />$show_newblog$ �˱����ʾ������־�б� <br />$show_comment$ �˱����ʾ���»ظ��б�<br />$show_subject$ �˱����ʾר����ࡣ <br />$show_subject_l$ �˱�Ǻ�����ʾר����ࡣ<br />$show_newblog$ �˱����ʾ������־�б�<br />$show_newmessage$ �˱����ʾ���������б�<br />$show_info$ �˱����ʾBlog���ƣ�ͳ����Ϣ�ȡ� <br />$show_login$ �˱����ʾ��¼���� <br />$show_links$ �˱����ʾ������Ϣ��<br />$show_blogname$ �˱����ʾ�û�blog���ƣ�������Ϊ������ʾ�û�id��<br />$show_search$ �˱����ʾ��������<br />$show_xml$ �˱����ʾrss���ӱ�־��<br />$show_blogurl$ �˱����ʾ�������ӡ�<br />$show_myfriend$ �˱�ǩ��ʾ�ҵĺ��ѡ�<br />$show_mygroups$ �˱�ǩ��ʾ�Ҽ����Ⱥ�顣<br />$show_photo$ �˱�ǩ������ᡣ</div>";
		}else{
			var str="<div style='height:150px;overflow:auto;z-index:999999;'>��ģ���ǩ<hr /> $show_topic$ �˱����ʾ��־��Ŀ�� <br />$show_loginfo$ �˱����ʾ��־���ߣ�����ʱ�����Ϣ�� <br />$show_logtext$ �˱����ʾ��־���ġ� <br />$show_more$ �˱����ʾ�Ķ�ȫ�ģ����õ����ӡ� <br />$show_emot$ �˱�ǽ���ʾ��ʾ����ͼ�ꡣ<br />$show_author$ �˱�ǽ���ʾ��������<br />$show_addtime$ �˱�ǽ���ʾ����ʱ�䡣<br />$show_topictxt$ �˱�ǽ���ʾ��־���⡣</div>";
		}
	}
	else
	{
		if (action==0){
			var str="<div style='height:200px;overflow:auto;z-index:999999;'><%=oblog.CacheConfig(69)%>��ģ���ǩ<hr />$group_id$ <%=oblog.CacheConfig(69)%>ID<br />$group_posts$ ��������<br /> $group_ico$  <%=oblog.CacheConfig(69)%>���ͼƬ <br /> $group_url$ <%=oblog.CacheConfig(69)%>���ʵ�ַ <br />$group_guide$ ��������<br /> $group_name$ <%=oblog.CacheConfig(69)%>���� <br /> $group_creater$ <%=oblog.CacheConfig(69)%>������ <br /> $group_bottom$ ��Ȩ��ʶ<br /> $group_comments$ �������  <br />$group_placard$ ����<br /> $group_links$ �������� <br /> $group_info$ <%=oblog.CacheConfig(69)%>��Ϣ<br /> $group_bestuser$ ��Ծ�û�<br /> $group_newuser$ ���¼����û�<br />  $group_admin$ ����Ա��Ϣ<br /> $group_bestposts$ �������� <br />$group_photo$ <%=oblog.CacheConfig(69)%>��Ƭ </div>";
		}else{
			var str="<div style='height:150px;overflow:auto;z-index:999999;'><%=oblog.CacheConfig(69)%>��ģ���ǩ<hr /> $group_list$ ���ݱ�ǩ <br /> $group_post_title$ ���ӱ��� <br />  $group_content$ ��������<br />  $group_post_userico$ ����ͷ��<br />  $group_post_user$ �������� <br />  $group_post_time$ ����ʱ�� <br />  $group_post_content$ �������� <br />  $group_post_id$ ����ID <br /> $group_post_replys$ �ظ���ť <br />  $group_post_userurl$ �������ߵ�ַ <br />   $group_post_high$ ����¥��  <br /> $group_post_m$���Ӳ������� <br /> </div>";
		}
	}
	 var oDialog = new dialog("<%=blogurl%>");
	 oDialog.init();
	 oDialog.event(str,'');
	 oDialog.button('dialogOk',"");
	 document.getElementById("ob_boxface").style.display="none";
}

</script>
<%
sub showconfig()
	Dim rs,SkinStrings,defaultskin,rst,sql,sClasses,sqlclass,classid
	Dim bookmark,sOrderby
	Set rs=Server.CreateObject("Adodb.Recordset")
	classid=Request("classid")
	If classid<>"" Then classid=Int(classid)
	If classid=0 Then classid=""
	'ȡĬ��ģ��
	rs.Open "select defaultskin from "&tableName1&" where "&tSQL,conn,1,3
	defaultskin=OB_IIF(rs(0),0)
	rs.Close
	If classid<>"" Then sqlclass=" And classid=" & classid:G_P_FileName="user_template.asp?teamid="&teamid&"&classid="&classid&"&page="
	'ȡ�û�/Ȧ��ģ����÷���
	If OB_IIF(teamid,"")="" Then
		If OB_IIF(oblog.l_Group(13,0),"")="" Then
			sql="select * from "&tableName&" where ispass=1 "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And icount>0")
		Else
			sql="select * from "&tableName&" where ispass=1 And classid in (" & oblog.l_Group(13,0) &") "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And classid in (" & oblog.l_Group(13,0) &")  And icount>0")
		End If
	Else
		If OB_IIF(oblog.l_Group(5,0),"")="" Then
			sql="select * from "&tableName&" where ispass=1 "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And icount>0")
		Else
			sql="select * from "&tableName&" where ispass=1 And classid in (" & oblog.l_Group(5,0) &") "  & sqlclass & " Order By Id"
			Set rst=oblog.Execute("Select * From oblog_skinclass Where itype="&itype&" And classid in (" & oblog.l_Group(5,0) &")  And icount>0")
		End If
	End If
	sClasses="<option value=""0"">�鿴ȫ��ģ��</option>"
	Do While Not rst.Eof
		sClasses=sClasses & "<option value=" & rst("classid") & ">" & rst("classname")& "(��" & rst("icount") &"��ģ��)</option>"
		rst.Movenext
	Loop
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
<!--					<fieldset id="Template" class="FieldsetForm">
						<legend>��ǰģ�壺</legend>
							<ul>
								<li>
								<%
								'If defaultskin=0 Then
									'Response.Write "									<strong class=""red"">�㵱ǰ��û��ѡ���κ�ģ��</strong>" & VBCRLF
								'Else
									'Set rst=oblog.Execute("Select * From "&tableName&" Where id=" & defaultskin)
									'If Not rst.Eof Then
										'Response.Write "<ul class=""Skin_onmouseover"">" & VBCRLF
										'Response.Write "										<li class=""l1""><a href=""showskin.asp?id=" & rst("id") & """ target=""_blank""><img src="""&oblog.filt_html(rst("skinpic"))&""" title=""���Ԥ��"" width=""200"" height=""122"" border=""0"" /></a></li>" & VBCRLF
										'Response.Write "										<li class=""l2"">���ƣ�" & "<strong>" & rst("userskinname") &"</strong></li>" & VBCRLF
										'Response.Write "										<li class=""l3"">���ߣ�<a href=""" & rst("skinauthorurl") &""" target=""_blank"">" & "<span class=""blue"">" & rst("skinauthor") &"</span></a></li>" & VBCRLF
										'Response.Write "									</ul>" & VBCRLF
									'End If
									'Set rst=Nothing
								'End If
								%>
								</li>
							</ul>
					</fieldset>-->
<%
Dim lPage,lAll,lPages
If P_USER_TEMPLATE_ORDERBY=1 Then sOrderby=" Desc"
rs.Open sql & sOrderby,conn,1,3
lAll=clng(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
				<div id="chk_idAll">
					<!-- û����ؼ�¼ -->
					<div class="msg"><%=sGuide & " û����ؼ�¼" %></div>
					<!-- û����ؼ�¼ end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
    	<%
    	Exit Sub
    End If

	'��ҳ
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = clng(Request("page"))
	End If

	'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
%>

					<fieldset id="Template" class="FieldsetForm">
						<legend>ѡ��ģ�壺<span class="red">��<%
								If defaultskin=0 Then
									Response.Write "<strong class=""red"">�㵱ǰ��û��ѡ���κ�ģ��</strong>" & VBCRLF
								Else
									Set rst=oblog.Execute("Select * From "&tableName&" Where id=" & defaultskin)
									If Not rst.Eof Then
										If teamid <>"" Then
											tSQL = "teamskinid=" & rst("id")
										Else
											tSQL = "id=" & rst("id")
										End if
										Response.Write "��ǰģ��Ϊ��<a href=""showskin.asp?" & tSQL & """ target=""_blank""><strong>" & rst("userskinname") &"</a></strong>" & VBCRLF
									End If
									Set rst=Nothing
								End If
								%>��</span></legend>
							<ul>
								<li id="form">
								<%If teamid = "" Then %>
									<form id="sClasses" name="formclass" method="post" action="user_template.asp">
										<select name="classid" onchange="javascript:window.location='user_template.asp?action=showconfig&classid='+this.options[this.selectedIndex].value;">
												<option>��ѡ��ģ�����</option>
												<%=sClasses%>
										</select>
									</form>
								<%End if%>
									<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
								</li>
								<li>
									<%
										SkinStrings = SkinStrings & GetSkinList(rs,defaultskin)
										rs.Close
										set rs=Nothing
										Response.Write SkinStrings
										SkinStrings=""
									%>
									<input type="hidden" value="<%=request("u")%>">
									<input type="hidden" name="teamid" value="<%=teamid%>">
								</li>
							</ul>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%
	set rs=nothing
End sub

sub savedefault()
	dim rs,rsskin,isdefaultID
	isdefaultID=clng(trim(request("radiobutton")))
	set rsskin=oblog.execute("select skinmain,skinshowlog from "&tableName&" where id="&isdefaultID)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select user_skin_main,user_skin_showlog,defaultskin from "&tableName1&" where "&tSQL,conn,1,3
	rs(0)=rsskin(0)
	rs(1)=rsskin(1)
	rs(2)=isdefaultID
	set rsskin=nothing
	rs.update
	rs.close
	set rs=nothing
	'���������ģʽ
	If teamid = "" Then
		If oblog.l_uNewbie=1 Then
			'�����µ��û�Ŀ¼
			If oblog.CacheConfig(59) = "0" Then oblog.CreateUserDir oblog.l_uname, 1
			oblog.execute "Update oblog_user Set newbie=0 Where userid=" & oblog.l_uid
			'���Session����ֹѡ����޷���ʱ��Ч
			Session ("CheckUserLogined_"&oblog.l_uName) = ""
			Oblog.CheckUserLogined()
			updateindex()
			oblog.ShowMsg "ģ��ѡ��ɹ���","user_template.asp?action=showconfig&teamid="&teamid
		Else
'			Session ("CheckUserLogined_"&oblog.l_uName) = ""
			updateindex()
			oblog.ShowMsg "�޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�","user_template.asp?action=showconfig&teamid="&teamid
		End If
	Else
		oblog.ShowMsg "�޸ĳɹ���","user_template.asp?action=showconfig&teamid="&teamid
	End if
End sub

sub saveconfig()
	dim rs,sql,sContent,iChk
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "�����ڵ��鲻�����޸�ģ�壡","user_template.asp"
		Exit Sub
	End If
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_main from "&tableName1&" where "&tSQL
	rs.open sql,conn,1,3
	sContent=oblog.filtpath(oblog.filt_badword(Trim(request("edit"))))
'	OB_DEBUG sContent,1
	'���ݼ��
	iChk=oblog.chk_badword(sContent)
	If iChk>0 Then
		oblog.ShowMsg "ģ���д���ϵͳ��ֹ���ַ�!",""
		Response.End
	End If
	sContent=Replace(Replace(sContent,"<%","&lt;%"),"%"&">","%&gt;")
	rs(0)=sContent
	'�ű�����
	If oblog.l_Group(15,0)=0 Then rs(0)=oblog.CheckScript(sContent)

	rs.update
	rs.close
	Set rs=Nothing
	sContent=""
	If teamID <> "" Then
		oblog.ShowMsg "�޸ĳɹ���",""
	Else
		updateindex()
		oblog.ShowMsg "�޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�",""
	End If
End sub

sub saveviceconfig()
	dim rs,sql,sContent
	Dim iChk
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "�����ڵ��鲻�����޸�ģ�壡",""
		Exit Sub
	End If
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_showlog from "&tableName1&" where "&tSQL
	rs.open sql,conn,1,3
	sContent=oblog.filtpath(oblog.filt_badword(trim(request("edit"))))
	'���ݼ��
	iChk=oblog.chk_badword(sContent)
	If iChk>0 Then
		oblog.ShowMsg "ģ���д���ϵͳ��ֹ���ַ�!",""
		Response.End
	End If
	'�ű�����
	If oblog.l_Group(15,0) = "0" Then rs(0)=oblog.CheckScript(sContent)
	rs(0)=sContent
	rs.update
	rs.close
	set rs=nothing
	sContent=""
	If teamID <> "" Then
		oblog.ShowMsg "�޸ĳɹ���",""
	Else
		updateindex()
		oblog.ShowMsg "�޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�",""
	End If
End sub

sub modiconfig()
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "�����ڵ��鲻�����޸�ģ�壡","user_template.asp"
		Exit Sub
	End If
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_main,defaultskin from "&tableName1&" where "&tSQL)
	If rs("defaultskin")=0  or rs("defaultskin")="" Or IsNull (rs(1)) Then
		set rs=nothing
		oblog.adderrstr("����ѡ��һ��Ĭ��ģ�壡")
		oblog.showusererr
		exit sub
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>�޸���ģ�壺</legend>
						<form method="POST" action="user_template.asp" id="oblogform" name="oblogform"  <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
							<ul>
								<li><strong>��ģ�����ҳ��������</strong>�������޸�ǰ��<a href="user_template.asp?action=bakskin&teamid=<%=teamid%>"><span class="blue">����ģ��</span></a>��<br /><a href="#" onclick="skin_help(0,<%=OB_IIF(teamid,0)%>);" ><span class="blue">��ģ���ǩ˵��</span></a></li>
								<li><span id="loadedit" class="red" style="display:<%=C_Editor_LoadIcon%>;"><img src="images/loading.gif" align="absbottom" /> ��������༭��...</span>
									<textarea id="edit" name="edit" style="	display:none"><%=Server.HtmlEncode(rs(0))%></textarea>
										<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
								</li>
								<input name="teamid" type="hidden" id="teamid" value="<%=teamid%>" />
								<li><input name="Action" type="hidden" id="Action" value="saveconfig" /><input name="cmdSave" type="submit" id="Submit" value=" �����޸� " style="height:30px;" /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",1,"580","300"%>
<%
set rs=nothing
End sub

sub modiviceconfig()
	If oblog.l_Group(14,0)=0 Then
		oblog.ShowMsg "�����ڵ��鲻�����޸�ģ�壡","user_template.asp"
		Exit Sub
	End If
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_showlog,defaultskin from "&tableName1&" where "&tSQL)
	If rs("defaultskin")=0  or rs("defaultskin")="" Or IsNull(rs(1)) Then
		set rs=nothing
		set rs=nothing
		oblog.adderrstr("����ѡ��һ��Ĭ��ģ�壡")
		oblog.showusererr
		exit sub
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll" style="overflow-y:auto;">
					<fieldset id="BackUp" class="FieldsetForm">
						<legend>�޸ĸ�ģ�壺</legend>
						<form method="POST" action="user_template.asp?id=<%=clng(request.QueryString("id"))%>" id="oblogform" name="oblogform" <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
							<ul>
								<li><strong>��ģ�������־������ʾ���</strong></a>��<br /><a href="#" onclick="skin_help(1,<%=OB_IIF(teamid,0)%>);"><span class="blue">��ģ���ǩ˵��</span></a></li>
								<li><span id="loadedit" class="red" style="display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> ��������༭��...</span>
									<textarea id="edit" name="edit" style="display:none;"><%=Server.HtmlEncode(rs(0))%></textarea>
									<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
									</li>
								<li><input name="action" type="hidden" id="action" value="saveviceconfig" />
								<input name="teamid" type="hidden" id="teamid" value="<%=teamid%>" /><input type="submit"  value="�����޸�" id="Submit" /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",1,"580","300"%>
<%
	set rs=nothing
End sub

sub bakskin()
	dim bak,rs
	bak=request("bak")
	If bak="bak" Then
		oblog.execute("update "&tableName1&" set bak_skin1=user_skin_main,bak_skin2=user_skin_showlog where "&tSQL)
		oblog.ShowMsg "����ģ��ɹ���",""
	ElseIf bak="restore" Then
		set rs=oblog.execute("select bak_skin1,bak_skin2 from "&tableName1&" where "&tSQL)
		If rs(0)="" or rs(1)="" or isnull(rs(0)) or isnull(rs(1)) Then
			set rs=nothing
			oblog.adderrstr("�����ݵ�ģ��Ϊ�գ�������ָ���")
			oblog.showusererr
		End If
		oblog.execute("update "&tableName1&" set user_skin_main=bak_skin1,user_skin_showlog=bak_skin2 where "&tSQL)
		set rs=nothing
		oblog.ShowMsg "�ָ�ģ��ɹ���",""
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUpTemplate" class="FieldsetForm">
						<legend>����ģ��</legend>
						<form name="bakform" method="post" action="user_template.asp?action=bakskin&teamid=<%=teamid%>">
							<ul>
								<li><input type="hidden" name="bak" id="bak" value="" /><input type="submit" id="Submit" name="Submit" value="����ģ��" onclick="setbak();" /></li>
								<li>��������ʹ���е�ģ�塣<a href="showskin.asp?<%=tSQL%>" target ="_blank"><span class="red">�鿴Ŀǰģ�屸��</span></a></li>
								<li></li>
							</ul>
					</fieldset>
					<fieldset id="BackUpTemplate" class="FieldsetForm">
						<legend>�ָ�ģ��</legend>
							<ul>
								<li><input type="submit" id="Submit" name="Submit2" value="�ָ�ģ��" onclick="setrestore();" /></li>
								<li>�ָ�������󱸷ݵ�ģ�壬����������ʹ���е�ģ�塣</li>
								<li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%End Sub

sub good()
	dim good,rs,rsu,skinname
	good=request("good")
	If good="save" Then
		skinname=request("skinname")
		If skinname="" or oblog.strLength(skinname)>50  Then oblog.adderrstr("�û�������Ϊ��(���ܴ���50)��")
		If oblog.errstr<>"" Then oblog.showusererr:exit sub
		set rsu=oblog.execute("select user_skin_main,user_skin_showlog from "&tableName1&" where "&tSQL)
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 1 * from ["&tableName&"]",conn,1,3
		rs.addnew
		rs("userskinname")=skinname
		rs("skinmain")=rsu(0)
		rs("skinshowlog")=rsu(1)
		rs("skinauthor")=oblog.l_uname
		rs("skinauthorurl")=oblog.l_udir&"/"&oblog.l_uid&"/index."&f_ext
		rs.update
		rs.close
		set rsu=nothing
		set rs=nothing
		oblog.ShowMsg "�Ƽ��ɹ�����ȴ�����Ա��ˣ�",""
	End If
%>
<table id="TableBody" class="template" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="BackUpTemplate" class="FieldsetForm">
						<legend>�Ƽ��ҵ�ģ��</legend>
						<form name="good" action="user_template.asp?action=good&good=save" method="post">
							<ul>
								<li>������ģ���Ư�����Ƽ�������Ա�����Է���ģ�����ݿ���ø�����ʹ��Ŷ!<br />ģ���ע�����ߺ����blog���ӡ��벻Ҫ�ύ�Ѿ����ڵ��û�ģ�壡</li>
								<li>ģ�����ƣ�<input name="skinname" type="text" value="" size="20" maxlength="20" /></li>
								<li><input type="submit" id="Submit" value="�Ƽ�" /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%End sub
sub updateindex()
	dim blog
	set blog=new class_blog
	blog.userid=oblog.l_uid
	blog.update_index 0
	blog.update_message 0
	blog.CreateFunctionPage
	set blog=nothing
End sub

Function GetSkinList(byref rst,byval defaultskin)
	Dim strSkins,ustr,iCount
	iCount=0
	strSkins=""
	Do While Not rst.eof
		If teamid <>"" Then
			tSQL = "teamskinid=" & rst("id")
		Else
			tSQL = "id=" & rst("id")
		End If
		If rst("skinauthorurl")<>"" Then
			ustr="<a href="""&rst("skinauthorurl")&""" target=""_blank""><span class=""blue"">"&rst("skinauthor")&"</span></a>"
		Else
			ustr=rst("skinauthor")
		End If
		strSkins = strSkins & "									<ul onmouseover=""this.className='Skin_onmouseover'"" onmouseout=""this.className='Skin_onmouseout'"" class=""Skin_onmouseout"">" & VBCRLF
		strSkins = strSkins & "										<li class=""l1""><a href='showskin.asp?"&tSQL&" ' target=_blank>" & VBCRLF
		If rst("skinpic")="" or isnull(rst("skinpic")) Then
			strSkins = strSkins & "<img src=""images/nopic.gIf"" title=""�Բ���,��ģ��û��Ԥ��ͼ"" width=""200"" height=""122"" border=""0"" />" & VBCRLF
		Else
			strSkins = strSkins & "<img src="""&oblog.filt_html(rst("skinpic"))&""" title=""���Ԥ��"" width=""200"" height=""122"" border=""0"" />" & VBCRLF
		End If
		strSkins = strSkins & "</a></li>" & VBCRLF
		strSkins = strSkins & "										<li class=""l2"" title=""ģ�壺" & rst("userskinname") & """>ģ�壺<strong>" & rst("userskinname") & "</strong></li>" & VBCRLF
		strSkins = strSkins & "										<li class=""l3"" title=""���ߣ�"&rst("skinauthor")&""">���ߣ�" & ustr & "</li>" & VBCRLF
'		strSkins = strSkins & "<div class=""skin_used""><a href=""user_template.asp?action=savedefault&teamid="&teamid&"&radiobutton=" & rst("id") & """ >Ӧ�ô�ģ��</a></div>"
		strSkins = strSkins & "										<li class=""l4""><input type=""submit"" value=""Ӧ�ô�ģ��"" onClick=""window.location='user_template.asp?action=savedefault&teamid="&teamid&"&radiobutton=" & rst("id") & "'"" /></li>"
		strSkins = strSkins & "									</ul>	" & VBCRLF & VBCRLF

		iCount = iCount+1
		if iCount >=iPage then exit do
		rst.movenext
	Loop
	GetSkinList = strSkins
End Function
%>
