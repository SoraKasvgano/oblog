<!--#include file="user_top.asp"-->
<%
Dim DivId
DivId=Request("div")
If DivId="" Then DivId=11
DivId=Cint(DivId)
%>
<script src="oBlogStyle/move.js" type="text/javascript"></script>
<script src="inc/function.js" type="text/javascript"></script>
<script type="text/javascript">
function getImg(){
	if (document.oblogform.ico.value!=""){
		document.oblogform.imgIcon.src=document.oblogform.ico.value;
	}
}
</script>
<table id="TableBody" class="Setting" cellpadding="0">
	<thead id="TableBody_thead">
		<tr>
			<th>
				<ul id="TabPage2">
					<li id="left_tab1" <%If divId=11 or divId=12 or divId=13 or divId=14 or divId=15 or divId=16 Then%>class="Selected"<%End If%> onClick="javascript:border_left('TabPage2','left_tab1');self.location='user_setting.asp?action=0&div=11'" title="��������">��������</li>
					<li id="left_tab2" <%If divId=21 or divId=22 or divId=23 Then%>class="Selected"<%End If%> onClick="javascript:border_left('TabPage2','left_tab2');self.location='user_setting.asp?action=userinfo&div=21'" title="��������">�û�����</li>
					<li id="left_tab3" <%If divId=31 or divId=32 or divId=33 Then%>class="Selected"<%End If%> onClick="javascript:border_left('TabPage2','left_tab3');self.location='user_setting.asp?action=blogteam&div=31'" title="��������">��ͬ׫д</li>
				</ul>

				<div id="left_menu_cnt">
					<ul id="dleft_tab1" <%If divId=11 or divId=12 or divId=13 or divId=14 or divId=15 or divId=16 Then%>class="Selected" style="display:block;"<%End If%>>
						<li id="now11" <%If divId=11 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=0&div=11" title="��������">��������</a></li>
						<li id="now12" <%If divId=12 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=placard&div=12" title="���͹���">���͹���</a></li>
						<li id="now13" <%If divId=13 Then%>class="Selected"<%End If%>><a href="user_friendurl.asp" title="������������">������������</a></li>
						<li id="now14" <%If divId=14 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=links&div=14" title="�߼��༭��������">�߼��༭��������</a></li>
						<li id="now15" <%If divId=15 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogpassword&div=15" title="���ܲ���">���ܲ���</a></li>
						<li id="now16" <%If divId=16 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogstar&div=16" title="���벩��֮��">���벩��֮��</a></li>
					</ul>
					<ul id="dleft_tab2" <%If divId=21 or divId=22 or divId=23 Then%>class="Selected" style="display:block;"<%End If%>>
						<li id="now21" <%If divId=21 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=userinfo&div=21" title="��������">��������</a></li>
						<li id="now22" <%If divId=22 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=userpassword&div=22" title="��������">��������</a></li>
						<li id="now23" <%If divId=23 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=userpassword&div=23" title="���뱣��">���뱣��</a></li>
					</ul>
					<ul id="dleft_tab3" <%If divId=31 or divId=32 or divId=33 Then%>class="Selected" style="display:block;"<%End If%>>
						<li id="now31" <%If divId=31 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogteam&div=31" title="�Ŷӳ�Ա����">�Ŷӳ�Ա����</a></li>
						<li id="now32" <%If divId=32 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogteam&div=32" title="�ҹ�����Ŷ�">�Ҽ�����Ŷ�</a></li>
						<li id="now33" <%If divId=33 Then%>class="Selected"<%End If%>><a href="user_setting.asp?action=blogteam&div=33" title="�ҹ�����Ŷ�">�������Ѽ���</a></li>
					</ul>
				</div>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
<%
	If oblog.ChkPost() = False Then
        oblog.adderrstr ("ϵͳ��������ⲿ�ύ��")
    End If
Dim action
action = Request("action")
select Case action
    Case "blogteam"
        Call blogteam
    Case "links"
        Call links
    Case "savelinks"
        Call savelinks
    Case "placard"
        Call placard
    Case "saveplacard"
        Call saveplacard
    Case "blogpassword"
        Call blogpassword
    Case "addblogpassword"
        Call addblogpassword
    Case "unblogpassword"
        Call unblogpassword
    Case "savesitesetup"
        Call savesitesetup
    Case "blogstar"
        Call blogstar
    Case "saveblogstar"
        Call saveblogstar
    Case "saveuserlostpass"
        Call saveuserlostpass
    Case "userpassword"
        Call userpassword
    Case "saveuserpassword"
        Call saveuserpassword
    Case "saveuserinfo"
        Call saveuserinfo
    Case "userinfo"
        Call userinfo
	Case "aoboaccount"
		Call aoboaccount
    Case Else
        Call sitesetup
End select
%>

			</td>
		</tr>
	</tbody>
</table>
</body>
</html>
<%
Sub blogteam()
%>
				<div id="dTab31">
					<iframe class="FrmID" src="user_blogteam.asp?action=userteam&div=<%=divid%>" frameborder="0" scrolling="no"></iframe>
				</div>
<%
End Sub

Sub saveplacard()
    Dim rs, userplacard
    userplacard = oblog.filt_astr(Request.Form("edit"), 20000)
	If oblog.l_Group(15,0)=0 Then userplacard = FilterJS(userplacard)
	'-============================ȡ������ע�����ù����޸���֤����֤
	'If Not oblog.codepass Then
	'		oblog.adderrstr ("��֤�������ˢ�º��������룡")
	'		oblog.showusererr
	'		Response.end
	'end if
	If oblog.chk_badword(userplacard) >0 Then
		oblog.adderrstr ("վ�㹫���д���ϵͳ��������ַ�!")
		oblog.showusererr
		Exit Sub
	End if
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select user_placard from [oblog_user] where userid=" & oblog.l_uid, conn, 1, 3
    rs(0) = oblog.filtpath(userplacard)
    rs.Update
    rs.Close
    Dim blog
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_placard oblog.l_uid
    Set rs = Nothing
    oblog.ShowMsg "�޸Ĺ���ɹ�", ""
End Sub

Sub placard()
Dim rs
Set rs = oblog.execute("select user_placard from [oblog_user] where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab12">
						<form name="oblogform" method="post" action="user_setting.asp " <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td>
									�������������������Ƭ��������Ľ��ܣ�������Ը�����ȥ���κ���Ϣ��
								</td>
							</tr>
							<tr>
								<td>
									<span id="loadedit"  style="display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> ��������༭��...</span>
									<textarea id="edit" name="edit" style="display:none">
										<%=Server.HtmlEncode(OB_IIF(rs(0),""))%>
									</textarea>
<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
								</td>
							</tr>
							<tr>
								<td><span style="display:none;float:left;width:470px;height:30px;">��֤��:<input name="codestr" id="codestr" type="text"  size="4" maxlength="20" style="display:inline;height:18px;border:1px #1B76B7 solid;"><%=oblog.getcode%></span>
								<input name="Action" type="hidden" id="Action" value="saveplacard" />
								<input type="submit" name="Submit" id="Submit" value="�ύ�޸�" style="display:block;float:left;width:120px;height:50px;" />
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "",1,"535","240"%>
<%
Set rs = Nothing
End Sub

Sub savelinks()
    Dim rs, links
    links = oblog.filt_astr(Request.Form("edit"), 20000)
	If oblog.l_Group(15,0)=0 Then links = FilterJS(links)
	If oblog.chk_badword(links) >0 Then
		oblog.adderrstr ("���������д���ϵͳ��������ַ�!")
		oblog.showusererr
		Exit Sub
	End if
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select user_links from [oblog_user] where userid=" & oblog.l_uid, conn, 1, 3
    rs(0) = oblog.filtpath(links)
    rs.Update
    rs.Close
    Dim blog
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_links oblog.l_uid
    Set rs = Nothing
    oblog.ShowMsg "�޸��������ӳɹ�", ""
End Sub

Sub links()
Dim rs
Set rs = oblog.execute("select user_links from [oblog_user] where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab14">
						<form name="oblogform" method="post" action="user_setting.asp" <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td>
									��������������ֻ���ͼƬ��Ȼ���� <img src="images/wlink.gif" align="absbottom" /> ��ť���볬�����ӣ��Ƽ�ʹ��<a href="user_friendurl.asp">�������ӹ���</a>��
								</td>
							</tr>
							<tr>
								<td>
									<span id="loadedit" style="display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> ��������༭��...</span>
									<textarea id="edit" name="edit" style="display:none">
										<%=Server.HtmlEncode(OB_IIF(rs(0),""))%>
									</textarea >
<%If C_Editor_Type=2 Then  Server.Execute C_Editor & "/edit.asp" %>
								</td>
							</tr>
							<tr>
								<td>
								<input name="Action" type="hidden" id="Action" value="savelinks" />
								<input type="submit" name="Submit" id="Submit"  value="�ύ�޸�" />
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "",1,"535","240"%>
<%
Set rs = Nothing
End Sub

Sub blogpassword()
%>
				<div id="chk_idAll">
					<div id="dTab15">
						<form name="oblogform" method="post" action="user_setting.asp?action=addblogpassword">
						<table class="Setting_Content" cellpadding="0">
<%
If oblog.l_Group(26,0)=1 Then
%>
							<tr>
								<td class="title">
									<label for="password">���ͷ������룺</label>
								</td>
								<td>
									<form name="form1" method="post" action="user_setting.asp?action=" ><input type="password" id="password" name="password" /></br><input type="submit" name="Submit" value="ȫվ����" /></form>
									<span>���ܺ���������־����Ҫͨ��������֤����ܷ��ʡ�</br>
	ע�⣺�����������Ժ���Ҫ<a href="user_update.asp" onclick="purl('user_setting.asp?action=userpassword&div=12','��������')">���·���ȫվ</a>��</span>
								</td>
							</tr>
<%Else%>
							<tr>
								<td class="title">
									ϵͳ��ֹ������վ���ܣ�
								</td>
								<td>
									���֮ǰ����������վ���ܹ��ܣ�֮ǰ���ܵ���������ͨ��ԭ���ķ�ʽ���ʡ�
								</td>
							</tr>
<%End If%>
							<tr>
								<td class="title">
								</td>
								<td>
									<form name="form2" method="post" action="user_setting.asp?action=unblogpassword" /><input type="submit" name="Submit" id="Submit" value="�����վ������뱣��" /></form>
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%end sub

Sub addblogpassword()
	If oblog.l_Group(26,0)=0 Then
		oblog.ShowMsg "ϵͳ��������վ����!",""
		Response.End
	End If
    Dim password, strtmp, blog
	password=Trim(Request("password"))
	if password="" then
		oblog.ShowMsg "���벻��Ϊ��!",""
		Response.End
	end if
    password = md5(password)
    oblog.execute("update [oblog_user] set blog_password='"&password&"' where userid="&oblog.l_uid)
    oblog.execute ("update [oblog_log] set blog_password=1 where userid=" & oblog.l_uid)
	oblog.execute ("update [oblog_album] set ishide=2 where userid=" & oblog.l_uid&" and ishide<>1")

    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_index 0
    blog.update_message 0
    Set blog = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName)=""
	oblog.CheckUserLogined()
    oblog.ShowMsg "������վ����ɹ�,�����¸���ȫվ�ſɻ�ð�ȫ�ļ��ܱ�����", ""
End Sub

Sub unblogpassword()
    Dim upath, blog
    oblog.execute ("update [oblog_user] set blog_password=null where userid=" & oblog.l_uid)
	oblog.execute ("update [oblog_album] set ishide=0 where userid=" & oblog.l_uid&" and ishide=2")
	oblog.execute ("update [oblog_log] set blog_password=0 where userid=" & oblog.l_uid)
	oblog.execute ("update [oblog_log] set isspecial=0 where userid=" & oblog.l_uid&" and ishide<>1 and (ispassword is null or ispassword='') and isneedlogin<>1 and (viewscores is null or viewscores=0) and (viewgroupid is null or viewgroupid=0 ) ")
    upath = Server.MapPath(oblog.l_udir)
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_index 0
    blog.update_message 0
    Set blog = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName)=""
	oblog.CheckUserLogined()
    oblog.ShowMsg "ȡ������ɹ�,�����¸���ȫվ�ſ�ȫ�����ܣ�", ""

End Sub

Sub sitesetup()
Dim rs, sstr, sublist, us, i,sstr1
If oblog.l_Group(6,0) = 0 Then sstr = "disabled"
If oblog.l_Group(7,0) = 0 Then sstr1 = "disabled"
Set rs = oblog.execute("select * from oblog_user where userid=" & oblog.l_uid)
us = rs("user_info")
If us = "" Or IsNull(us) Then
    sublist = 0
Else
    us = Split(us, "$")
    If us(0) <> "" Then sublist = CInt(us(0)) Else sublist = 0
End If

Dim user_domain,custom_domain
user_domain = oblog.filt_html(Trim(rs("user_domain")))
if true_domain=1 Then
	custom_domain = oblog.filt_html(Trim(rs("custom_domain")))
End if
If user_domain = "" Or IsNull(user_domain) Then
	sstr = ""
End If
If custom_domain = "" Or IsNull(custom_domain) Then
	sstr1 = ""
End If
%>
				<div id="chk_idAll">
					<div id="dTab11">
						<form name="oblogform" action="user_setting.asp" method="post">
						<table class="Setting_Content" cellpadding="0">
<%if Trim(oblog.cacheConfig(4))<>"" and oblog.cacheConfig(5) = "1" then%>
							<tr>
								<td class="title">
									<label for="user_domain">������</label>
								</td>
								<td>
									<input name="user_domain" id="user_domain" type="text" value="<%=user_domain%>" size="10" maxlength="20" <%=sstr%> /> <select name="user_domainroot" <%=sstr%>><%=oblog.type_domainroot(rs("user_domainroot"),0)%></select><input type="hidden" name="old_userdomain" value="<%=user_domain%>">
								</td>
							</tr>
<%end if%>
<%if true_domain=1 and oblog.l_Group(7,0) = "1" then%>
								<tr>
									<td class="title">
										<label for="custom_domain">���ҵĶ���������</label>
									</td>
									<td>
										<input name="custom_domain" id="custom_domain" type="text" value="<%=custom_domain%>" size="30" maxlength="50" <%=sstr1%> />
									<span>��ǰ��ȷ������ip�Ѿ����������ͷ�������</span>
									</td>
								</tr>
<%end if%>
								<tr>
									<td class="title">
										<label for="blogname">վ�����ƣ�</label>
									</td>
									<td>
										<input name="blogname" id="blogname" type="text" value="<%=oblog.filt_html(rs("blogname"))%>" size="30" maxlength="20" />
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_classid">վ�����</label>
									</td>
									<td>
										<select name="user_classid" id="user_classid" >
											<%=oblog.show_class("user",rs("user_classid"),0)%>
										</select>
									</td>
								</tr>
								<tr>
									<td class="title">
										�����Ҽ��벩���Ŷӣ�
									</td>
									<td>
										<label><input type="radio" value="1" name="en_blogteam" <%if rs("en_blogteam")<>0 then Response.write "checked"%> />��&nbsp;&nbsp;</label>
										<label><input type=radio value="0" name="en_blogteam" <%if rs("en_blogteam")=0 then Response.write "checked"%> />��</label>
										<span>������˽��Լ��������������Ŷӡ�</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										�Ƿ�����ת��URL��
									</td>
									<td>
										<label><input type="radio" value="1" name="hideurl" <%if rs("hideurl")=1 then Response.write "checked"%> />�� &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="hideurl" <%if rs("hideurl")=0 Or  OB_IIF(rs("hideurl"),"")="" then Response.write "checked"%> />��</label>
										<span>��������˽������������ʵ������ֻ�ܿ�����ѡ��Ķ���������</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										������־�Ƿ����б���ʾ��
									</td>
									<td>
										<label><input type="radio" value="1" name="sublist" <%if sublist=1 then Response.write "checked"%> />�� &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="sublist" <%if sublist=0 then Response.write "checked"%> />��</label>
										<span>�������������·��ർ����������־��������У��ر�����ʾ��־���ݡ�</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										��ҳ��־�Ƿ����б���ʾ��
									</td>
									<td>
										<label><input type="radio" value="1" name="indexlist" <%if rs("indexlist")=1 then Response.write "checked"%> />�� &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="indexlist" <%if rs("indexlist")=0 then Response.write "checked"%> />��</label>
										<span>������������ҳ������������־��������С�</span>
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_showlogword_num">��־Ĭ�ϲ�����ʾ������</label>
									</td>
									<td>
										<input name="user_showlogword_num" id="user_showlogword_num" type="text"  value="<%=OB_IIF(rs("user_showlogword_num"),"500")%>" size="5" />
										<span>���ó�0����ʾȫ�ġ�</span>
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_showlog_num">ÿҳ��ʾ��־ƪ����</label>
									</td>
									<td>
										<input name="user_showlog_num" id="user_showlog_num" type="text" id="user_showlog_num" value="<%=OB_IIF(rs("user_showlog_num"),"20")%>" size="5" />
										<span>��ҳ��ʾ��־�������벻Ҫ����Ϊ0����̫������֡�</span>
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_photorow_num">ÿ����ʾ��Ƭ����</label>
									</td>
									<td>
										<input name="user_photorow_num" id="user_photorow_num" type="text" id="user_photorow_num" value="<%=OB_IIF(rs("user_photorow_num"),"4")%>" size="5" />
										<span>���ҳ��ÿ����ʾ��Ƭ����</span>
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_shownewcomment_num">��ʾ���»ظ�������</label>
									</td>
									<td>
										<input name="user_shownewcomment_num" id="user_shownewcomment_num" type="text" value="<%=OB_IIF(rs("user_shownewcomment_num"),"8")%>" size="5" />
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_shownewlog_num">��ʾ������־������</label>
									</td>
									<td>
										<input name="user_shownewlog_num" id="user_shownewlog_num" type="text" value="<%=OB_IIF(rs("user_shownewlog_num"),"8")%>" size="5" />
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="user_shownewmessage_num">��ʾ��������������</label>
									</td>
									<td>
										<input name="user_shownewmessage_num" id="user_shownewmessage_num" type="text" value="<%=OB_IIF(rs("user_shownewmessage_num"),"8")%>" size="5" />
										<span class="red">��Ҫ������ҳ��Ż���Ч��</span>
									</td>
								</tr>
								<tr>
									<td class="title">
										��־��������˳��
									</td>
									<td>
										<label><input type="radio" value="1" name="comment_isasc" <%if rs("comment_isasc")=1 then Response.write "checked"%> />ʱ��˳�� &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="comment_isasc" <%if rs("comment_isasc")=0 then Response.write "checked"%> />ʱ�䵹��</label>
									</td>
								</tr>
								<tr>
									<td class="title">
										�༭�����ͣ�
									</td>
									<td>
										<label><input type="radio" value="2" name="isubbedit" <%if rs("isubbedit")=2 then Response.write "checked"%> />3.x�汾(�޷�����֧�ַ�IE�����) &nbsp;&nbsp;</label>
										<label><input type="radio" value="1" name="isubbedit" <%if rs("isubbedit")=1 then Response.write "checked"%> />4.x�汾</label>
									</td>
								</tr>
								<tr>
									<td class="title">
										�Ƿ�������־�������û��Ƽ���
									</td>
									<td>
										<label><input type="radio" value="1" name="isdigg" <%if OB_iif(rs("isdigg"),1)=1 then Response.write "checked"%> />���� &nbsp;&nbsp;</label>
										<label><input type="radio" value="0" name="isdigg" <%if rs("isdigg")=0 then Response.write "checked"%> />������</label>
									</td>
								</tr>
								<tr>
									<td class="title">
										<label for="siteinfo">վ���飺</label>
									</td>
									<td>
										<textarea name="siteinfo" id="siteinfo" cols="40" rows="5"><%=oblog.filt_html(rs("siteinfo"))%></textarea>
									</td>
								</tr>
								<tr>
									<td class="title">
									</td>
									<td>
										<input name="action" type="hidden" value="savesitesetup" />
										<input type="submit" id="Submit" value="�����޸�" />
									</td>
								</tr>
							</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
Set rs = Nothing
End Sub

Sub savesitesetup()
    Dim user_domain, user_domainroot, blogname, user_classid, en_blogteam, user_showlogword_num, user_showlog_num, user_shownewcomment_num, user_shownewlog_num, user_shownewmessage_num, hideurl, comment_isasc, siteinfo, user_photorow_num, custom_domain,isdigg
    Dim rs, blog
    user_domain = LCase(Trim(Request("user_domain")))
    user_domainroot = Trim(Request("user_domainroot"))
    blogname = EncodeJP(Trim(Request("blogname")))
    user_classid = Trim(Request("user_classid"))
    en_blogteam = Trim(Request("en_blogteam"))
    user_showlogword_num = Trim(Request("user_showlogword_num"))
    user_showlog_num = Trim(Request("user_showlog_num"))
    user_photorow_num = Trim(Request("user_photorow_num"))
    user_shownewcomment_num = Trim(Request("user_shownewcomment_num"))
    user_shownewlog_num = Trim(Request("user_shownewlog_num"))
    user_shownewmessage_num = Trim(Request("user_shownewmessage_num"))
    hideurl = Trim(Request("hideurl"))
    comment_isasc = Trim(Request("comment_isasc"))
    siteinfo = EncodeJP(Trim(Request("siteinfo")))
    custom_domain = Trim(Request("custom_domain"))
    isdigg = Trim(Request("isdigg"))
    If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1 And oblog.l_Group(6,0) = 1 Then
        If user_domain = "" Or oblog.strLength(user_domain) > 20 Then oblog.adderrstr ("��������Ϊ��(���ܴ���14���ַ�)��")
        If user_domain <> Request("old_userdomain") And oblog.strLength(user_domain) < 4 Then oblog.adderrstr ("��������С��4���ַ���")
        'If oblog.chk_regname(user_domain) Then oblog.adderrstr ("������ϵͳ������ע�ᣡ")
        If oblog.chk_badword(user_domain) > 0 Then oblog.adderrstr ("�����к���ϵͳ��������ַ���")
        If oblog.chkdomain(user_domain) = False Then oblog.adderrstr ("�������Ϲ淶��ֻ��ʹ��Сд��ĸ�����֣�")
        If user_domainroot = "" Then oblog.adderrstr ("����������Ϊ�գ�")
		If oblog.CheckDomainRoot(user_domainroot,0) = False Then oblog.adderrstr  ("���������Ϸ���")
    End If
    If oblog.strLength(siteinfo) > 255 Then oblog.adderrstr ("վ���鲻�ܴ���255���ַ���")
    If oblog.chk_badword(blogname) > 0 Then oblog.adderrstr ("blog���к���ϵͳ��������ַ���")
    If Not IsNumeric(user_showlogword_num) Then
        oblog.adderrstr ("��־Ĭ�ϲ�����ʾ��������Ϊ���֣�")
    End If
	If oblog.CacheConfig(48)="1" Then
		Dim rsreg
		Set rsreg=oblog.execute("select Count(userid) From oblog_user Where blogname='" & ProtectSQL(blogname) & "' and userid<> " & oblog.l_uid)
    	If rsreg(0)>0 Then
    		oblog.adderrstr  ("��ʹ�õĲ�������: " & blogname & " �ѱ�����ʹ�ã��������������")
    	End If
    	rsreg.Close
	End If
    If Not IsNumeric(user_showlog_num) Then
        oblog.adderrstr ("ÿҳ��ʾ��־��������Ϊ���֣�")
    Else
        user_showlog_num = CLng(user_showlog_num)
        If user_showlog_num > 50 Then oblog.adderrstr ("ÿҳ��ʾ��־��������С��50��")
    End If
    If Not IsNumeric(user_photorow_num) Then
        oblog.adderrstr ("ÿ����ʾ��Ƭ��������Ϊ���֣�")
    Else
        user_photorow_num = CLng(user_photorow_num)
        If user_photorow_num > 50 Then oblog.adderrstr ("ÿ����ʾ��Ƭ��������С��50��")
    End If

    If Not IsNumeric(user_shownewcomment_num) Then
        oblog.adderrstr ("��ʾ���»ظ���������Ϊ���֣�")
    Else
        user_shownewcomment_num = CLng(user_shownewcomment_num)
        If user_shownewcomment_num > 50 or user_shownewcomment_num<1 Then oblog.adderrstr ("��ʾ���»ظ��������ܴ���50����С��1��")
    End If

    If Not IsNumeric(user_shownewlog_num) Then
        oblog.adderrstr ("��ʾ������־��������Ϊ���֣�")
    Else
        user_shownewlog_num = CLng(user_shownewlog_num)
        If user_shownewlog_num > 50 or user_shownewlog_num<1 Then oblog.adderrstr ("��ʾ������־�������ܴ���50����С��1��")
    End If

    If Not IsNumeric(user_shownewmessage_num) Then
        oblog.adderrstr ("��ʾ����������������Ϊ���֣�")
    Else
        user_shownewmessage_num = CLng(user_shownewmessage_num)
        If user_shownewmessage_num > 50 or user_shownewmessage_num<1 then oblog.adderrstr ("��ʾ���������������ܴ���50����С��1��")
    End If
   If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1 And oblog.l_Group(6,0) = 1 Then
        Set rs = oblog.execute("select userid from oblog_user where user_domain='" & oblog.filt_badstr(user_domain) & "' and user_domainroot='" & oblog.filt_badstr(user_domainroot) & "' and userid<>" & oblog.l_uid)
        If Not rs.EOF Or Not rs.bof Then oblog.adderrstr ("ϵͳ���Ѿ�������������ڣ������������")
    End If
    If true_domain = 1 And custom_domain <> "" Then
        If oblog.chk_badword(custom_domain) > 0 Then oblog.adderrstr ("�󶨵Ķ��������к���ϵͳ��������ַ���")
        Set rs = oblog.execute("select userid from oblog_user where custom_domain='" & oblog.filt_badstr(custom_domain) & "'" & " and userid<>" & oblog.l_uid)
        If Not rs.EOF Or Not rs.bof Then oblog.adderrstr ("ϵͳ���Ѿ��������˰�������������������������������ϵ����Ա��")
    End If
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    If hideurl = "" Or IsNull(hideurl) Then hideurl = 0
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select * from oblog_user where userid=" & oblog.l_uid, conn, 1, 3
    If Not rs.EOF Then
		rs("blogname") = oblog.filt_astr(blogname, 50)
		If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1 Then
			If oblog.l_Group(6,0) = 1 Or Trim (rs("user_domain")) = "" Or IsNull(rs("user_domain")) Then
				rs("user_domain") = user_domain
				rs("user_domainroot") = user_domainroot
			End if
        End If
		If true_domain = 1 Then
			If oblog.l_Group(7,0) = 1 Then
				rs("custom_domain") = custom_domain
            End If
        End If
        rs("user_classid") = user_classid
        rs("en_blogteam") = en_blogteam
        rs("user_showlogword_num") = user_showlogword_num
        rs("user_showlog_num") = user_showlog_num
        rs("user_photorow_num") = user_photorow_num
        rs("user_shownewcomment_num") = user_shownewcomment_num
        rs("user_shownewlog_num") = user_shownewlog_num
        rs("user_shownewmessage_num") = user_shownewmessage_num
        rs("hideurl") = hideurl
        rs("comment_isasc") = comment_isasc
        rs("siteinfo") = siteinfo
        rs("isubbedit") = OB_IIF(Request.Form("isubbedit"),2)
        rs("user_info") = CInt(Request.Form("sublist")) & "$0"
		rs("indexlist")=cint(Request.Form("indexlist"))
		rs("isdigg") = OB_IIF(isdigg,1)
        rs.Update
        rs.Close
    End If
    Set rs = Nothing
    Set blog = New class_blog
    blog.userid = oblog.l_uid
    blog.update_blogname
    Set blog = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName) = ""
	Oblog.CheckUserLogined
    oblog.ShowMsg "�������óɹ�!", ""
End Sub

Sub saveblogstar()
    Dim rs, picurl, bloginfo, blogname
    picurl = Trim(Request("ico"))
    bloginfo = Trim(Request("bloginfo"))
    blogname = Trim(Request("blogname"))
    If picurl = "" Or oblog.strLength(picurl) > 250 Then oblog.adderrstr ("ͼƬ���ӵ�ַ����Ϊ��,�Ҳ��ܴ���250���ַ���")
    If bloginfo = "" Or oblog.strLength(bloginfo) > 250 Then oblog.adderrstr ("վ����ܲ���Ϊ��,�Ҳ��ܴ���250���ַ���")
    If blogname = "" Or oblog.strLength(blogname) > 50 Then oblog.adderrstr ("����������Ϊ��,�Ҳ��ܴ���50���ַ���")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select top 1 * from oblog_blogstar Where userid=" & oblog.l_uid, conn, 1, 3
    If rs.EOF Then
        rs.addnew
        rs("userid") = oblog.l_uid
    End If
    rs("picurl") = picurl
    rs("info") = bloginfo
    If Trim(oblog.CacheConfig(4)) <> "" And oblog.CacheConfig(5) = 1  Then
		If oblog.l_Group(6,0) = 0 Then
			rs("userurl") = "http://" & oblog.l_udomain
		Else
			rs("userurl")=oblog.CacheConfig(3)&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext
		End if
    Else
        rs("userurl")=oblog.CacheConfig(3)&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext
    End If
    rs("blogname") = blogname
	rs("username") = Oblog.l_uName
	rs("usernickname") = Oblog.l_uNickname
    rs.Update
    rs.Close
    Set rs = Nothing
    oblog.ShowMsg "�ύ��ɣ���ȴ�����Ա���ͨ����", ""
End Sub

Sub blogstar()
    Dim rs, strTitle, strBlogName, strPicUrl, strBlogInfo, intState
    Set rs = Server.CreateObject("ADODB.RecordSet")
    rs.open "select * from oblog_blogstar Where userid=" & oblog.l_uid, conn, 1, 1
    If rs.EOF Then
        strTitle = "��Ŀǰ��û������"
        intState = -1
    Else
        strPicUrl = rs("picurl")
        strBlogName = rs("blogname")
        strBlogInfo = rs("info")
        intState = rs("ispass")
        If intState = 1 Then
            strTitle = "��Ŀǰ�Ѿ��ǲ���֮�ǣ����ϲ��ɸ���"
        Else
            strTitle = "��Ŀǰ���ڵȴ�����У������޸�֮ǰ�ύ������"
        End If
    End If
    rs.Close
    Set rs = Nothing
%>
				<div id="chk_idAll">
					<div id="dTab21">
						<form name="oblogform" method="post" action="user_setting.asp?action=saveblogstar">
						<table  class="UserInfo" align="center" cellpadding="0" cellspacing="1">
							<tr>
								<td colspan="4">
									<div class="red"><%=strTitle%></div>
								</td>
							</tr>
							<tr>
								<td class="title">
									�������֣�
								</td>
								<td colspan="3">
									<input type="text" maxlength="50" size="60" name="blogname" value="<%=strBlogname%>" />
									<span>blog�����벻Ҫ����50�֡�</span>
								</td>
							</tr>
							<tr>
								<td class="title">�û�ͷ��</td>
								<td colspan="3"><div class="user_face">
									<span><img src="<%=ProIco(strPicUrl,1)%>" class="face" id="imgIcon" width=<%=C_UserIcon_Width%> height=<%=C_UserIcon_Width%> /></span>
									<p><iframe id="d_file" frameborder="0" src="upload.asp?tMode=9&re=" width="400" height="30" scrolling="no"></iframe></p>
									<p>ֻ֧��jpg��gif��png��С��200k��Ĭ�ϳߴ�Ϊ48px*48px<br /><br /></p>
									<p><select name="usertile" id="usertile" onchange="setusertile();"><option value="0">Ĭ��</option><%=GetUserTile%></select>����<label>ͷ���ַ��<input name="ico"  id = "ico" type="text" value="<%=oblog.filt_html(strPicUrl)%>" size="60" maxlength="200"  onblur="getImg();" /></label></p>
								</div></td>
							</tr>
<!-- 							<tr>
								<td class="title">
									ͼƬ��ַ��
								</td>
								<td colspan="3">
									<input type="text" maxlength="250" size="60" name="picurl" value="<%=strPicUrl%>" />
									<span>ͼƬ��ַ���Է��������Ƭ��վ��logo����վ������ͼ��<br />��ͼƬ�ߴ������С��130*100���ң��Ա����Ա��������</span>
								</td>
							</tr>
 -->							<tr>
								<td class="title">
									blog���ܣ�
								</td>
								<td colspan="3">
									<textarea name="bloginfo" cols="50" rows="5"><%=strBlogInfo%></textarea><br />
									<span>��дBlog���ܺ��������ɣ�����ͨ���󽫹�����ʾ������Ա��Ȩ����Щ�������ʵ�������</span>
								</td>
							</tr>
							<tr>
								<td colspan="4" align="center">
<%
select Case intState
	Case -1
%>
									<input type="submit" id="Submit" value="�ύ��������" />
<%
Case 0
%>
									<input type="submit" id="Submit" value="�޸���������" />
<%
Case 1
%>
									�Ѿ���ȷ��Ϊ����֮�ǣ����ϲ��ɸ��ģ������Ҫ�޸Ļ����������Ա��ϵ��
<%End select%>
								</td>
							</tr>
						</table>
						</form>
					</div>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
End Sub
Sub userpassword()
if is_ot_user=1 then
	Response.Redirect(ot_modifypass1)
end if
Dim rs
Set rs = oblog.execute("select question from oblog_user where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab22"<%If divId=23 Then%>style="display:none;"<%End If%>>
						<form  method="post" action="user_setting.asp">
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td class="title">
									ԭʼ���룺
								</td>
								<td>
									<input name="oldpassword" type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
									�����룺
								</td>
								<td>
									<input name="newpassword"  type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
									�ظ����룺
								</td>
								<td>
									<input name="newpassword1" type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
								</td>
								<td>
									<input name="action" type="hidden" value="saveuserpassword" />
									<input type="submit" id="Submit"  value=" �޸����� " />
								</td>
							</tr>
						</table>
						</form>
					</div>
					<div id="dTab23" <%If divId=22 Then%>style="display:none;"<%End If%>>
						<form name="oblogform" action="user_setting.asp" method="post">
						<table class="Setting_Content" cellpadding="0">
							<tr>
								<td class="title">
									��¼���룺
								</td>
								<td>
									<input name="password"  type="password" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
									������ʾ���⣺
								</td>
								<td>
									<input name="question"  type="text" size="30" maxlength="20" value="<%=oblog.filt_html(rs(0))%>">
								</td>
							</tr>
							<tr>
								<td class="title">
									�һ�����𰸣�
								</td>
								<td>
									<input name="answer" type="text" size="30" maxlength="20" />
								</td>
							</tr>
							<tr>
								<td class="title">
								</td>
								<td>
									<input name="action" type="hidden" value="saveuserlostpass">
									<input type="submit" id="Submit" value="ȷ���޸�">
								</td>
							</tr>
						</table>
						</form>
					</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
End Sub

Sub saveuserlostpass()
    Dim password, question, answer, rs
    password = Trim(Request("password"))
    question = Trim(Request("question"))
    answer = Trim(Request("answer"))
    If password = "" Then oblog.adderrstr ("���󣺵�¼���벻��Ϊ��!")
    If question = "" Then oblog.adderrstr ("������ʾ���ⲻ��Ϊ�գ�")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select question,answer from oblog_user where userid="&oblog.l_uid&" and password='"&md5(password)&"'",conn,1,3
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("���󣺵�¼�����������!")
        oblog.showusererr
        Exit Sub
    Else
	    If API_Enable Then
	          Dim blogAPI
              Set blogAPI = New DPO_API_OBLOG
			  blogAPI.LoadXmlFile True
              blogAPI.UserName=oblog.l_uName
              blogAPI.PassWord=password
			  blogAPI.Question=Question
              blogAPI.Answer=Answer
		      Call blogAPI.ProcessMultiPing("update")
			  Set blogAPI=Nothing
		End If

        rs("question") = question
        If answer <> "" Then rs("answer") = md5(answer)
        rs.Update
        rs.Close
        Set rs = Nothing
        oblog.ShowMsg "�޸��һ��������ϳɹ���", ""
    End If

End Sub

Sub saveuserpassword()
    Dim oldpassword, newpassword, rs
    oldpassword = Trim(Request("oldpassword"))
    newpassword = Trim(Request("newpassword"))
    If oldpassword = "" Then oblog.adderrstr ("����ԭ���벻��Ϊ��!")
    If newpassword = "" Or oblog.strLength(newpassword) > 14 Or oblog.strLength(newpassword) < 4 Then oblog.adderrstr ("���������벻��Ϊ��(���ܴ���14С��4)��")
    If newpassword <> Trim(Request("newpassword1")) Then oblog.adderrstr ("�����ظ������������!")
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select password,TruePassWord from oblog_user where userid="&oblog.l_uid&" and password='"&md5(oldpassword)&"'",conn,1,3
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        oblog.adderrstr ("����ԭ�����������!")
        oblog.showusererr
        Exit Sub
    Else

	    If API_Enable Then
			Dim blogAPI,j
			Set blogAPI = New DPO_API_OBLOG
			blogAPI.LoadXmlFile True
			blogAPI.UserName=oblog.l_uName
			blogAPI.PassWord=newpassword
			Call blogAPI.ProcessMultiPing("update")
			Set blogAPI=Nothing
			For j=0 To UBound(aUrls)
				strUrl=Lcase(aUrls(j))
				If Left(strUrl,7)="http://" Then
					Response.write("<script src="""&strUrl&"?syskey="&MD5(oblog.l_uName&oblog_Key)&"&username="&oblog.l_uName&"&password="&MD5(newpassword)&"&savecookie=0""></script>")
				End If
			Next
		End If

		Dim TruePassWord
		TruePassWord = RndPassword(16)
        rs("password") = md5(newpassword)
		rs("TruePassWord") = TruePassWord
        rs.Update
        rs.Close
        Set rs = Nothing
		oblog.savecookie oblog.l_uname, TruePassWord, 0
        oblog.ShowMsg "�޸�����ɹ�,�´���Ҫ���µ�¼��", ""
    End If
End Sub

Sub userinfo()
Dim rs
Set rs = oblog.execute("select * from oblog_user where userid=" & oblog.l_uid)
%>
				<div id="chk_idAll">
					<div id="dTab21">
						<form name="oblogform" action="user_setting.asp" method="post">
						<table  class="UserInfo" align="center" cellpadding="0" cellspacing="1">
							<tr>
								<td class="title">��¼�ɣģ�</td>
								<td><span class="user_id"><%=rs("userName")%></span></td>
								<td class="title">�û��ȼ���</td>
								<td class="userlevel"><span class="red"><%=oblog.l_Group(1,0)%></span></td>
							</tr>
							<tr>
								<td class="title">�û�ͷ��</td>
								<td colspan="3"><div class="user_face">
									<span><img src="<%=ProIco(rs("user_icon1"),1)%>" class="face" id="imgIcon" width=<%=C_UserIcon_Width%> height=<%=C_UserIcon_Width%> /></span>
									<p><iframe id="d_file" frameborder="0" src="upload.asp?tMode=9&re=" width="400" height="30" scrolling="no"></iframe></p>
									<p>ֻ֧��jpg��gif��png��С��200k��Ĭ�ϳߴ�Ϊ48px*48px<br /><br /></p>
									<p><select name="usertile" id="usertile" onchange="setusertile();"><option value="0">Ĭ��</option><%=GetUserTile%></select>����<label>ͷ���ַ��<input name="ico"  id = "ico" type="text" value="<%=oblog.filt_html(rs("user_icon1"))%>" size="60" maxlength="200"  onblur="getImg();" /></label></p>
								</div></td>
							</tr>
							<tr>
								<td class="title"><label for="nickname">�ǳƣ�</label></td>
								<td colspan="3"><input name="nickname" id="nickname" type="text" value="<%=oblog.filt_html(rs("nickname"))%>" size="30" maxlength="20" /><input type="hidden" name="o_nickname" value="<%=oblog.filt_html(rs("nickname"))%>" /></td>
							</tr>
							<tr>
								<td class="title"><label for="truename">��ʵ������</label></td>
								<td><input name="truename" id="truename" type="text" value="<%=oblog.filt_html(rs("truename"))%>" size="30" maxlength="20" /></td>
								<td class="title">�Ա�</td>
								<td>
									<label><input type="radio" value="1" name="sex" <%if rs("Sex")=1 then Response.write "checked"%> />��</label>
									&nbsp;&nbsp;
									<label><input type="radio" value="0" name="sex" <%if rs("Sex")=0 then Response.write "checked"%> />Ů</label>
								</td>
							</tr>
							<tr>
								<td class="title"><label for="y">�������ڣ�</label></td>
								<td>
									<label><input value="<%=year(rs("birthday"))%>" name="y" id="y" size="2" maxlength="4" />��</label>
									<label><input value="<%=month(rs("birthday"))%>" name="m" size="2" maxlength="2" />��</label>
									<label><input value="<%=day(rs("birthday"))%>"  name="d" size="2" maxlength="2" />��</label>
								</td>
								<td class="title">ʡ/�У�</td>
								<td><%=oblog.type_city(rs("province"),rs("city"))%></td>
							</tr>
							<tr>
								<td class="title">ְҵ��</td>
								<td><%oblog.type_job(rs("job"))%></td>
								<td class="title"><label for="Email">E-mail��</label></td>
								<td><input name="Email" id="Email" value="<%=oblog.filt_html(rs("userEmail"))%>" size="30" maxlength="50" /></td>
							</tr>
							<tr>
								<td class="title"><label for="homepage">��ҳ��</label></td>
								<td colspan="3"><input maxlength="100" size="30" name="homepage" id="homepage" value="<%=oblog.filt_html(rs("Homepage"))%>" /></td>
							</tr>
							<tr>
								<td class="title"><label for="qq">QQ���룺</label></td>
								<td><input name="qq" id="qq" value="<%=oblog.filt_html(rs("qq"))%>" size="30" maxlength="20" /></td>
								<td class="title"><label for="msn">MSN��</label></td>
								<td><input name="msn" id="msn"value="<%=oblog.filt_html(rs("Msn"))%>" size="30" maxlength="50" /></td>
							</tr>
							<tr>
								<td class="title"><label for="tel">�绰��</label></td>
								<td><input name="tel" id="tel" value="<%=oblog.filt_html(rs("tel"))%>" size="30" maxlength="50" /></td>
								<td class="title"><label for="address">ͨ�ŵ�ַ��</label></td>
								<td><input name="address" id="address" value="<%=oblog.filt_html(rs("address"))%>" size="30" maxlength="250" /></td>
							</tr>
<%
If oblog.CacheConfig(51)="1" Then
If oblog.l_Group(34,0)="1" Then
%>
							<tr>
								<td colspan="4"><font class="red"><strong>��վ����ͨ���ʼ����ֻ�������־</strong></font><br />�ڴ˴����õ�������ֻ���,������ͨ�������ַ�ʽ����Ҫ���������ݷ��͵�<font class="red"><%=oblog.CacheConfig(52)%></font>,ϵͳ���Զ��������ݲ����з���<br/>�˴��������ַ���ֻ����������վ�㹫����ʾ���ֻ�����Ŀǰֻ֧���й��ƶ�GSM����,��135~139�Ŷ�</td>
							</tr>
							<tr>
								<td class="title"><label for="postmail">�����ַ��</label></td>
								<td><input name="postmail" id="postmail" value="<%=oblog.filt_html(rs("postmail"))%>" size="30" maxlength="100" /></td>
								<td class="title"><label for="postmobile">�ֻ����룺</label></td>
								<td><input name="postmobile" id="postmobile" value="<%=oblog.filt_html(rs("postmobile"))%>" size="30" maxlength="13" /></td>
							</tr>
<%
End If
End If
%>
							<tr>
								<td colspan="4" align="center">
									<input name="action" type="hidden" value="saveuserinfo" />
									<input type="submit" id="Submit" value="���¸�������" />
								</td>
							</tr>
						</table>
						</form>
					</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
<%
Set rs = Nothing
End Sub

Sub saveuserinfo()
    Dim rs, nickname, email, birthday,usericon
    nickname = Trim(Request("nickname"))
    email = Trim(Request("email"))
    birthday = Trim(Request("y")) & "-" & Trim(Request("m")) & "-" & Trim(Request("d"))
    usericon=Trim(Request("ico"))
	If InStr(nickname,"$$$") > 0 Then oblog.adderrstr ("���ǳ�ϵͳ������ע�ᣡ")
    'If oblog.chk_regname(nickname) Then oblog.adderrstr ("���ǳ�ϵͳ������ע�ᣡ")
    If oblog.chk_badword(nickname) > 0 Then oblog.adderrstr ("�ǳ��к���ϵͳ��������ַ���")
    If oblog.strLength(nickname) > 50 Then oblog.adderrstr ("�ǳƲ��ܲ��ܴ���50�ַ���")
    '�ǳ�Ψһ���ж�
    If oblog.cacheConfig(47) = "1" And nickname <> "" And nickname <> Trim(Request("o_nickname")) Then
        Set rs = oblog.execute("select userid from oblog_user where nickname='" & ProtectSQL(nickname) & "'")
        If Not rs.EOF Or Not rs.bof Then oblog.adderrstr ("ϵͳ���Ѿ�������ǳƴ��ڣ�������ǳƣ�")
    End If
    If birthday = "--" Then
        birthday = ""
    Else
        If Not IsDate(birthday) Then oblog.adderrstr ("�������ڸ�ʽ����")
        If CLng(Trim(Request("y"))) > 2060 Then oblog.adderrstr ("������ݹ���")
        If CLng(Trim(Request("y"))) < 1900 Then oblog.adderrstr ("������ݹ�С��")
    End If
    If email = "" Then oblog.adderrstr ("�����ʼ���ַ����Ϊ�գ�")
    If Not oblog.IsValidEmail(email) Then oblog.adderrstr ("�����ʼ���ַ��ʽ����")
    If oblog.cacheConfig(22) = "1" Then
        Set rs = oblog.execute("select COUNT(userid) from oblog_user where useremail='" & ProtectSQL(email) & "' and userid<>" & oblog.l_uid)
        If rs(0) > 0 Then
			oblog.adderrstr ("ϵͳ���Ѿ������Email���ڣ������Email��")
		End If
		rs.close
    End If
	email = Replace(email,"��","-")
    Dim rsreg
	Set rsreg=Nothing
	'---------------------------------------
		'Plus: Mail to Blog Start
		Dim sMail,sMobile,rstMail
		If oblog.cacheconfig(51)="1" Then
			If oblog.l_Group(34,0)="1" Then
				sMail=Trim(Request("postmail"))
				If  sMail<>"" Then
					if not oblog.IsValidEmail(sMail) then oblog.adderrstr("���������ַ��ʽ����")
				End If
				sMobile=Trim(Request("postmobile"))
				If  sMobile<>"" Then
					If Len(sMobile) = 11 And IsNumeric(sMobile) Then
						If CInt(Left(sMobile, 3)) >= 134 And CInt(Left(sMobile, 3)) <= 139 Or CInt(Left(sMobile, 3)) = 159 Then
							'bMobile = True
						Else
							oblog.adderrstr("��������ֻ�����������ϵͳ�ݲ�֧�֣�")
						End If
					Else
						oblog.adderrstr("��������ֻ�����������ϵͳ�ݲ�֧�֣�")
					End If
				End If

				set rstMail=Server.CreateObject("adodb.recordset")
				'�ж�Mail�Ƿ��ظ�
				If  sMail<>"" Then
					rstMail.open "select * from oblog_user where postmail='" & LCase(Trim(sMail)) & "' And Userid<>" & oblog.l_uid,conn,1,1
					If Not rstMail.Eof Then
						oblog.adderrstr(sMail & " �Ѿ���ʹ��,�������������!" )
					End If
					rstMail.Close
				End If
				'�ж��ֻ������Ƿ��ظ�
				If  sMobile<>"" Then
					rstMail.open "select * from oblog_user where postMobile='" & sMobile & "' And Userid<>" & oblog.l_uid,conn,1,1
					If Not rstMail.Eof Then
						oblog.adderrstr(sMobile & " �Ѿ���ʹ��,�������������!" )
					End If
					rstMail.Close
				End If
			End If
		End if
		Set rstMail=Nothing
		'Plus: Mail to Blog End
		'---------------------------------------
    If oblog.errstr <> "" Then oblog.showusererr: Exit Sub
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open "select * from oblog_user where userid=" & oblog.l_uid, conn, 1, 3
    If Not rs.EOF Then
		If API_Enable Then
			  Dim blogAPI
			  Set blogAPI = New DPO_API_OBLOG
			  blogAPI.LoadXmlFile True
			  blogAPI.UserName=rs("UserName")
			  blogAPI.EMail=Email
			  blogAPI.Sex=CLng(Request("sex"))
			  blogAPI.QQ=oblog.filt_astr(Request("qq"),20)
			  blogAPI.MSN=oblog.filt_astr(Request("msn"),50)
			  blogAPI.truename=oblog.filt_astr(Request("truename"),20)
			  If birthday <> "" Then blogAPI.birthday=birthday
			  blogAPI.telephone=oblog.filt_astr(Request("tel"),50)
			  blogAPI.homepage=oblog.filt_astr(Request("homepage"),100)
			  blogAPI.province=oblog.filt_astr(Request("province"),18)
			  blogAPI.city=oblog.filt_astr(Request("city"),18)
			  blogAPI.address=oblog.filt_astr(Request("address"),250)
			  Call blogAPI.ProcessMultiPing("update")
			  Set blogAPI=Nothing
		End If
        rs("nickname") = oblog.filt_astr(nickname, 50)
        rs("truename") = oblog.filt_astr(Request("truename"), 20)
        rs("sex") = CLng(Request("sex"))
        rs("province") = oblog.filt_astr(Request("province"),20)
        rs("city") = oblog.filt_astr(Request("city"),20)
        If birthday <> "" Then rs("birthday") = oblog.filt_astr(birthday,20)
        rs("job") = oblog.filt_astr(Request("job"),20)
        rs("useremail") = oblog.filt_astr(email, 50)
        rs("homepage") = oblog.filt_astr(Request("homepage"), 100)
        rs("qq") = oblog.filt_astr(Request("qq"), 20)
        rs("msn") = oblog.filt_astr(Request("msn"), 50)
        rs("tel") = oblog.filt_astr(Request("tel"), 50)
        rs("address") = oblog.filt_astr(Request("address"), 250)
        rs("user_icon1") = RemoveHtml(oblog.filt_html(oblog.filt_astr(Request("ico"),200)))
		If oblog.cacheconfig(51)="1" Then
			If oblog.l_Group(34,0)="1" Then
				rs("postmail")=oblog.filt_astr(Request("postmail"),100)
				rs("postmobile")=oblog.filt_astr(Request("postmobile"),13)
			End If
		End if
        rs.Update
        rs.Close
    End If
    Set rs = Nothing
	Session ("CheckUserLogined_"&oblog.l_uName) = ""
	Oblog.CheckUserLogined
    oblog.ShowMsg "�������ϳɹ�!", ""
End Sub

Function GetUserTile()
	Dim i
	Dim strTemp
	For i = 1 To 12
		strTemp = strTemp &"<option value=""usertile"&i&".gif"">ͷ��"&i&"</option>"
	Next
	GetUserTile = strTemp
	strTemp = ""
End Function
%>
