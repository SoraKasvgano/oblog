<!--#include file="user_top.asp"-->
<div id="main">
  <div class="submenu">
    <div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
    <div class="submenu_content">
    </div>
  </div>
  <div class="content">
    <div class="content_top">
            <div class="side_d1 side11"></div>
            <div class="side_d2 side21"></div>
    </div>

    <div class="content_body">
<%
G_P_PerMax = 20
Dim sAction,Sql, SqlQuery, sYear, sMonth, sTitle, rs
If sYear = "" Then
    sYear = Year(Date)
    sMonth = Month(Date)
Else
    sYear = CInt(Reuqest("q_year"))
    sMonth = CInt(Reuqest("q_month"))
End If
sAction = Trim(Request("action"))


G_P_FileName = "user_myactions.asp?action=" & sAction & "&q_year=" & sYear & "&q_month=" & sMonth
If Request("page") <> "" Then
    currentPage = CInt(Request("page"))
Else
    currentPage = 1
End If
SqlQuery = " And Year(addtime)=" & sYear & " And Month(addtime)=" & sMonth
select Case sAction
    Case "c"
        sTitle = "评论"
        G_P_Guide="我发表的评论"
        Sql = "select * From oblog_comment Where comment_user='" & oblog.l_uname & "' And isguest=0 " & SqlQuery
        Call Main
    Case "m"
        sTitle = "留言"
        G_P_Guide="我发表的留言"
        Sql = "select * From oblog_message Where message_user='" & oblog.l_uname & "' And isguest=0 " & SqlQuery
        Call Main
    Case "t"
        sTitle = "引用"
        G_P_Guide="谁引用了我的文章"
        Sql = "select * From oblog_trackback Where userid=" & oblog.l_uid & SqlQuery
        Call Main
    Case "del"
    	Call DelTrackBacks
    Case Else
    	sTitle = "评论"
        G_P_Guide="我发表的评论"
        Sql = "select * From oblog_comment Where comment_user='" & oblog.l_uname & "' And isguest=0 " & SqlQuery
        Call Main
End select
Set rs = Nothing
%>
    </div>

    <div class="content_bot">
            <div class="side_e1 side12"></div>
            <div class="side_e2 side22"></div>
    </div>

  </div>
</div>




</body>
</html>
<%
Sub Main()
%>
<style type="text/css">
<!--
    #list ul{ width: 95%}
    #list ul li.t0 { width: 50px}
    #list ul li.t1 { width: 600px}
    #list ul li.t2 { width: 150px}
    #list ul li.t3 { width: 50px}
-->
</style>
<%
Dim j,nYear
nYear=Year(Date)
%>
<div id="list">
    <h2>
  <form name="form1" action="user_myactions.asp?action=<%=sAction%>" method="get">
快速查找:
        <select size=1 name="usersearch">
          <%For j=nYear-5 To nYear%>
          	<option value="<%=j%>"><%=j%></option>
          <%Next%>
        </select>年
         <select name="q_month" id="q_month">
        <%For j=1 To 12%>
          	<option value="<%=j%>"><%=j%></option>
         <%Next%>
      </select>月
      <input type="submit"  value=" 搜索 ">
  </form>
    </h2>
<%
    Dim G_P_Guide
    G_P_Guide = "<h1>当前选择&nbsp;&gt;&gt;&nbsp;"
    Set rs = Server.CreateObject("Adodb.RecordSet")
 '   Response.Write(sql)
    rs.Open Sql, Conn, 1, 1
    If rs.EOF And rs.bof Then
        G_P_Guide = G_P_Guide & " (共有0个"& sTitle &")</h1></div>"
        Response.write G_P_Guide
    Else
        G_P_AllRecords = rs.recordcount
        G_P_Guide = G_P_Guide & " (共有" & G_P_AllRecords & "个"& sTitle &")</h1>"
        Response.write G_P_Guide
        If currentPage < 1 Then
            currentPage = 1
        End If
        If (currentPage - 1) * G_P_PerMax > G_P_AllRecords Then
            If (G_P_AllRecords Mod G_P_PerMax) = 0 Then
                currentPage = G_P_AllRecords \ G_P_PerMax
            Else
                currentPage = G_P_AllRecords \ G_P_PerMax + 1
            End If
        End If
        If currentPage = 1 Then
            showContent
            Response.write oblog.showpage(True, True, "篇"& sTitle )
        Else
            If (currentPage - 1) * G_P_PerMax < G_P_AllRecords Then
                rs.Move (currentPage - 1) * G_P_PerMax
                Dim bookmark
                bookmark = rs.bookmark
                showContent
                Response.write oblog.showpage(True, True, "篇"& sTitle )
            Else
                currentPage = 1
                showContent
                Response.write oblog.showpage(True, True, "篇"& sTitle )
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Sub showContent()
    Dim i,sTitle,sUrl
    i = 0%>
    <form name="myform" method="Post" action="user_myactions.asp" onsubmit="return confirm('确定要执行选定的操作吗？');">
    <ul class="list_top">
    <%If sAction="t" Then%><li class="t0">选中</li><%End If%>
    <li class="t1">内容</li>
    <li class="t2">时间</li>
    <%If sAction="t" Then%><li class="t3">操作</li><%End If%>
    </ul>
    <%Do While Not rs.EOF %>
	    <ul class="list_content" onmouseover="fSetBg(this)" onmouseout="fReBg(this)">
	    <%
	    select Case sAction
	    	Case "c"
	    	%>
		    	<li class="t1"> <a href=go.asp?logid=<%=rs("mainid")%>#<%=rs("commentid")%> target=_blank><%=oblog.filt_html(rs("commenttopic"))%><br/></a><%=Left(oblog.filt_html(rs("comment")),100) & "..."%> </li>
				<li class="t2"> <%=rs("addtime")%></li>
			<%
			Case "m"
			%>
				<li class="t1"><a href=go.asp?messageid=<%=rs("messageid")%> target=_blank><%=oblog.filt_html(rs("messagetopic"))%></a><br/><%=Left(oblog.filt_html(rs("message")),100) & "..."%>   </li>
	    		<li class="t2"><%if rs("addtime")<>"" then  Response.write rs("addtime") else Response.write "&nbsp;" %> </li>
	    	<%
	    	Case "t"
	    		%>
	    		<li class="t0"><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("id"))%>'></li>
	    		<li class="t1"><%=oblog.filt_html(rs("tb_url"))%><br/><%=oblog.filt_html(rs("IP"))%><%=Left(oblog.filt_html(rs("except")),100) & "..."%>   </li>
	    		<li class="t2"><%if rs("addtime")<>"" then  Response.write rs("addtime") else Response.write "&nbsp;" %> </li>
	    		<li class="t3"><a href="user_myactions.asp?action=del&id=<%=rs("ID")%>" onClick='return confirm("确定要删除此引用吗？");'>删除</a></li>
	    	<%
	    End select
	    	%>
	    </ul>
	    <%
	    i = i + 1
	    If i >= G_P_PerMax Then Exit Do
	    rs.movenext
	Loop
	If sAction="t" Then
	%>

	    <ul class="list_bot"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
	              全选</li>
	    <input name="action"  type="hidden" value="del" >
	    <input type="submit" name="Submit" value=" 删除 ">
	    </ul>
	    </form>
	</div>
	<%
	End If
End Sub

Sub DelTrackBacks()
	Dim sIDs
	sIDs=Request("Id")
    If  sIDs = "" Then
        oblog.adderrstr ("错误：请指定要删除的留言！")
        oblog.showusererr
        Exit Sub
    End If
    If InStr(sIDs, ",") > 0 Then
        sIDs = FilterIDs(sIDs)
		If sIDs<>"" Then oblog.execute("Delete From oblog_trackback Where userid=" & oblog.l_uid & " And Id In (" & sIDs  &")")
    Else
    	If CheckInt(sIDs)<>"" Then oblog.execute("Delete From oblog_trackback Where userid=" & oblog.l_uid & " And Id =" & sIDs)
    End If
    oblog.ShowMsg "删除留言成功!", ""
End Sub


%>