<html>
<head>
<title>oBlog 4.x ��ǩ����˵��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/admin/style.css" type="text/css" />
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">

<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">oBlogģ����˵��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 colspan="2" class="topbg"><strong>ϵͳģ����˵��</strong>
  <tr>
<td height=23 colspan="2" class="tdbg"><p><strong><font color=#0000ff>��ģ�壨��ϵͳ��ҳ��ֻ��index.asp�ļ���Ч����</font><br>
        ע��</strong>����ʹ����ȷ�ı�ǩ������������ܳ��ֲ���Ԥ֪�Ĵ���һ��Ϊ��ʾΪ�±�Խ�磩<br>
        ������<font color=red>��ɫ����Ϊ4.0�汾��������3.x�汾��ͬ�����ı�ǩ��</font></p></td>
  </tr>
  <tr>
    <td width="25%" height=23 class="tdbg"><p>$show_log(����1,����2,����3,����4,����5,����6,����7,����8,����9)$<br>
      </p></td>
    <td width="75%" class="tdbg">�˱�ǵ�����־�����б����Ϣ������˵�����£�<br>
      ��������1��������־������<br>
      ��������2����־���ⳤ�ȣ����ַ�Ϊ��λ������������ʾ��...����<br>
      ��������3�����򷽷���Ϊ1������ʱ�䣬Ϊ2���������Ϊ3���ظ�����<br>
      ��������4���Ƿ񾫻���Ϊ1����������־��Ϊ2���þ�����־��<br>
      ��������5�����ö������ڵ���־������Ϊ��λ��<br>
      ��������6����־����id��Ϊ0��������з������־��<br>
      ��������7���Ƿ���ʾ��־ϵͳ��������Ϊ1��ʾ��Ϊ0����ʾ��<br>
      ��������8���Ƿ���ʾ��־ר������Ϊ1��ʾ��Ϊ0����ʾ��<br>
      ��������9����ʾ��Ϣ��1Ϊ��ʾ����ʱ����û���2Ϊ��ʾ����ʱ�䣬3Ϊ��ʾ�����û���4Ϊ��ʾ�����û��͵������5Ϊ��ʾ�������6Ϊ��ʾ�������ں��û���7Ϊ��ʾ�������ڣ�8Ϊ��ʾ�ظ�����0Ϊ����ʾ��</td>
  </tr>
  <tr>
<td width="25%" height=23 class="tdbg"><p>$show_userlog(����1,����2,����3,����4,����5,����6)$<br>
      </p></td>
    <td width="75%" class="tdbg">�˱�ǵ���ĳ�û���־�����б����Ϣ������˵�����£�<br>
      	��������1��userid��<br>
		��������2��������־������<br>
		��������3����־���ⳤ�ȣ����ַ�Ϊ��λ������������ʾ��...����<br>
		��������4�����򷽷���Ϊ1������ʱ�䣬Ϊ2���������Ϊ3���ظ�����<br>
		��������5���û�ר��id��Ϊ0����ø��û�������־��<br>
		��������6����ʾ��Ϣ��1Ϊ��ʾ�������ڣ�0Ϊ����ʾ��</td>
  </tr>

  �˱�ǵ�����־�����б����Ϣ������˵�����£�
  <tr>
    <td height=23 class="tdbg">$show_class(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾϵͳ�����б�<br>
	��������1��������ʾʱ��ÿ��������Ϊ0��������ʾ��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_bloger(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ�����б�<br>
	��������1��������ʾʱ��ÿ��������Ϊ0��������ʾ��
	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_treeclass(����1)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ���͵����˵����༶�����չ���رա�<br>
	��������1��Ϊlog��ʾ������־���ࣻ<br>
	���������� Ϊuser��ʾ�����û����ࣻ<br>
	���������� Ϊphoto��ʾ������Ƭ���<br>
	���������� Ϊgroup��ʾ����Ⱥ����ࡣ</font></td>
  </tr>
     <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_template(����1)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ����ϵͳģ�塣<br>
	��������1����ʾ����Ŀ</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_placard$</td>
    <td height=23 class="tdbg">�˱����ʾϵͳ���档</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_count$</td>
    <td height=23 class="tdbg"> �˱��վ��ͳ����Ϣ��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogupdate(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ��־���������б�<br>
	��������1������������</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_userlogin$</td>
    <td height=23 class="tdbg">�˱����ʾ��¼���ڡ�</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_userlogin_l$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ������ʾ��¼���ڡ�</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_comment(����1,����2)$</td>
    <td height=23 class="tdbg"> �˱����ʾ���»ظ��б�<br>
��������1������������<br>
��������2���ظ����ⳤ�ȡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_subject(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ�û�ר�������б�<br>
	��������1������������</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_bestblog(����1)$</td>
    <td height=23 class="tdbg">�˱����ʾ�Ƽ����͡�<br>
	��������1������������</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_hotblog(����1,����2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ���Ų���,���������,�ظ���,������,���������ۺ�������<br>
    ��������1������������<br>
	��������2��Ϊ0����ʾ����ͷ��Ϊ1����ʾ����ͷ��</font>
    	</td>
  </tr>
    <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_album(����1,����2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ�û������<br>
    ��������1������������<br>
	��������2��Ϊ0��ʱ������Ϊ1���Է���������</font>
    	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_pic(����1,����2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ�û�����Ƭ��<br>
    ��������1������������<br>
	��������2��Ϊ0��ʱ������Ϊ1���Է���������Ϊ2��������������</font>
    	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_diggs(����1,����2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾDIGG��<br>
    ��������1������������<br>
	��������2��Ϊ0��ʱ������Ϊ1���Ա�DIGG���Ƽ�����Ŀ��������</font>
    	</td>
  </tr>
   <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_userdiggs(����1,����2)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾ�û���DIGG���Ƽ����Ĵ�����<br>
    ��������1������������<br>
	��������2��Ϊ0�Ա�DIGG���Ƽ�������Ŀ����Ϊ1�����û�ע��ʱ������</font>
    	</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogstar$</td>
    <td height=23 class="tdbg">�˱�ǵ������²���֮�ǡ�</td>
  </tr>
    <tr>
    <td height=23 class="tdbg">$show_blogstar2(����1,����2,����3,����4)$</td>
    <td height=23 class="tdbg">�˱�ǵ������²���֮�ǡ�<br>
	��������1��������Ŀ��<br>
	��������2��ÿ����ʾ��Ŀ��<br>
	��������3��ͼƬ��ȣ�<br>
	��������4��ͼƬ�߶ȡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newblogger(����1)$</td>
    <td height=23 class="tdbg">�˱����ʾ����ע���û���<br>
	��������1������������</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_search(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ��������<br>
	��������1��Ϊ0������ʾ��Ϊ1��������ʾ��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_cityblogger(����1)$</td>
    <td height=23 class="tdbg">�˱����ʾͬ�ǲ�����������<br>
	��������1��Ϊ0������ʾ��Ϊ1��������ʾ��<font color="#FF0000">�˱�ǩ�������ڸ�ģ�塣</font>
    </td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newphoto(����1,����2,����3,����4)$</td>
    <td height=23 class="tdbg">�˱�ǩ�������ͼƬ��<br>
	��������1������������<br>
	��������2��ÿ����ʾͼƬ����Ŀ��<br>
	��������3��ͼƬ��ȣ���λ�����أ���<br>
	��������4��ͼƬ�߶ȣ���λ�����أ���</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_teams(����1,����2,����3,����4,����5,����6)$</font></td>
    <td height=23 class="tdbg"><font color=red> �˱����ʾȺ���־��<br/>
    ��������1:�������� 1- ���´���/2-���ԾȺ��(�������)/3-��ģ��(�������)��<br/>
    ��������2:������Ŀ��<br/>
    ��������3:��Ŀ��ʾ���ȣ�<br/>
    ��������4:�Ƿ���ʾͼ�ꣻ<br/>
    ��������5:ͼ���ȣ���д��Ĭ��50���أ�<br/>
	��������6:ͼ��߶ȣ���д��Ĭ��50���ء�<br/></font>
    	</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_posts(����1,����2,����3,����4,����5)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾȺ��������־��<br/>
    ��������1:��ʾĳЩȺ�����־: 0 ����Ⱥ��;�����ѡ����Ⱥ��,���Ⱥ��ID��|�ָ���,��1|2|8��<br/>
	��������2: ������Ŀ��<br/>
	��������3:����������ʾ������<br/>
	��������4:�Ƿ���ʾ�û��� 0/1��<br/>
	��������5:�Ƿ���ʾ����ʱ�� 0/1��<br/></font>
    	</td>
  </tr>
  <tr align="left" >
    <td height=23 class="tdbg"><font color=red>$show_hottag(����1,����2,����3,����4)$</font></td>
    <td height=23 class="tdbg"> <font color=red>�˱����ʾTAG��־��<br/>
    ��������1: ������ʽ 1-�б���ʽ,2-��ͼ��ʽ��<br/>
	��������2: ��ǩ��Ŀ��<br/>
	��������3:��ȡ����<br/>
	��������4:ÿ����ʾ��Ŀ��</font>
    	</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_friends$</td>
    <td height=23 class="tdbg"> �˱����ʾ�������ӡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> �˱����ʾrss���ӱ�־��</td>
  </tr>
  <tr>
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>��ģ�壨�Գ�index.asp�������ϵͳҳ����Ч����list.asp,listblogger.asp�ļ��ȣ���</font></strong></td>
  </tr>
  <tr>
    <td height=23 colspan="2" class="tdbg">����������ģ��ı��<font color=#ff0000>($show_cityblogger(����1)$ ��ǩ����)</font>��������ͬ��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_list$</td>
    <td height=23 class="tdbg"> ��Ҫ���˱����ʾ����ϵͳ��ҳ���������壬����ȥ������ֻ���ڸ�ģ��ʹ�á�</td>
  </tr>
</table>
<br>
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 colspan="2" class="topbg"><strong>�û�ģ����˵��</strong>
  <tr>
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>��ģ�壺</font><br>
      </strong>��ģ��Ϊҳ������岿�֣�����css��ʽ���õȣ�������Dreamweave��Frontpage�б༭����ɺ󽫴���copy������

    </td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_log$</td>
    <td height=23 class="tdbg"><strong>��Ҫ���˱����ʾ��־���岿�֣��������۵���Ϣ��</strong></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_subject$</td>
    <td height=23 class="tdbg"> �˱�����б���ʽ��ʾר����ࡣ</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_subject_l$</td>
    <td height=23 class="tdbg"> �˱�Ǻ�����ʾר����ࡣ </td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_login$</td>
    <td height=23 class="tdbg"> �˱����ʾ��¼���ڡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_calendar$</td>
    <td height=23 class="tdbg"> �˱����ʾ������</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_placard$ </td>
    <td height=23 class="tdbg">�˱����ʾ�û����档</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_photo$</td>
    <td height=23 class="tdbg"><font color=red>�˱����ʾ�������������Ƭ��FLASH�������š�</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> �˱����ʾ������־�б�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> �˱����ʾ������־�б�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_comment$</td>
    <td height=23 class="tdbg"> �˱����ʾ���»ظ��б�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_newmessage$ </td>
    <td height=23 class="tdbg">�˱����ʾ���������б�</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_myfriend$</td>
    <td height=23 class="tdbg"><font color=red>�˱����ʾ�ҵĺ���������</font></td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_mygroups$</td>
    <td height=23 class="tdbg"><font color=red>�˱����ʾ�����Ⱥ�顣</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_info$</td>
    <td height=23 class="tdbg"> �˱����ʾBlog���ƣ�ͳ����Ϣ�ȡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_search$</td>
    <td height=23 class="tdbg"> �˱����ʾ��������</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_links$</td>
    <td height=23 class="tdbg"> �˱����ʾ������Ϣ��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogname$</td>
    <td height=23 class="tdbg">�˱����ʾ�û�blog���ƣ�������Ϊ������ʾ�û�id��</td>
  </tr>
  <tr align="left">
    <td height=23 class="tdbg"><font color=red>$show_blogurl$</font></td>
    <td height=23 class="tdbg"><font color=red>�˱����ʾ���͵��������ӡ�</font></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> �˱����ʾrss���ӱ�־��</td>
  </tr>
  <tr>
<td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>��ģ�壺</font><br>
      </strong>��ģ��Ϊ��ʾ��־���ݲ��֡�������־���⣬����ʱ�䣬��־���ݵ���Ϣ�İ������á�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_topic$</td>
    <td height=23 class="tdbg"> �˱����ʾ����ͼ�꣬ר�������Ƿ���ܣ��Ƿ��ö�����־���⡣</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_emot$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ�������ͼ�ꡣ</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_topictxt$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ��־���⡣</td>
  </tr>
  <tr>
    <td width="18%" height=23 class="tdbg">$show_loginfo$</td>
    <td width="82%" class="tdbg">�˱����ʾ��־����(�����ߣ�������)������ʱ�䡣</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_author$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ������(�����ߣ�������)��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_addtime$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ����ʱ�䡣</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_logtext$</td>
    <td height=23 class="tdbg"> �˱����ʾ��־���ġ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_more$</td>
    <td height=23 class="tdbg"> �˱����ʾ�Ķ�ȫ��(����)���ظ�(����)���������ӡ�</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_blogzhai$</td>
    <td height=23 class="tdbg">�˱����ʾ���뵽��ժ�����ӡ�</td>
  </tr>
</table>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
</body>
</html>