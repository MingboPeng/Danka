<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!-- #include file="../inc/AntiAttack.asp" -->
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/web_config.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<%
search_q=request.querystring("q")
%>
<title>������<%=search_q%>_�����д󿵻������޹�˾</title>
<meta name="keywords" content="$Class_Keywords$" />
<meta name="description" content="$Class_Description$" />
<link href="/css/HituxBlue/inner.css" rel="stylesheet" type="text/css" />
<link href="/css/HituxBlue/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/functions.js"></script>
<script type="text/javascript" src="/images/iepng/iepngfix_tilebg.js"></script>
</head>

<body>
<%
keywords=split(search_q," ")
c=ubound(keywords)
for i=0 to c
if i=0 then
search_sql1=search_sql1&"where  ( [title] like '%"&keywords(i)&"%'"
keywords_all=keywords(i)
else
search_sql1=search_sql1&" or   [title] like '%"&keywords(i)&"%'"
keywords_all=keywords_all&"+"&keywords(i)
end if
next

s_sql="select [title],[content],[file_path],[time],ArticleType from [article] "&search_sql1&" )  and view_yes=1 order by [time] desc"
%>
<div id="wrapper">

<!--head start-->
<div id="head">

<!--top start -->
<div class="top">
<div class="TopInfo"><div class="link">ѡ�����ԣ�<a  href="javascript:zh_tran('s');" class="zh_click" id="zh_click_s">��������</a> | <a href="javascript:zh_tran('t');" class="zh_click" id="zh_click_t">���w����</a> </div>
</div>
<div class="clearfix"></div>
<div class="TopLogo">
<div class="logo"><a href="/"><img src="/images/up_images/201382164659.png" alt="�����д󿵻������޹�˾"></a></div>
</div>

</div>
<!--top end-->

<!--nav start-->
<div id="NavLink">
<div class="NavBG">
<!--Head Menu Start-->
<ul id='sddm'><li class='CurrentLi'><a href='/'>��վ��ҳ<p>home</p></a></li> <li><a href='/About/' onmouseover=mopen('m2') onmouseout='mclosetime()'>���ڹ�˾<p>about</p></a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/'>��˾����</a> <a href='/About/gongsitupian/'>��˾ͼƬ</a> <a href='/About/Groups/'>���Ŷ�</a> <a href='/About/Honour/'>������</a> </div></li> <li><a href='/Product/' onmouseover=mopen('m3') onmouseout='mclosetime()'>��Ʒչʾ<p>product</p></a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Product/Welding'>����</a> <a href='/Product/Welder'>����</a> <a href='/Product/Hardware'>���</a> <a href='/Product/Pump'>ˮ��</a> <a href='/Product/Mechanicals'>��ҵ����</a> <a href='/Product/Air Compressor/'>��ѹ��</a> <a href='/Product/Cutting/'>�и�Ƭ</a> <a href='/Product/Scales/'>����</a> </div></li> <li><a href='/news/' onmouseover=mopen('m4') onmouseout='mclosetime()'>������Ѷ<p>news</p></a> <div id='m4' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/news/CompanyNews'>��˾����</a> <a href='/news/IndustryNews'>��ҵ����</a> </div></li> <li><a href='/Support' onmouseover=mopen('m5') onmouseout='mclosetime()'>����֧��<p>Support</p></a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Support/Services'>�ۺ����</a> <a href='/Support/Download'>��������</a> </div></li> <li><a href='/Recruit' onmouseover=mopen('m6') onmouseout='mclosetime()'>�˲���Ƹ<p>Recruit</p></a> <div id='m6' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/jobs'>��Ƹְλ</a> </div></li> <li><a href='/contact/'>��ϵ��ʽ<p>contact</p></a></li> <li><a href='/Feedback/'>�ÿ�����<p>feedback</p></a></li> </ul>
<!--Head Menu End-->
</div>
<div class="clearfix"></div>
</div>
<!--nav end-->

</div>
<!--head end-->
<!--body start-->
<div id="body">
<!--focus start-->
<div id="InnerBanner">

</div>
<!--foncus end-->
<div class="HeightTab clearfix"></div>
<!--inner start -->
<div class="inner">
<!--left start-->
<div class="left">
<div class="Sbox">
<div class="topic">��ϵ����&nbsp;&nbsp;&nbsp;Contact</div>
<div class="txt ColorLink">
<p>��ַ���㽭ʡ���������������нֵ������Է��������10-14��</p>
<p>�绰��0577-86860303 86866000</p>
<p>���棺0577-86885103</p>
<p>��վ��<a href='http://www.wzdk86860303.com' target='_blank'>www.wzdk86860303.com</a> </p>
<p align='center'><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=314237795&site=qq&menu=yes"><img border="0" src="http://pub.idqqimg.com/wpa/images/counseling_style_52.png" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ"></a> <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2621423199&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2621423199:42" alt="���������ҷ���Ϣ" title="���������ҷ���Ϣ"></a>   </p></div>
</div>
<div class="HeightTab clearfix"></div>

<div class="Sbox">
<div class="topic">����&nbsp;&nbsp;&nbsp;Search</div>
<div class="SearchBar">
<form method="get" action="/Search/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='������ؼ���';" 
onfocus="if(this.value=='������ؼ���') this.value='';" value="������ؼ���" /><input type="submit" id="search-submit" value="����" />
			</form>
</div>
</div>

</div>
<!--left end-->
<!--right start-->
<div class="right">
<div class="Position"><span>���λ�ã�<a href="/">��ҳ</a> > ����</span></div>
<div class="HeightTab clearfix"></div>
<!--main start-->
<div class="main">

<!--search content start-->
<div id="search_content" class="clearfix">

<%
if search_q<>"" then 

set rs=server.createobject("adodb.recordset")
rs.open(s_sql),cn,1,1
%>

<%'=============��ҳ���忪ʼ��Ҫ�������ݿ��֮��
if err.number<>0 then '������
response.write "���ݿ����ʧ�ܣ�" & err.description
err.clear
else
if not (rs.eof and rs.bof) then '����¼���Ƿ�Ϊ��
r=cint(rs.RecordCount) '��¼����
rowcount = 10 '����ÿһҳ�����ݼ�¼�����ɸ���ʵ���Զ���
rs.pagesize = rowcount '��ҳ��¼��ÿҳ��ʾ��¼��
maxpagecount=rs.pagecount '��ҳҳ��
page=request.querystring("page")
  if page="" then
  page=1
  end if
rs.absolutepage = page 
rcount1=0
pagestart=page-5
pageend=page+5
if pagestart<1 then
pagestart=1
end if
if pageend>maxpagecount then
pageend=maxpagecount
end if
rcount=rs.RecordCount
'=============��ҳ�������%>

<!--position start-->
<div class="searchtip">������������<span class="FontRed"><%=search_q%></span>��,�ҵ������Ϣ <span class="font_brown"><%=rcount%></span> ��</div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">��ʾ���ÿո���������Ѱ�ؼ��ʿɻ�ȡ�����������硰���� ��Ʒ����</div>
<dl>

<%'===========ѭ���忪ʼ
do while not rs.eof and rowcount%>
<%
select case rs("ArticleType")
case 1
Content_FolderName=Article_FolderName
case 2
Content_FolderName=Product_FolderName
end select

title1=left(rs("title"),30)
for i=0 to c
title1=Replace(title1, keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next

content1=left(nohtml(rs("content")),110)
for i=0 to c
content1=Replace(content1,keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next
%>
<dt ><a href='<%="/"&Content_FolderName&"/"&rs("file_path")%>' target='_blank' title='<%=rs("title")%>'><%=title1%></a></dt>
<dd><%=content1%>...</dd>
<dd class="font12 arial font_green line"><a href='<%="/"&Content_FolderName&"/"&rs("file_path")%>' target='_blank'><span class="font_green"><%=web_url&"/"&Content_FolderName&"/"&rs("file_path")%></span></a><%=year(rs("time"))%>-<%=month(rs("time"))%>-<%=day(rs("time"))%></dd>
<%
rowcount=rowcount-1 
rs.movenext
loop
 '===========ѭ�������%>

</dl>
</div>
<!--list end-->

<!--page start-->
<div class="result_page clearfix">
<!--#include file="../inc/page_list.asp"-->
</div>
<!--page end-->

<%
else
response.write "<div class='search_welcome'>�ܱ�Ǹ,û���ҵ��� <span class='FontRed'>"&search_q&"</span> ��ص���Ϣ��<p >��ʾ���ÿո���������Ѱ�ؼ��ʿɻ�ȡ�����������硰���� ��Ʒ����</p></div>"
end if
end if
end if%>
</div>
<!--search content end-->	

</div>
<!--main end-->
</div>
<!--right end-->
</div>
<!--inner end-->
</div>
<!--body end-->
<div class="HeightTab clearfix"></div>
<!--footer start-->
<div id="footer">
<div class="inner">
<div class='BottomNav'><a href="/">��վ��ҳ</a> | <a href="/About">��������</a> | <a href="/Recruit">�˲���Ƹ</a>  | <a href="/Sitemap">��վ��ͼ</a> | <a href="/RSS">����RSS</a></div>
<div class='HeightTab'></div>
<p>Copyright 2013 <a href='http://www.wzdk86860303.com' target='_blank'>www.wzdk86860303.com</a> �����д󿵻������޹�˾ ��Ȩ���� All Rights Reserved </p>
<p>��˾��ַ�����������������нֵ������Է��������10-14 ��ϵ�绰��0577-86860303 </p>
<p>Built By <a href="http://www.wzdk86860303.com/" target="_blank">Mingbo</a> <a href="http://www.mingbopeng.com/" target="_blank">������</a> ����֧�� <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a></p>

<script type="text/javascript" src="http://zjnet.zjaic.gov.cn/wzqybswj/3303030001009043.js"></script>
<a target="_blank" href="http://idinfo.zjaic.gov.cn/bscx.do?method=hddoc&amp;id=3303030001009043"><img src="http://idinfo.zjaic.gov.cn/images/i_lo2.gif" border="0"></a>
</div>
</div>
<!--footer end -->


</div>
<script type="text/javascript">
window.onerror=function(){return true;}
</script>
</body>
</html>
<!--
Powered By huiguerCMS ASP V2.O   
-->





