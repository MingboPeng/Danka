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
<title>搜索：<%=search_q%>_温州市大康机电有限公司</title>
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
<div class="TopInfo"><div class="link">选择语言：<a  href="javascript:zh_tran('s');" class="zh_click" id="zh_click_s">简体中文</a> | <a href="javascript:zh_tran('t');" class="zh_click" id="zh_click_t">繁體中文</a> </div>
</div>
<div class="clearfix"></div>
<div class="TopLogo">
<div class="logo"><a href="/"><img src="/images/up_images/201382164659.png" alt="温州市大康机电有限公司"></a></div>
</div>

</div>
<!--top end-->

<!--nav start-->
<div id="NavLink">
<div class="NavBG">
<!--Head Menu Start-->
<ul id='sddm'><li class='CurrentLi'><a href='/'>网站首页<p>home</p></a></li> <li><a href='/About/' onmouseover=mopen('m2') onmouseout='mclosetime()'>关于公司<p>about</p></a> <div id='m2' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/About/'>公司介绍</a> <a href='/About/gongsitupian/'>公司图片</a> <a href='/About/Groups/'>大康团队</a> <a href='/About/Honour/'>大康荣誉</a> </div></li> <li><a href='/Product/' onmouseover=mopen('m3') onmouseout='mclosetime()'>商品展示<p>product</p></a> <div id='m3' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Product/Welding'>焊材</a> <a href='/Product/Welder'>焊机</a> <a href='/Product/Hardware'>五金</a> <a href='/Product/Pump'>水泵</a> <a href='/Product/Mechanicals'>工业机电</a> <a href='/Product/Air Compressor/'>空压机</a> <a href='/Product/Cutting/'>切割片</a> <a href='/Product/Scales/'>衡器</a> </div></li> <li><a href='/news/' onmouseover=mopen('m4') onmouseout='mclosetime()'>新闻资讯<p>news</p></a> <div id='m4' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/news/CompanyNews'>公司新闻</a> <a href='/news/IndustryNews'>行业新闻</a> </div></li> <li><a href='/Support' onmouseover=mopen('m5') onmouseout='mclosetime()'>技术支持<p>Support</p></a> <div id='m5' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/Support/Services'>售后服务</a> <a href='/Support/Download'>下载中心</a> </div></li> <li><a href='/Recruit' onmouseover=mopen('m6') onmouseout='mclosetime()'>人才招聘<p>Recruit</p></a> <div id='m6' onmouseover='mcancelclosetime()' onmouseout='mclosetime()'><a href='/recruit/jobs'>招聘职位</a> </div></li> <li><a href='/contact/'>联系方式<p>contact</p></a></li> <li><a href='/Feedback/'>访客留言<p>feedback</p></a></li> </ul>
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
<div class="topic">联系我们&nbsp;&nbsp;&nbsp;Contact</div>
<div class="txt ColorLink">
<p>地址：浙江省温州市龙湾区永中街道南洋锦苑丁香三幢10-14号</p>
<p>电话：0577-86860303 86866000</p>
<p>传真：0577-86885103</p>
<p>网站：<a href='http://www.wzdk86860303.com' target='_blank'>www.wzdk86860303.com</a> </p>
<p align='center'><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=314237795&site=qq&menu=yes"><img border="0" src="http://pub.idqqimg.com/wpa/images/counseling_style_52.png" alt="点击这里给我发消息" title="点击这里给我发消息"></a> <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=2621423199&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:2621423199:42" alt="点击这里给我发消息" title="点击这里给我发消息"></a>   </p></div>
</div>
<div class="HeightTab clearfix"></div>

<div class="Sbox">
<div class="topic">搜索&nbsp;&nbsp;&nbsp;Search</div>
<div class="SearchBar">
<form method="get" action="/Search/index.asp">
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='请输入关键词';" 
onfocus="if(this.value=='请输入关键词') this.value='';" value="请输入关键词" /><input type="submit" id="search-submit" value="搜索" />
			</form>
</div>
</div>

</div>
<!--left end-->
<!--right start-->
<div class="right">
<div class="Position"><span>你的位置：<a href="/">首页</a> > 搜索</span></div>
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

<%'=============分页定义开始，要放在数据库打开之后
if err.number<>0 then '错误处理
response.write "数据库操作失败：" & err.description
err.clear
else
if not (rs.eof and rs.bof) then '检测记录集是否为空
r=cint(rs.RecordCount) '记录总数
rowcount = 10 '设置每一页的数据记录数，可根据实际自定义
rs.pagesize = rowcount '分页记录集每页显示记录数
maxpagecount=rs.pagecount '分页页数
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
'=============分页定义结束%>

<!--position start-->
<div class="searchtip">您正在搜索“<span class="FontRed"><%=search_q%></span>”,找到相关信息 <span class="font_brown"><%=rcount%></span> 条</div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">提示：用空格隔开多个搜寻关键词可获取更理想结果，如“最新 产品”。</div>
<dl>

<%'===========循环体开始
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
 '===========循环体结束%>

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
response.write "<div class='search_welcome'>很抱歉,没有找到与 <span class='FontRed'>"&search_q&"</span> 相关的信息！<p >提示：用空格隔开多个搜寻关键词可获取更理想结果，如“最新 产品”。</p></div>"
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
<div class='BottomNav'><a href="/">网站首页</a> | <a href="/About">关于我们</a> | <a href="/Recruit">人才招聘</a>  | <a href="/Sitemap">网站地图</a> | <a href="/RSS">订阅RSS</a></div>
<div class='HeightTab'></div>
<p>Copyright 2013 <a href='http://www.wzdk86860303.com' target='_blank'>www.wzdk86860303.com</a> 温州市大康机电有限公司 版权所有 All Rights Reserved </p>
<p>公司地址：温州市龙湾区永中街道南洋锦苑丁香三幢10-14 联系电话：0577-86860303 </p>
<p>Built By <a href="http://www.wzdk86860303.com/" target="_blank">Mingbo</a> <a href="http://www.mingbopeng.com/" target="_blank">彭铭博</a> 技术支持 <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a></p>

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






