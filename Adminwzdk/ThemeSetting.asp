<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Model_to_html.asp" -->
<!-- #include file="page_next.asp" -->

<% '搜索模块
act=request.querystring("act")
keywords=trim(request.form("keywords"))
cid=request("cid")


if act="search" then
s_sql="select * from web_theme where [name]  like '%"&keywords&"%'  order by [time] desc"
else
s_sql="select * from web_theme order by [time] desc"
end if

%>

<% '主题激活模块
action1=request.querystring("action")
ThemeFolder=request.querystring("ThemeFolder")
ThemeID=request.querystring("ThemeID")
if action1="Edit" then
set rs1=server.createobject("adodb.recordset")
sql="select web_theme,web_ThemeID from web_settings "
rs1.open(sql),cn,1,3
rs1("web_theme")=ThemeFolder
rs1("web_ThemeID")=ThemeID
rs1.update
rs1.close
set rs1=nothing

'生成该主题模板文件
set rs_create=server.createobject("adodb.recordset")
sql="select [id],ModelType,ModelTheme from web_models where  ModelTheme="&ThemeID
rs_create.open(sql),cn,1,1
Do While not rs_create.eof 
l_id=rs_create("id")
ModelType=rs_create("ModelType")
ModelTheme=rs_create("ModelTheme")
Call Model_to_html(l_id)
rs_create.movenext
loop
rs_create.close
set rs_create=nothing

'先生成首页效果
call index_to_html()

response.Write "<script language='javascript'>alert('该主题启用成功，请点击‘预览首页’查看首页效果！查看全部页面效果需要先生成其它页面！');location.href='ThemeSetting.asp';</script>"
end if
%>
<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='警告：删除后将不可恢复，可能造成意想不到后果？';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
	<%
Call header()
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>主题模板</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">

          <tr>

        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td>
<div class='ThemeArea'>
<% '文章列表模块
strFileName="ad_list.asp" 
pageno=10
set rs = server.CreateObject("adodb.recordset")
rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>
<% if p_i mod 2 =0 then
class_style="forumRow"
else
class_style="forumRowHighLight"
end if%>
<div class='ThemeBlock'>
<div class='preview'><a href='<%=rs("memo")%>' target='_blank'><%=rs("Folder")%></a></div>
<div class='inner'>
<div class='img'><a href='<%=rs("memo")%>' target='_blank'><img src="/images/up_images/<%=rs("image")%>" width="200" height="225" border="0" alt="<%=rs("name")%>"></a>
<p><a href='<%=rs("memo")%>' target='_blank'><img src="images/view_icon.gif"  border="0"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a></p></div>
</div>
</div>
<%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>暂无主题！</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>

<div class="clearfix"></div>

</div>
</td>
</tr>
		    <tr  >
              <td height="35"  colspan="9" ><div align="center">
           </div></td>
		    </tr>
		    <tr  >
              <td height="35"  colspan="9" ></td>
		    </tr>
</table>
	    

              </div>
            </form>
            </td>
          </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>