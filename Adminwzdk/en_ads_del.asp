<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/index_to_html.asp" -->

	<%
Call header()
%>


	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>删除广告</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" bgcolor="#B1CFF8"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="100">
			<%page=request.querystring("page")
			act=request.querystring("act")
			keywords=request.querystring("keywords")
			article_id=cint(request.querystring("id"))
			set rs=server.createobject("adodb.recordset")
sql="select id from en_web_ads where id="&article_id&""
rs.open(sql),cn,1,3
rs.delete
rs.close
set rs=nothing
'先判断文件是否存在，否则删除
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath("/ADs/"&article_id&".js")) then
FilePath="/ADs/"&article_id&".js"
call DelFile(FilePath)
end if
call index_to_html()
response.Write "<script language='javascript'>alert('删除成功！');location.href='en_ads_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>