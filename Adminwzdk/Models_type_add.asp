<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<%'�ж�ģ�����Ȩ���Ƿ���
set  rs_a=server.createobject("adodb.recordset")
sql="select web_ModelEdit from web_settings"
rs_a.open(sql),cn,1,1
if rs_a("web_ModelEdit")=0 then
response.Write "<script language='javascript'>alert('����ģ�����Ȩ����δ������');history.go(-1);</script>"
else
%>
<% '�������ݵ����ݱ�
act=Request("act")
If act="save" Then 
l_name=trim(request.form("name"))
l_url=replace(LCase(trim(request.form("url"))),".asp",".html")
FolderName=trim(request.form("FolderName"))
l_image=trim(request.form("web_image"))
l_memo=trim(request.form("memo"))
l_view_yes=trim(request.form("view_yes"))
l_time=now()

'����ļ������Ƿ���ϵͳ�ļ������ظ�
			set rs_f=server.createobject("adodb.recordset")
			sql="select [name] from web_SystemFolder where [name]='"&FolderName&"'"
			rs_f.open(sql),cn,1,1
if not rs_f.eof then
response.Write "<script language='javascript'>alert('�����ļ���������ϵͳ�ļ������ظ���������������');history.go(-1);</script>"
else
set rs=server.createobject("adodb.recordset")
sql="select * from web_models_type where  FileName='"&l_url&"'"
rs.open(sql),cn,1,3
if not rs.eof then
response.Write "<script language='javascript'>alert('ģ���ļ����ظ���������������');history.go(-1);</script>"
else
rs.addnew
rs("name")=l_name
rs("FileName")=l_url
rs("FolderName")=FolderName
rs("image")=l_image
rs("memo")=l_memo
rs("view_yes")=cint(l_view_yes)
rs("time")=l_time
rs.update
rs.close
set rs=nothing
%>
<% '�ж��ļ����Ƿ���ڣ����򴴽�
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/"&FolderName))=false Then
NewFolderDir="/"&FolderName
call CreateFolderB(NewFolderDir)
end if
%>
<%
response.Write "<script language='javascript'>alert('���ӳɹ���');location.href='Models_Type_List.asp';</script>"
end if 
end if
end if%>
 

	<%
Call header()

%>

  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.name.value == '' ) {
window.alert('������ģ���������^_^');
document.form1.name.focus();
return false;}

if ( document.form1.url.value == '' ) {
window.alert('������ģ���ļ���^_^');
document.form1.url.focus();
return false;}

return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>����ģ�����</th>
	<tr>
	<td width="15%" height=23 class='forumRow'>�������� (����) </td>
	<td class='forumRow'><input name='name' type='text' id='name' size='40' maxlength="70">
	  &nbsp;</td>
	</tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>ģ���ļ��� (����) </td>
	    <td class='forumRowHighLight'><input name='url' type='text' id='url' size='40' maxlength="70">
          <span style="color: #FF0000">&nbsp;���磺index_model.html�����������������������޸ģ�</span></td>
      </tr>
	  <tr>
	    <td class='forumRow' height=23>Ŀ���ļ��� (����) </td>
	    <td class='forumRow'><input name='FolderName' type='text' id='FolderName' size='40' maxlength="70">
          <span style="color: #FF0000">&nbsp;���磺Cool���ļ����ڸ�Ŀ¼�£�����������������ϵͳ�ļ������ظ������������޸ģ�</span></td>
      </tr>	  
	  <tr>
	    <td class='forumRowHighLight' height=23>ͼƬ</td>
	    <td width="85%" class='forumRowHighLight'><table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td width="22%" ><input name="web_image" type="text" id="web_image"  size="30"></td>
           <td width="78%"  ><iframe width="500" name="ad" frameborder=0 height=30 scrolling=no src=upload.asp></iframe></td>
         </tr>
       </table></td>
      </tr><tr>
	  <td class='forumRow' height=11>��ע</td>
	  <td class='forumRow'><textarea name='memo'  cols="100" rows="6" id="memo" ></textarea></td>
	</tr>
	  
	  <tr>
	  <td class='forumRowHighLight' height=23>�Ƿ����</td>
	  <td class='forumRowHighLight'><label>
	    <input name="view_yes" type="radio" value="1" checked="checked">
      ��
      &nbsp;
      <input name="view_yes" type="radio" value="0">
      ��</label></td>
	</tr>
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='�ύ' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
end if
rs_a.close
set rs_a=nothing
%>
<%
Call DbconnEnd()
 %>