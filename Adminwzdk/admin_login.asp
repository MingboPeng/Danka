

<%
Dim errmsg

username=Request("username")
password=Request("password")
verifycode=Request("verifycode")

If username="" Or password="" Then 
	Response.Redirect "login.asp?errno=2"
	Response.End 
End If 


%>
<!-- #include file="Inc/md5.asp" -->
<!-- #include file="Inc/Functions.asp" -->
<%
username=getSafeStr(username)
password=getSafeStr(password)
verifycode=getSafeStr(verifycode)
password_m=md5(password,16)
sql="select id,username,password,class from Web_admin where username='"&username&"' and password='"&password_m&"'"
response.Write sql
Set rs=cn.execute(sql)
If Not (rs.bof Or rs.eof) Then 
	Session("log_name")=rs("username")
			Session("log_role")=rs("class")
			session.Timeout=1000
			Response.Redirect "index.asp"
Else 
	Response.Redirect "login.asp?errno=1"
End If 

Call DBconnEnd()
%>