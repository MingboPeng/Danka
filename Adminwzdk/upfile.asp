<!--#include file="../inc/access.asp"-->
</style><!--#include file="upload.inc"-->
<%  
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''�����ϴ�����
formPath="/images/up_images"
 ''��Ŀ¼���(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
response.Write "<script language='javascript'>alert('��ѡ����Ҫ�ϴ���ͼƬ��');location.href='upload.asp';</script>"
	response.end
 end if
 	
 if file.filesize>1000000 then '�ϴ�ͼƬ��С���ƣ�100000Ϊ100KB�����Լ������󡭡���������
response.Write "<script language='javascript'>alert('ͼƬ��С���������ƣ��������ϴ���');location.href='upload.asp';</script>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" and fileEXT<>".bmp" and fileEXT<>".png" then
response.Write "<script language='javascript'>alert('�ļ���ʽ���ԣ��������ϴ���');location.href='upload.asp';</script>"
	response.end
 end if 
 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
 file.SaveAs Server.mappath(filename)   ''�����ļ�
 'response.write "<script>parent.form1.textfield6.value='"&FileName1&"'</'script>"
 response.write  fielname1
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���
'response.write "<script>parent.form1.textfield6.value="&Filename1&"<'/script>"
session("upface")="done"
sub HtmEnd(Msg)
 set upload=nothing
 response.write "ͼƬ�ϴ��ɹ�"
 response.end
end sub
%>
</body>
</html>
<%response.redirect("upload.asp?FileName="&Filename1&"")%>