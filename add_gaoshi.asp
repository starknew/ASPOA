<html>
<head>
<title><%=Request.Cookies("homename")%></title>
<LINK href="img/Forum.css" type=text/css rel=stylesheet>
<!--#include file="conn.asp"--> <!--#include file="asp/checkuser.asp"-->
<!--#include file="asp/top.asp"-->
<%if request("action")="add" then
if request("title")="" then
response.write"<SCRIPT language=JavaScript>alert('请输入标题!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if
if request("content")="" then
response.write"<SCRIPT language=JavaScript>alert('请输入内容!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if
     set rs=server.createobject("ADODB.recordset") 
  sql = "select * from newnotice1"
 rs.Open sql,conn,1,3
 rs.addnew
 rs("title")=trim(request("title"))
 rs("content")=request("content")
 rs("username")=oabusyusername
 rs("jointime")=now() 
  rs.update
 rs.close
 Response.write "<script language='javascript'>" & chr(13)
  Response.write "alert('填加成功！');" & Chr(13)
  Response.write "window.document.location.href='admin_gaoshi1.asp';"&Chr(13)
  Response.write "</script>" & Chr(13)
        Response.End 
 end if
 %>

<body topmargin="0" leftmargin="0">

<html>
<body topmargin="5" leftmargin="0"> 
  <center>
    　<p style="line-height: 200%; margin-top: 0; margin-bottom: 0"><b>发布告示</b>&nbsp;&nbsp;&nbsp;&nbsp;"<font color=red>*</font>"号的表示必须填写 
   
	</p>
	<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">　</p>
<form method=post name="form1" action="add_gaoshi.asp?action=add">   
		

  <table width="85%" border="0" cellpadding="1" cellspacing="0" align="center" id="table4">  
    <tr> 
                <td style="font-family: 宋体; font-size: 9pt"> 
                  <table width="100%" border="0" id="table5">

                    <tr> 
                      <td style="font-family: 宋体; font-size: 9pt"> 
                        
              <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#cccccc" id="table6">
			  	<tr bgcolor="#eeeeee"> 
                  <td width="10%" align="center" style="font-family: 宋体; font-size: 9pt" height="25"> 
					<p style="margin-left: 5px; margin-right: 5px">告示标题： </td>
                  <td width="90%" style="font-family: 宋体; font-size: 9pt" height="25"> 
                    <p style="margin-left: 5px; margin-right: 5px"><input type=text name="title" size=70 style="border:1 solid #000000;background:#ffffff" value="">
                    <font color=red>*</font> </td>
                </tr>
                <tr bgcolor="#eeeeee"> 
                  <td width="10%" style="font-family: 宋体; font-size: 9pt" height="25"> 
                    <div align="center">
						<p style="margin-left: 5px; margin-right: 5px">内 容： </div>
                  </td>
                  <td width="90%" style="font-family: 宋体; font-size: 9pt" height="25"> 
                    <p style="margin-left: 5px; margin-right: 5px">
                    <input type=hidden name="content" value="">
                      <iframe ID="eWebEditor1" src="Hokilly_Editor/ewebeditor.htm?id=content&style=blue" frameborder="0" scrolling="no" width="98%" HEIGHT="350"></iframe>             
                    <font color=red>*</font> </td>
                </tr>
                  <tr bgcolor="#eeeeee"> 
                    <td align="center" colspan=2 style="font-family: 宋体; font-size: 9pt" height="25">　</td>
                  </tr>
              </table>
            </table></table>
		<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">
		<input type="submit" value="提交" name="B1" style="border:1px solid #000000;background:#EEEEEE; ">&nbsp;&nbsp;
		<input type="reset" value="重置" name="B2" style="border:1px solid #000000;background:#EEEEEE; "></p>
	</form>
	<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">　</p>
  </center>
<center>
  

</body>
</html>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>