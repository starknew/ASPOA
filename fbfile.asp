
<html>
<head><httpRuntime   executionTimeout="90"   maxRequestLength="100000"   useFullyQualifiedRedirectUrl="false"   />
<title>∷电子公文传输系统∷</title>
<LINK href="css/global.css" rel=stylesheet type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
.comments {
 width:95%;/*自动适应父布局宽度*/
 overflow:auto;
 word-break:break-all;
}
</style>

<!--#include file="conn.asp"-->
<!--#include file="asp/checkuser.asp"-->
<SCRIPT language=javascript src="css/init.js">
</SCRIPT>
<% if instr(1,"||supermanage||chksuper||chkmeet||chfiles|",oabusyuserpower)<2 then%>
<%Response.write "<script language='javascript'>" & chr(13)
		Response.write "alert('您没有这个权限！');" & Chr(13)
		Response.write "window.document.location.href='document.asp';"&Chr(13)
		Response.write "</script>" & Chr(13)
        Response.End%>
<%end if%>
<%
if Request("action") = "add" Then
	if request("joinman")="" then
response.write"<SCRIPT language=JavaScript>alert('接收人员不能为空!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if
if request("title")="" then
response.write"<SCRIPT language=JavaScript>alert('请输入文件标题!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if
if request("fbdw")="" then
response.write"<SCRIPT language=JavaScript>alert('请输入发文人员!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if
if request("filename")="" then
response.write"<SCRIPT language=JavaScript>alert('文件类型不能没有!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if
if request("content")="" then
response.write"<SCRIPT language=JavaScript>alert('文件内容不能为空!!');"                                      
response.write"javascript:history.go(-1)</SCRIPT>"                                      
response.end
end if


	function makefilename(fname)
		fname = now()
		fname = replace(fname,"-","")
		fname = replace(fname," ","") 
		fname = replace(fname,":","")
		fname = replace(fname,"PM","")
		fname = replace(fname,"AM","")
		fname = replace(fname,"上午","")
		fname = replace(fname,"下午","")
		makefilename=fname
	end function 
	
	vfname = makefilename(now())
	
	set rs=server.createobject("ADODB.recordset") 
	sql="select * from path"
	rs.open sql,conn,1,1
	
	if not rs.eof and not rs.bof then
		filepath=rs("fileword")
	else
		filepath="upfile/"
	end if
	rs.close
   df=request("joinman")
   joinman = trim(request("joinman"))
   title=request("title")
   fbdw=request("fbdw")
   filename=request("filename")
   fk=request("fk")
   file1=request("file1")
   wenhao=request("wenhao")
   content=request("content")

	set rs=server.createobject("ADODB.recordset") 
	sql = "select * from filedata"
	rs.Open sql,conn,2,3
	rs.addnew 
	 if instr(1,"||supermanage|chksuper|",oabusyuserpower)>1   then
	    if dfdw<>"" then
			rs("fbdw")=df
			rs("manageman")=dfdw
		else
			rs("fbdw")=fbdw
			rs("manageman")=oabusyusername
		end if
	else
		rs("fbdw")=fbdw
		rs("manageman")=oabusyusername
	end if
	
	rs("title")=trim(title)
	rs("typename")=filename
	if instr(joinman,rs("manageman")) then 
		joinman=joinman&"|"
	else
		joinman="|"&rs("manageman")&"|"&joinman&"|"
	end if
	rs("joinman")=joinman
	rs("fk")=fk
	rs("content")=content
        rs("jointime")=now()
	rs("filepath")=filepath
	rs("wenhao")=wenhao
	rs("filename")=file1
	rs.update 
rs.close
sql = "select top 1 * from filedata order by id desc"
	rs.Open sql,conn,1,1
	id1=rs("id")
        title1=rs("title")
rs.close
	
	if filename1<>"" then xup.SaveFile "file1",filepath&vfname&filename1

	set xup=nothing
set rs2=server.createobject("ADODB.recordset") 
	sql2="select * from fileuser"
	rs2.Open sql2,conn,1,3
	mysendto=split(joinman,"|",-1,1)
	
	for each sendtoinf in mysendto
		if trim(sendtoinf)<>"" then
			set rs1=server.createobject("ADODB.recordset") 
			sql="select * from userinf  where username='"&trim(sendtoinf)&"'"
			rs1.Open sql,conn,1,3
			rs2.addnew
			rs2("username")=sendtoinf
			if not rs1.EOF then  rs2("name")=rs1("name")
			if sendtoinf=oabusyusername or rs1("userlevel")="超级管理员" or rs1("userlevel")="最高管理员"  or rs1("userlevel")="高级发文者"   then
                        if sendtoinf=oabusyusername then
                            rs2("ip")="传阅者"
                        else
                            rs2("ip")="管理人员"
                        end if
                            rs2("qshou")=1
                            rs2("qstime")=now()+0.001
else
  rs2("qshou")=0
  rs2("qstime")=null
end if
			rs2("reid")=id1
			rs2("title")=title
                        rs2("gzgw")=rs1("gzgw")
			rs2.update
			rs1.Close
			set rs1=nothing
		end if
	next
	rs2.close
	set rs2=nothing
	'写入反馈
	if fk=1 then
	set rs3=server.createobject("ADODB.recordset") 
	sql2="select * from filefk"
	rs3.Open sql2,conn,1,3
	mysendto=split(joinman,"|",-1,1)
	
	for each sendtoinf in mysendto
		if trim(sendtoinf)<>"" then
			set rs1=server.createobject("ADODB.recordset") 
			sql="select * from userinf  where username='"&trim(sendtoinf)&"'"
			rs1.Open sql,conn,1,3
			rs3.addnew
			rs3("username")=trim(sendtoinf)
			if not rs1.EOF then  rs3("name")=rs1("name")
			if sendtoinf=oabusyusername or rs1("userlevel")="超级管理员" or rs1("userlevel")="最高管理员" or rs1("userlevel")="最高管理员"  then
                        rs3("fk")=1
                        rs3("ip")="管理人员"
                        rs3("fkdate")=now()+0.001
else
rs3("fk")=0
rs3("fkdate")=null
end if

			rs3("reid")=id1
			rs3("title")=title
			rs3.update
			rs1.Close
			set rs1=nothing
		end if
	next
	rs3.close
	set rs3=nothing
end if
	set conn=nothing 

%>
<style>
<!--
input        { border-left: 1px solid #FFFFFF; border-right: 1px solid #FFFFFF; 
               border-top: 1px solid #FFFFFF; border-bottom: 1px solid #C0C0C0 }
-->
</style><%Response.write "<script language='javascript'>" & chr(13)
		Response.write "alert('文件发布成功！');" & Chr(13)
		Response.write "window.document.location.href='document.asp';"&Chr(13)
		Response.write "</script>" & Chr(13)
        Response.End%>
<%
else
%>
  <center>
    <form method=post name="form1" action="fbfile.asp?action=add">   
		

                   
		

  <div align="center">
		

                   
		

  <table cellSpacing="0" cellPadding="0" width="95%" align="center" bgColor="#ffffff" border="0" id="table12">
					<tr>
						<td><img src="images/lban.gif"></td>
						<td width="100%" background="images/bg.gif">
						<table cellSpacing="0" cellPadding="0" width="100%" border="0" id="table13">
							<tr>
								<td width="100%">　</td>
								<td>
								<input onClick="form_check()" type="image" src="images/rb_save_up.png" style="border: 1px solid #004477; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" name="sub0"> 
								</td>
								<td>　</td>
								<td>
								<a href='javascript:history.back();' onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','images/rb_back_up.png',1)"><img src='images/rb_back_up.png' name='Image9' border='0'></a> 
								</td>
								<td>　</td>
							</tr>
						</table>
						</td>
						<td><img src="images/rban.gif"></td>
					</tr>
					<tr bgColor="#00bfef">
						<td>
						<img height="16" src="images/ljiao.gif" width="19"></td>
						<td class="font9pt" align="right">　</td>
						<td align="right">
						<img height="16" src="images/rjiao.gif" width="19"></td>
					</tr>
				</table>
		

                   
		

  <table width="98%" border="0" cellspacing="0" id="table7" bgcolor="#E8E8E8">  
    <tr> 
                <td style="font-family: 宋体; font-size: 9pt"> 
                  <table width="100%" border="0" id="table8">

                    <tr> 
                      <td style="font-family: 宋体; font-size: 9pt"> 
                        
              <div align="center">
				<p style="line-height: 200%; margin-top: 0; margin-bottom: 0">　</p>
                        
              <table width="90%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FF0000" id="table9">
			  	<tr>
              <td style="font-family: 黑体; font-size: 9pt" height="50" bgcolor="#FFFFFF" colspan="4" align="center"> 
              <b><font size="5" color="#FF0000">发 布 新 文 电</font></b></td>
            	</tr>
			  	<tr>
              <td width="18%" style="font-family: 宋体; font-size: 9pt" height="30" bgcolor="#FFFFFF"> 
              <div align="center"><b><font color="#000000">接收人员：</font></b></div>              </td>
              <td align="left" valign="top" style="font-family: 宋体; font-size: 9pt" height="50" bgcolor="#FFFFFF" colspan="3"> 
                <p style="line-height: 150%; margin-top: 0; margin-bottom: 0"> 
				<a href="USERLIST.ASP?opper=fbfile.asp" style="color: #0000FF; text-decoration: none; font-family: 宋体; font-size: 11pt"><b><u>
				点击选择接收人员:</b></u></a></p>
				<p style="line-height: 150%; margin-top: 0; margin-bottom: 0">
				<textarea class="comments" name="joinman"  cols="100" id="joinman" readonly="      " rows="3" style="font-family:宋体;font-size:11pt;border-left:1px solid #FFFFFF; border-right:1px solid #FFFFFF; border-top:1px solid #FFFFFF; border-bottom:1px solid #C0C0C0; background:#FFFFFF; "><%=trim(request("joinman"))%></textarea>              <font color=red>*</font></td>
            	</tr>
				<tr>
              <td width="18%" style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                <div align="center"><b><font color="#000000">文电标题：</font></b></div>
              </td>
              <td style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                &nbsp;<input type=text name="title" size=61 style="border-left:1px solid #FFFFFF; border-right:1px solid #FFFFFF; border-top:1px solid #FFFFFF; border-bottom:1px solid #C0C0C0; background:#FFFFFF; "><font color="#FF0000">*
				</font> </td>
              <td width="305" style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                <p align="center"><b><font color="#000000">需要反馈：</font></b></td>
              <td width="305" style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                <input type="radio" value="0" checked name="fk">不需要&nbsp; 
				<input type="radio" name="fk" value="1">需要&nbsp;</td>
            	</tr>
				<tr>
              <td width="18%" style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                <p align="center"><b><font color="#000000">文电号：</font></b></td>
              <td style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                
                <input type=text name="wenhao" style="border-left:1px solid #FFFFFF; border-right:1px solid #FFFFFF; border-top:1px solid #FFFFFF; border-bottom:1px solid #C0C0C0; background:#FFFFFF; " size=29 style="background:#FFFFFF; " value="﹝<%=year(date)%>﹞号"></td>
              <td width="10%" style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                
                <p align="center"><b><font color="#000000">传阅日期：</font></b></td>
              <td width="26%" style="font-family: 宋体; font-size: 9pt" height="40" bgcolor="#FFFFFF"> 
                
                　<%=year(date)%>年<%=month(date)%>月<%=day(date)%>日</td>
            	</tr>
				<tr>
				
				
              <td width="18%" height="40" style="font-family: 宋体; font-size: 9pt" bgcolor="#FFFFFF"> 
              <div align="center"><b><font color="#000000">传阅人：</font></b></div>              </td>
              <td height="40" style="font-family: 宋体; font-size: 9pt" bgcolor="#FFFFFF" width="41%"> 
                &nbsp;<input type="text" style="border-left:1px solid #FFFFFF; border-right:1px solid #FFFFFF; border-top:1px solid #FFFFFF; border-bottom:1px solid #C0C0C0; background:#FFFFFF; " name="fbdw" size="60" value="<%=oabusyname%>" style="background:#FFFFFF; ">
                <font color=red>*</font> 
              </td>
              <td height="40" style="font-family: 宋体; font-size: 9pt" bgcolor="#FFFFFF" width="10%"> 
                <p align="center"><b><font color="#000000">文件类型：</font></b></td>
              <td height="40" style="font-family: 宋体; font-size: 9pt" bgcolor="#FFFFFF" width="26%"> 
                
                <div align="left">&nbsp;<select name="filename" size="1" style="border:1 solid #000000;background:#ffffff; font-family:宋体; font-size:9pt">
                    <%Set rs=Server.CreateObject("ADODB.recordset")
 		sql="select * from fbgtype order by id asc"
		rs.open sql,conn,1
		if not rs.bof and not rs.eof then
		do while not rs.eof %>
        <option value="<%=rs("typename")%>"><%=rs("typename")%></option>
                    <%rs.movenext
                    loop
                    end if
                    rs.close%>
                  </select>
                </div> </td>
            	</tr>
				<tr>
              <td width="18%" style="font-family: 宋体; font-size: 9pt" height="30" bgcolor="#FFFFFF"> 
                <p align="center"><b><font color="#000000">领导要求：</font></b></td>
              <td style="font-family: 宋体; font-size: 9pt" height="30" bgcolor="#FFFFFF" colspan="3">
 <p style="margin-left: 5px; margin-right: 5px"> <textarea rows="10" name="content"  class=SpanHeight style="width:100%;border-left:1px solid #FFFFFF; border-right:1px solid #FFFFFF; border-top:1px solid #FFFFFF; border-bottom:1px solid #C0C0C0; background:#FFFFFF;font-family: 楷体;font-size: 15pt;color:#0000ff"> 请各岗位严格按要求执行！</textarea>
</td>
            	</tr>
				<tr>
              <td width="18%" style="font-family: 宋体; font-size: 9pt;height:80"  bgcolor="#FFFFFF"> 
                <div align="center"><b><font color="#FF0000">上传文件：</font></b><p style="margin-top: 0; margin-bottom: 0" align="center">　<p style="margin-top: 0; margin-bottom: 0" align="center"><font color="#0000ff">(上传附件时如果文件名包含<p style="margin-top: 0; margin-bottom: 0" align="center">引号，请改名后再上传。)<p align="center"></div>
              </td>
              <td style="font-family: 宋体; font-size: 10pt;height:60" bgcolor="#FFFFFF" colspan="3">
				<p style="margin-top: 0; margin-bottom: 0"><input name="file1" style="border-left:1px solid #FFFFFF; border-right:1px solid #FFFFFF; border-top:1px solid #FFFFFF; border-bottom:1px solid #C0C0C0; background:#FFFFFF; " type="text" style="background:#FFFFFF; " id="file1" size="100">
                <font color=red>*</font><p style="margin-top: 0; margin-bottom: 0">　<p style="margin-top: 0; margin-bottom: 0">
				<iframe frameborder=0 width=80% height=60 scrolling=no src="upfile.asp?action=pic" name="I5"></iframe>
				</p>
				</td>
            	</tr>
                  </table>
            	</div>
            </table></table>
		</div>
	</form>
  </center>
<center>
  
<%end if%><!--#include file="footer.asp"-->