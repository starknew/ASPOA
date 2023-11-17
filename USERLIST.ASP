<!--#include file="conn.asp"--> <SCRIPT language=javascript src="css/init.js">
</SCRIPT>
<!--#include file="conn.asp"--> <!--#include file="asp/checkuser.asp"-->
<script language="JavaScript">
<%
opper=request("opper")
condition=""
for each idno in request("typename")
	count=count+1
	condition=condition+"|" & idno
	condition2=condition2+"reid=" & idno
	if count<request("delid").count then
	condition=condition+" or "
	condition2=condition2+" or "
	end if
next
'response.write condition
set rs1=server.CreateObject("adodb.recordset")
 if instr(1,"||supermanage||chksuper|",oabusyuserpower)>1 then
   sql="select * from dept   order by dept.no asc"
' else
'    sql="select * from dept  where dept='"&oabusyuserdept&"'"
'         if oabusyuserdept="皂角铺"  then
'        sql="select * from dept  where dept='皂角铺' or dept='绵阳北'"
'         end if
'        if oabusyuserdept="绵竹"  then
'         sql="select * from dept  where dept='绵竹' or dept='汉旺' or dept='孝泉'"
'         end if
'        if oabusyuserdept="什邡"  then
'         sql="select * from dept  where dept='什邡' or dept='永兴' or dept='高骈' or dept='两路口'"
'         end if
'        if oabusyuserdept="江油"  then
'         sql="select * from dept  where dept='江油' or dept='三合场' or dept='双河口'"
'         end if
'        if oabusyuserdept="黄许镇"  then
'         sql="select * from dept  where dept='黄许镇' or dept='德阳'"
'         end if
end if  
rs1.Open sql,conn,1,1
if rs1.EOF then
Response.Write request.cookies("oabusyuserdept")
	Response.Write "没有系统部门"
	Response.End 
end if

%><head>
<LINK href="style.css" rel=stylesheet>
<title>接收人员列表</title>
<SCRIPT language=Javascript>
<!-- hide
function insertsmilie(smilieface){
	window.opener.document.form2.joinman.value=window.opener.document.form2.joinman.value+'|'+smilieface;
	//self.close()	
}
// -->
</SCRIPT>
</head>
<body topmargin="0" leftmargin="0">
<br>
<center>
  <font size="5" color="#FF0000"><b>请选择车站</b></font> 
</center>
<form name="form2" method="post" action="userlist.asp">
  <font size="2"> 
  <table border="0" width="96%" cellspacing="0" bgcolor="#eeeeee" cellpadding="0" align="left">
    <%
    i=0
    
    do while not rs1.eof
		 i=i+1
   %>
		
      <td width="16%" align="left " valign="center" height="30">
      &nbsp; &nbsp; <input   type=checkbox name="typename" value="<%=rs1("dept")%>">
        <font size="3" color="#0000ff"><b> <%=rs1("dept")%> </b></font></td>
      <%if i mod 5=0  then%></tr><%end if
         rs1.MoveNext  
   loop
  rs1.Close
  set rs1=nothing
  %>
  <center></table><hr size="3" width="98%"> <input type="hidden" name="id" value="<%=request("id")%>">
    &nbsp;&nbsp; <input type="submit" name="typename" value="选择" style="border:1 solid #000000;background:#ff9966;font-size: 15px">
    &nbsp; <input type="reset" name="Submit3" value="重选" style="border:1 solid #000000;background:#ff9966;;font-size: 15px">
    <input type="hidden" name="opper" value=<%=opper%>>
    <font size="4" color="#0000FF"><b>注:选中上面的部门,部门内的人员将会全部选中! </b></font>

</form>
<font size="2">
<form id =form1 name=form1 method="post" action="<%=opper%>">
  <table border="0" width="96%" cellspacing="0" align="center">
    <%set rs=server.createobject("adodb.recordset")
'sql="select userinf.id,userinf.userdept,userinf.name,userinf.username  from userinf inner join dept on userinf.userdept=dept.dept order by userinf.userdept asc"
 if instr(1,"||supermanage||chksuper|",oabusyuserpower)>1 then
   sql="select userinf.id,userinf.no,userinf.userdept,userinf.name,userinf.gzgw,userinf.username,userinf.pb  from userinf inner join dept on userinf.userdept=dept.dept order by dept.no,cint(userinf.no) asc"
' else
'    sql="select userinf.id,userinf.no,userinf.userdept,userinf.name,userinf.gzgw,userinf.username,userinf.pb  from userinf inner join dept on userinf.userdept=dept.dept where userinf.userdept='"&oabusyuserdept&"'   order by dept.no,cint(userinf.no) asc"
'        if oabusyuserdept="皂角铺"  then
'    sql="select userinf.id,userinf.no,userinf.userdept,userinf.name,userinf.gzgw,userinf.username,userinf.pb  from userinf inner join dept on userinf.userdept=dept.dept where userinf.userdept='皂角铺' or userinf.userdept='绵阳北'  order by dept.no,cint(userinf.no) asc"
'        end if
'        if oabusyuserdept="绵竹"  then
'    sql="select userinf.id,userinf.no,userinf.userdept,userinf.name,userinf.gzgw,userinf.username,userinf.pb  from userinf inner join dept on userinf.userdept=dept.dept where userinf.userdept='绵竹' or userinf.userdept='汉旺' or userinf.userdept='孝泉'  order by dept.no,cint(userinf.no) asc"
'        end if
'        if oabusyuserdept="什邡"  then
'    sql="select userinf.id,userinf.no,userinf.userdept,userinf.name,userinf.gzgw,userinf.username,userinf.pb  from userinf inner join dept on userinf.userdept=dept.dept where userinf.userdept='什邡' or userinf.userdept='永兴' or userinf.userdept='高骈'  or userinf.userdept='两路口'  order by dept.no,cint(userinf.no) asc"
'         end if
'        if oabusyuserdept="江油"  then
'    sql="select userinf.id,userinf.no,userinf.userdept,userinf.name,userinf.gzgw,userinf.username,userinf.pb  from userinf inner join dept on userinf.userdept=dept.dept where userinf.userdept='江油' or userinf.userdept='三合场' or userinf.userdept='双河口'  order by dept.no,cint(userinf.no) asc"
'        end if
end if
rs.open Sql,Conn,1,1
if rs.EOF then
	Response.Write "没有选择人员"
	Response.End 
end if

Response.Write condition
 'Response.Write rs("userdept")
  do while not rs.eof%>
    <%if rs("userdept")=olduser  then
        else
         i=0%>
          <tr>
            <td width="57%" bgcolor="#cccccc" colspan="5" height="25"><font size="2"> 
            <p align="center"><font size="3"> <b><%=rs("userdept")%> </b> </font> </td>
        </tr>
  
   <%end if
      
      pb=rs("pb")
      if pb=1 then
          else
            i=i+1%>
       

           <td width="21%" height="30"><input type=checkbox name=nid value=<%=rs("username")%> 
           <%if instr(1,condition,"|"&trim(rs("userdept"))&"|")>0 then %>
              checked
           <%end if%>><font size="2"> 
           <%=rs("name")%>
          <font color="#0000ff">
           <%if rs("gzgw")="货运员" or rs("gzgw")="门检员" then %> <font color="#000000"> <%end if%>
           <%if rs("gzgw")="外勤货运员" then%> <font color="#2894ff"> <%end if%>
           <%if rs("gzgw")="内勤货运员" then%> <font color="#FF00ff"> <%end if%>
  
 (<%=rs("gzgw")%>)
     <%end if%>
      </font></td>
     <%
     if i mod 5 =0 then Response.Write "<tr>" 
	 olduser=rs("userdept")
	 rs.movenext
  loop
	%>
</table>

  <hr width="96%" size="1">
  <p align="center"> <input type="hidden" name="id" value="<%=request("id")%>"><input type="hidden" name="joinman">
  
    <input type="Submit" name="Submit" onclick="return submitup();" value="确定" style="border:1 solid #000000;background:#00FF00;font-weight: bold;font-size: 16px">
&nbsp;&nbsp; 
   <input type="reset" name="Submit2" value="清除" style="border:1 solid #000000;background:#00FF00;font-weight: bold;font-size: 16px">
  </p>
</form>
<script>
function submitup()
{
	var group = document.getElementsByTagName("input");
	var str="";
	for(var i=0;i<group.length;i++)
	{
		if(group[i].type=="checkbox" && group[i].name!="typename" && group[i].checked)
			str+=group[i].value+"|";
	}
	document.form1.joinman.value = str;
	return true;
}
</script>
</body>
<%
rs.Close
set rs=nothing 
Conn.Close
set Conn = nothing
 %><!--#include file="footer.asp"-->