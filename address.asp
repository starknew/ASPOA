<HTML>
<!-- Lotus-Domino (Release 5.0.6a - January 17, 2001 on Windows NT/Intel) -->
<HEAD>
<TITLE>选择管理员</TITLE><META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script src="global.js"></script>
<script src="images.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- 

// For convenience
var fieldAddress ;
var fieldAddressRef;
var fieldEntryList;
var origfield;

function OKClick()
{
	// put the values back in the underlying form in the right format
	if(window.opener)
	{
	eval("origfield=window.opener.document.forms[0]."+window.opener.document.forms[0].Address.value);
     origfield.value=selectToString(fieldAddress);
	}
	self.close();
}


function CancelClick() 
{
	self.close();
}


function AddClick(field)
{
	if(fieldEntryList.selectedIndex != -1)
	{
		for (var i=0; i < fieldEntryList.options.length; ++i)
		{
		if (fieldEntryList.options[i].selected)
			{		
				var selection = FixName(fieldEntryList.options[i].text);
				if (checkInList(fieldAddress,selection)==-1)
				field.options[field.options.length] = new Option(selection);
			}
		}
	}
}

function substring2(string, start, length) {
	return string.substring(start,start+length);
}


function FixName(name)
{
	if (name.indexOf("CN=")!= -1)
		{
			tmpstring=name.substring(3)
			
			var fullname=substring2(tmpstring,0,tmpstring.indexOf("/"));
			
			tmpstring=substring2(tmpstring,tmpstring.indexOf("/"),tmpstring.length);
			strpos = tmpstring.indexOf("=")+1;
			
			
			while (strpos != -1)
			{
			     tmpstring=substring2(tmpstring,strpos,tmpstring.length);     
				if (tmpstring.indexOf("/")!= -1)
				{
				restofname=substring2(tmpstring,0,tmpstring.indexOf("/")); 
				tmpstring = substring2(tmpstring,tmpstring.indexOf("/"),tmpstring.length);
				
				if (restofname.length!=0)
					{				  	
					fullname=fullname+"/"+restofname;
					}
				strpos=tmpstring.indexOf("=") + 1;  
				} else {
					fullname=fullname+"/"+tmpstring;
					strpos=-1;
					}		
				
			}
			return fullname;
		} else return name;
}

function RemoveClick(field)
{
	if (field.length != 0){ 
	if (field.selectedIndex != -1)
	{
		for(var i=field.options.length-1; i>=0 ; i--)
			{
				if (field.options[i].selected)
					{
						field.options[i] = null;
					}
			}
	}
}
}

function RemoveAllClick()
{
	fieldAddress.options.length=0;
	
}


function trim(str)
{
	for(var i = 0 ; i<str.length && str.charAt(i)==" " ; i++ ) ;
	return str.substring(i,str.length); 
}


function stringToSelect(str,field)
{
	for(var  beg=0 ; beg < str.length ; beg = end+1)
	{
		if(-1 == (end = str.indexOf(",",beg))) end = str.length;
		var entry = trim(str.substring(beg,end));
		if(entry!="") field.options[field.options.length++].text = entry;
	}
}


function selectToString(field)
{
	if (field.length!=0)
	{
	var str = "";
	for(i=0 ; i < field.options.length-1 ; i++)
		str += field.options[i].text + ",";
	return str += field.options[i].text;
	} else return "";
}

function checkInList(field,sText)
{
	if (field.length!=0)
	{
		var str = "";
		for(i=0 ; i < field.options.length ; i++)
		{
			str = trim(field.options[i].text).toLowerCase();
			if (str==trim(sText).toLowerCase()) return i;
		}
	} 
	return -1;
}

function onEnter()
{
    if (window.event.keyCode==13)
    {
      	onSearch();
		window.event.returnValue=false;
    }
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function mysubmit() {
	fieldAddressRef.value=selectToString(fieldAddress);
	document._wAddress.submit();
}

//-->
</SCRIPT>
</HEAD>
<BODY TEXT="000000" BGCOLOR="FFFFFF" onLoad="
fieldAddress = document.forms[0].tmpAddress;
fieldAddressRef = document.forms[0].tmpAddressRef;
fieldEntryList = document.forms[0].EntryList;
">
<FORM METHOD=post ACTION='address.asp' NAME='_wAddress'><table width='676' border='0' cellspacing='0' cellpadding='0' height='470'>  <tr>    <td background='images/bj-4.jpg'>    <table width=560 border='0' cellspacing='0' cellpadding='0' align='center'>      <tr>    <td colspan=3 height='30'><br></td>  </tr>  <tr>        <td colspan=3>          <br>    </td>  </tr>  <tr>    <td colspan=2>　</td>        <td></td>  </tr>  <tr>    <td><SELECT NAME="EntryList" MULTIPLE onDblClick="AddClick(fieldAddress);document.forms[0].tmpAddressRef.value=selectToString(fieldAddress);" onKeyPress="if (window.event.keyCode==13){    AddClick(fieldAddress);    document.forms[0].tmpAddressRef.value=selectToString(fieldAddress);}" ID="EntryList_1" STYLE="width:200" size=15>  <option value=刑警队>刑警队</option>  <option value=派出所>派出所</option></SELECT></td>        <td align=left> <table height=200><tr><td><a onClick="AddClick(fieldAddress); document.forms[0].tmpAddressRef.value=selectToString(fieldAddress);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('addone','','images/rd_zj_down.gif',1)">         <img style="cursor:hand" name="addone" border="0" src="images/rd_zj_up.gif"></a>        </td> </tr>        <tr>        <td align="center">          <a onClick="RemoveClick(tmpAddress);document.forms[0].tmpAddressRef.value=selectToString(fieldAddress);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('delete','','images/rb_del_down.gif',1)">          <img style=cursor:hand name="delete" border="0" src="images/rb_del_up.gif"></a>        </td>  </tr>        <tr>        <td align="center">          <a onClick="RemoveAllClick();document.forms[0].tmpAddressRef.value='';" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('delall','','images/rb_delall_down.gif',1)">          <img style=cursor:hand name="delall" border="0" src="images/rb_delall_up.gif"></a>        </td>        </tr></table></td>    <td><SELECT NAME="tmpAddress" MULTIPLE onDblClick="RemoveClick(tmpAddress);document.forms[0].tmpAddressRef.value=selectToString(fieldAddress);" onKeyPress="if (window.event.keyCode==13){    RemoveClick(tmpAddress);    document.forms[0].tmpAddressRef.value=selectToString(fieldAddress);}" ID="tmpSendTo_1" STYLE="width:200" SIZE=15><option value=123456>123456</option></SELECT></td></tr><tr>    <td colspan=3 align=left>         <input name='flggroup' type='radio' value='user' checked onClick='mysubmit();'>用户         <input name='flggroup' type='radio' value='group' checked onClick='mysubmit();'>群组 </td></tr> <tr>    <td colspan=3 align=center>        <HR height=1 color='#000000'><br>          <a onClick="OKClick()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('okbutton','','images/rd_qr_down.gif',1)">          <img style="cursor:hand" name="okbutton" border=0 src="images/rd_qr_up.gif"></a>          &nbsp; &nbsp; &nbsp; &nbsp; <a  onClick="CancelClick()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('cancelbutton','','images/rb_cancel_down.gif',1)">          <img style="cursor:hand" name="cancelbutton" border=0 src="images/rb_cancel_up.gif"></a></td>  </tr></table></td>  </tr></table><INPUT NAME="tmpAddressRef" TYPE=hidden VALUE="a"></FORM>

</BODY>
</HTML>
