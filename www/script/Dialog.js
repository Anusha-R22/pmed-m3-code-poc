//ic 12/12/2001
//function returns a string value to calling function from modal dialog
function fnReturn(sVal)
{
	this.returnValue=sVal;
	this.close();
}
//ic 03/05/2002
//function checks a string doesnt contain forbidden chars
function fnIllegalChars(sStr)
{
	var sPattern = /[`|~"]/;
	return sPattern.exec(sStr);
}	
//ic 12/12/2001
//function checks a passed string [sStr] isnt longer than a passed length [nLength]
function fnLength(sStr,nLength)
{
	if (sStr.length>nLength) return false;
	return true;
}
//ic 28/02 2008
//function trims whitespace from beginning and end of string
String.prototype.trim=function()
{
    return this.replace(/^\s*|\s*$/g,'');
}
//ic 12/12/2001
//function builds a delimited string from input, validates and returns if ok
function fnReturn1(sStr)
{
	var sRtn="";
	
	sStr = sStr.trim();
	if (sStr=="") return;
	if (!fnIllegalChars(sStr))
	{
		if (fnLength(sStr,255))
		{
			sRtn=sStr;
		}
		else
		{
			alert("The total length of input cannot exceed 255 characters")
			txtInput.select();
			return;
		}
	}
	else
	{
		alert("Input cannot contain double or backward quotes or the | character");
		txtInput.select();
		return;
	}		
	fnReturn(sRtn);
}
function fnSplitQS(sQS)
{
	var aItem;
	var aQS=sQS.substring(1).split("&")
	var aHashedArray=new Array();
	for (var n=0;n<aQS.length;n++)
	{	
		aItem=aQS[n].split("=");
		aHashedArray[aItem[0]]=aItem[1];
	}
	return aHashedArray;
}
function SelectChange(sSource,sDestination) 
{
	var oSelect=window.document.all[sSource];
	var oText=window.document.all[sDestination];
	oText.value = oSelect.value;
	oText.focus();
}
function fnRemoveEscapeChars(sStr)
{
	sStr=sStr.replace(/[+]/g," ");
	return sStr;
}
//function keyPress()
//{
//	if (window.event.keyCode==13)
//	{
//		fnReturn1(txtInput.value);
//	}
//}
//window.document.onkeypress=keyPress;
