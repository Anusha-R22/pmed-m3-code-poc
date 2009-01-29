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
function keyPress()
{
	if (window.event.keyCode==13)
	{
		btnReturn.click();
	}
}
//ic 12/12/2001
//function builds a delimited RFC string from input, validates and returns if ok
function fnReturnRFC(sStr)
{
	var sRtn="";
		
	if (sStr=="") return;
	if (!fnIllegalChars(sStr))
	{
		if (fnLength(sStr,255))
		{
			sRtn=sStr;
		}
		else
		{
			alert("The total length of reason for change cannot exceed 255 characters")
			txtRFC.select();
			return;
		}
	}
	else
	{
		alert("Reason for change cannot contain double or backward quotes or the | character");
		txtRFC.select();
		return;
	}		
	fnReturn(sRtn);
}
function fnPageLoaded()
{
	var aQS=fnSplitQS(window.location.search);
	document.all["divName"].innerHTML="<b>Question: "+aQS["name"]+"</b>";
	document.all["divLabel"].innerHTML="Please enter or choose the reason for changing "+aQS["name"];
	txtRFC.focus();
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
window.document.onkeypress=keyPress;
