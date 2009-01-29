///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	SelectDatabase page functions
//
///////////////////////////////////////////////////////////////////////////////////////////
var undefined;
var sDel1="|";
var sDel2="`";
var bSubmitted=false;
			
function fnLoadDb(sSelectedDb)
{
	fnLoadSelect(document.Form1.db,sDatabases,false,sSelectedDb);
	fnDbClick();
}
function fnDbClick()
{
	var aList=sRoles.split(sDel1);
	fnLoadSelect(document.Form1.rl,aList[document.Form1.db.selectedIndex],true,"");
}
function fnLoadSelect(oSelect,sList,bClear,sSelectedItem)
{
	var lOptions=oSelect.options;
	var aList=sList.split(sDel2);
	
	if (bClear)
	{
		for(var n=lOptions.length;n>=0;n--){lOptions[n]=null;}
	}
	var nCount=(lOptions.length==undefined)?0:lOptions.length;
	for (n=0;n<aList.length;n++)
	{
		lOptions[nCount] = new Option(aList[n],aList[n]);
		nCount++;
	}
	if (nCount>0) 
	{
		if (sSelectedItem!="")
		{
			for (n=0;n<oSelect.options.length;n++)
			{
				if (oSelect.options[n].value==sSelectedItem) oSelect.selectedIndex=n;
			}
		}
		if (oSelect.selectedIndex==-1) oSelect.selectedIndex=0;
	}	
}
function fnSubmit()
{
	if(bSubmitted) 
	{
		return;
	}
	else
	{
		Form1.btnSubmit.disabled=true;
		bSubmitted=true;
	}
	
	if(document.Form1.db.value!=sDatabase)
	{
		//database has changed, dont allow app state
		document.Form1.app.value="";
	}
	 document.Form1.submit();
}
function fnClick()
{
	if (window.event.keyCode==13) fnSubmit();
}
window.document.onkeypress=fnClick;