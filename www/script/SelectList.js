//function fnChange() called from a select list change event
//function expects a delimited list to be global:
//lSelect=item1subitem1|item1subitem2`item2subitem1`item3subitem1|item3subitem2|item3subitem3
//it then loads the array index (ndx) into the passed select list
//ie used to populate a 2nd select, based on the the item chosen in the 1st
function fnChange(oSelect,sList,ndx)
{
	var sDel1="`";
	var aList=sList.split(sDel1);
	fnLoadSelect(oSelect,aList[ndx],true);
}
function fnLoadSelect(oSelect,sList,bClear)
{
	var sDel2="|";
	var undefined;
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
	if (nCount>0) oSelect.selectedIndex=0;
}
			
		