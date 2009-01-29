///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	add onmouseover='fnOnMouseOver(this,nNumHeaderRows)' for row hover
//	add onmouseout='fnOnMouseOut(this)' for row hover
//	add onclick/onmousedown/onmouseup='fnOnClick(this)' for row select
//
///////////////////////////////////////////////////////////////////////////////////////////

var aTableObjects=new Array()

var sHOVERBCOLOUR='#e0e0ff';
var sHOVERFCOLOUR='#0000ff';
var sSELECTBCOLOUR='#0000ff';
var sSELECTFCOLOUR='#ffffff';

function fnPushTableObject(oTableObject)
{
	aTableObjects.push(oTableObject);
}
function fnAddTableObject(oTable)
{
	var oTO=new Object();
	oTO.sID=oTable.id;
	
	oTO.oCurrentRow=null;
	oTO.sHoverBColour="";
	oTO.sHoverFColour="";
	
	oTO.oSelectRow=null;		  //tracks currently selected row
	oTO.sSelectBColour="";		//tracks previous backgroundColor of selected row
	oTO.sSelectFColour="";		//tracks previous Color of selected row
	fnPushTableObject(oTO);
	return oTO;
}
function fnGetTableObject(oTable)
{
	for(var n=0;n<aTableObjects.length;n++)
	{
		if(aTableObjects[n].sID==oTable.id)
		{
			return aTableObjects[n]
		}
	}
	return fnAddTableObject(oTable);
}
function fnSelectRow(oTable,nRowIndex)
{
	var oTableObject=fnGetTableObject(oTable);
	oRow=oTable.rows[nRowIndex];
	if (oTableObject.oSelectRow!=null)
	{
		fnColourRow(oTableObject.oSelectRow,oTableObject.sSelectFColour,oTableObject.sSelectBColour);
	}	
	oTableObject.sSelectFColour="";
	oTableObject.sSelectBColour="";
	fnColourRow(oRow,sSELECTFCOLOUR,sSELECTBCOLOUR);
	oTableObject.oSelectRow=oRow;
	oTableObject.oCurrentRow=null;
}
function fnColourRow(oRow,sFColour,sBColour)
{
	oRow.style.backgroundColor=sBColour;
	oRow.style.color=sFColour;
}
function fnOnClick(oTable,nNumHeaderRows)
{
	var oTableObject=fnGetTableObject(oTable);
	nNumHeaderRows=(nNumHeaderRows!=undefined)?nNumHeaderRows:0;
	var oRow=GetRow(window.event.srcElement)
	if (oRow==null) return true;
	if (oRow!=oTableObject.oSelectRow)
	{	
		if (oTableObject.oSelectRow!=null)
		{	
			fnColourRow(oTableObject.oSelectRow,oTableObject.sSelectFColour,oTableObject.sSelectBColour);
		}
		if (oRow.rowIndex>=nNumHeaderRows)
		{
			oTableObject.sSelectFColour=oTableObject.sHoverFColour;
			oTableObject.sSelectBColour=oTableObject.sHoverBColour;
			fnColourRow(oRow,sSELECTFCOLOUR,sSELECTBCOLOUR)
			oTableObject.oSelectRow=oRow;
		}
		oTableObject.oCurrentRow=null;
	}
}
function fnOnMouseOver(oTable,nNumHeaderRows)
{
	var oTableObject=fnGetTableObject(oTable);
	nNumHeaderRows=(nNumHeaderRows!=undefined)?nNumHeaderRows:0;
	var oRow=GetRow(window.event.srcElement)
	if (oRow==null) return true;
	if ((oRow!=oTableObject.oCurrentRow)&&(oRow!=oTableObject.oSelectRow))
	{
		if (oTableObject.oCurrentRow!=null)
		{
			fnColourRow(oTableObject.oCurrentRow,oTableObject.sHoverFColour,oTableObject.sHoverBColour);
		}
		if (oRow.rowIndex>=nNumHeaderRows)
		{
			oTableObject.sHoverFColour=oRow.currentStyle.color;
			oTableObject.sHoverBColour=oRow.currentStyle.backgroundColor;
			fnColourRow(oRow,sHOVERFCOLOUR,sHOVERBCOLOUR)
			oTableObject.oCurrentRow=oRow;
		}
	}
}
function fnOnMouseOut(oTable)
{
	var oTableObject=fnGetTableObject(oTable);
	if (oTableObject.oCurrentRow!=null)
	{
		fnColourRow(oTableObject.oCurrentRow,oTableObject.sHoverFColour,oTableObject.sHoverBColour);
	}
	oTableObject.oCurrentRow=null;
}
function GetRow(oElem)
{
	while (oElem)
	{
		if (oElem.tagName.toLowerCase() == "tr"
		    && oElem.parentElement.tagName.toLowerCase() == "tbody") 
		    return oElem;
		if (oElem.tagName.toLowerCase() == "table") return null;
		oElem = oElem.parentElement;
	}
}