////////////////////
////////////////////
// Public methods
////////////////////
////////////////////

//
// Create a slect list control.
//
// sLocation		= DOM textual reference to the container object eg "window.document"
// sID				= identification string for this control. name to be used for form element
// sValues			= delimited value-pair list of select list items
//						(pairs delimited with "|", items within pair delimited with "`")
// nTabIndex		= the control's tab index
// nWidth			= pixel width
// nHeight			= pixel height
// nAlign			= justification of text

//
// Returns a reference to the select object,
//	which supports the following public methods:
//
//global array of select objects
var aSelectObjects=new Array();

//global enums
var eAlignLeft=0;
var eAlignCentre=1;
var eAlignRight=2;

function fnPushSelect(oSelect)
{
	aSelectObjects.push(oSelect);
}
function fnPopSelect(sID)
{
	for(var n=0;n<aSelectObjects.length;n++)
	{
		if(aSelectObjects[n].sID==sID) return aSelectObjects[n];
	}
	return null;
}
function fnInitZOrder()
{
	//short term solution to z-index problems when innerhtml is changed
	for (var n=0;n<aSelectObjects.length;n++)
	{
		fnSelectOpen(aSelectObjects[n]);
		fnSelectClose(aSelectObjects[n]);
	}
}
function fnDocumentOnClick()
{
	var sID="";
	oElement=window.event.srcElement;
	while ((oElement.id.indexOf("sel_table")==-1)&&(oElement.tagName.toLowerCase()!="body"))
	{
		oElement=oElement.parentElement;
	}
	if(oElement.id.indexOf("sel_table")!=-1)
	{
		//click event originated on select list
		sID=oElement.id.slice(10);
	}
	fnSelectCloseAll(sID);
}
function fnSelectCreate(sLocation,sID,nTabIndex,nWidth,nHeight,nAlign)
{
	var oSelect=new Object();
	oSelect.sLocation=sLocation+".all."+sID+"_c";  //container
	oSelect.sPLocation=sLocation+".all."+sID+"_p"; //pane
	oSelect.sVLocation=sLocation+".all."+sID+"_v"; //value display
	oSelect.sFLocation=sLocation+".all."+sID	   //hidden field
	oSelect.sTLocation=sLocation+".all."+sID+"_t"; //tooltip
	
	oSelect.sID=sID;
	oSelect.nWidth=nWidth;
	oSelect.nHeight=nHeight;
	oSelect.nAlign=nAlign;
	oSelect.sIndex=null;

//
// Assign the object's methods - see functions called for detailed descriptions
//
	oSelect.setValue=function(sIndex)
	{
		return fnSetValue(this,sIndex);
	}
	oSelect.populate=function(sList,bInit)
	{
		return fnPopulate(this,sList,bInit);
	}
	oSelect.getIndex=function()
	{
		return fnGetIndex(this);
	}
	oSelect.getValue=function()
	{
		return fnGetValue(this);
	}
	oSelect.getText=function()
	{
		return fnGetText(this);
	}
	fnPushSelect(oSelect);
	fnDraw(oSelect);
	return oSelect;
}

//////////////////////////////////////////////////
// Private Methods - used by the object methods
//////////////////////////////////////////////////

//
// Useful keyboard definitions - global scope
//
var kDown=40;
var kUp=38;
var kSpace=32;
var kEnter=13;
var kTab=9;

function fnGetIndex(oSelect)
{
	return oSelect.sIndex;
}
function fnGetValue(oSelect)
{
	if (oSelect.sIndex===null)
	{
		return "";
	}
	else
	{
		return oSelect.sIndex;
	}
	
}
function fnGetText(oSelect)
{
	if (oSelect.sIndex===null)
	{
		return "";
	}
	else
	{
		return oSelect.olNames[oSelect.sIndex];
	}
}
function fnSetValue(oSelect,sIndex)
{
//	if(sIndex!=null)
//	{
		oSelect.sIndex=sIndex;
		eval(oSelect.sVLocation).innerHTML=oSelect.getText();
		eval(oSelect.sFLocation).value=oSelect.getValue();
		eval(oSelect.sTLocation).title=oSelect.getText();
//	}
}

///////////////////////////////////////////////////////////
// Private functions - used by the private methods above
///////////////////////////////////////////////////////////

function fnSelectOpenClose(sID)
{
	var oSelect=fnPopSelect(sID);
	var oPane=eval(oSelect.sPLocation);
	if(oPane.style.visibility=="visible")
	{
		fnSelectClose(oSelect);
	}
	else
	{
		fnSelectOpen(oSelect);
	}
}
function fnSelectOpen(oSelect)
{
	var oSelectDiv=eval(oSelect.sLocation);
	var oPane=eval(oSelect.sPLocation);
	oPane.style.top=fnOffsetTop(oSelect)+oSelectDiv.offsetHeight;
	oPane.style.left=fnOffsetLeft(oSelect); //oSelectDiv.style.left;
	oPane.style.visibility="visible";
}
function fnSelectClose(oSelect)
{
	var oPane=eval(oSelect.sPLocation);
	oPane.style.top=0;
	oPane.style.left=0;
	oPane.style.visibility="hidden";
}
function fnSelectClick(sID,sIndex)
{
	var oSelect=fnPopSelect(sID);
	if(oSelect.sIndex!=sIndex)
	{
		//value has changed		
		fnSetValue(oSelect,sIndex);		
		fnSelectChange(sID,fnGetValue(oSelect));
	}
	fnSelectClose(oSelect);
}
function fnOffsetTop(oSelect)
{
	var n=0;
	var oElement=eval(oSelect.sLocation);
	while (oElement.tagName.toLowerCase()!="body")
	{
		if(oElement.tagName.toLowerCase()=="div")
		{
			n+=oElement.offsetTop;
		}
		oElement=oElement.parentElement;
	}
	return n;
}
function fnOffsetLeft(oSelect)
{
	var n=0;
	var oElement=eval(oSelect.sLocation);
	while (oElement.tagName.toLowerCase()!="body")
	{
		if(oElement.tagName.toLowerCase()=="div")
		{
			n+=oElement.offsetLeft;
		}
		oElement=oElement.parentElement;
	}
	return n+1;
}
function fnDraw(oSelect)
{
	//select box
	var osHTML=new Array();
	osHTML.push("<table id='sel_table_"+oSelect.sID+"' onclick='fnSelectOpenClose(\""+oSelect.sID+"\");' cellpadding='0' cellspacing='0' border='0'>");
	osHTML.push("<tr id='"+oSelect.sID+"_t' class='clsMSelectList' valign='middle'>");
	osHTML.push("<td align='"+fnAlign(oSelect.nAlign)+"'>");
	osHTML.push("<div id='"+oSelect.sID+"_s' class='clsMSelectList' style='z-index:9; height:"+oSelect.nHeight+"px;width:"+(oSelect.nWidth-16)+"px;'>");
	osHTML.push("<div id='"+oSelect.sID+"_v'></div>");
	osHTML.push("</div>");
	osHTML.push("</td><td valign='top'><div style='height:"+oSelect.nHeight+"px; width:16px' class='clsMSelectButton'><img onmouseout='fnMU(this)' onmousedown='fnMD(this)' onmouseup='fnMU(this)' src='../img/exp_inactive.gif'></div></td></tr>");
	osHTML.push("<input type='hidden' value='' name='"+oSelect.sID+"'>")
	eval(oSelect.sLocation).innerHTML=osHTML.join('');	
}
function fnPopulate(oSelect,sList,bInit)
{
	oSelect.olNames=null;
	oSelect.olNames=new Array();
	var sStart='\\|';
	var sTerm = "([^|]*)";
	var sDel = '\`';
	var sHTML;

	if((sList!=undefined)&&(sList!=""))
	{
		var sAlign="align='"+fnAlign(oSelect.nAlign)+"' ";
		var aValues=sList.split("|");	

		for(var n=1;n<aValues.length;n++)
		{
			var aValue=aValues[n].split("`");
			//populate select object's internal arrays
			oSelect.olNames[aValue[0]]=aValue[1];           
		}
		//use regular expression to build select html
		//ic 04/02/2004 use ~ as separator to avoid name clashes with other ids
		sHTML=sList.replace(new RegExp(sStart+sTerm+sDel+sTerm,'g'),'<tr height="10" id="'+oSelect.sID+'~$1"><a style="text-decoration:none;" href="javascript:fnSelectClick(\''+oSelect.sID+'\',\'$1\');"><td '+sAlign+' title="$2">$2</td></a></tr>')
		eval(oSelect.sPLocation).innerHTML="<table id='"+oSelect.sID+"_pt' cellpadding='0' cellspacing='0' onclick='fnOnClick(this)' onmouseover='fnOnMouseOver(this,0);' onmouseout='fnOnMouseOut(this);' width='100%' class='clsTableText'>"+sHTML+"</table>";

		if(aValues.length>0)
		{ 
			fnSelectRow(document.all[oSelect.sID+"_pt"],0);			
			fnSetValue(oSelect,aValues[1].split("`")[0]);
		}
	}
	else
	{
		eval(oSelect.sPLocation).innerHTML="";
		fnSetValue(oSelect,null);
	}
	if(bInit) fnInitZOrder();	
}
function fnSelectCloseAll(sID)
{
	for (var n=0;n<aSelectObjects.length;n++)
	{
		if(aSelectObjects[n].sID!=sID) fnSelectClose(aSelectObjects[n]);
	}
}
function fnMD(oImg)
{
//	oImg.src="../img/exp_active.gif"
}
function fnMU(oImg)
{
//	oImg.src="../img/exp_inactive.gif"
}
function fnAlign(nAlign)
{
	switch (nAlign)
	{
		case eAlignLeft: return "left";break;
		case eAlignCentre: return "center";break;
		case eAlignRight: return "right";break;
	}
}
