////////////////////
////////////////////
// Public methods
////////////////////
////////////////////

//
// Create a radio button control.
//
// oLocation		= object *containing* the "DIV" to hold the control;
// sLocation		= DOM textual reference to the above object;
// sID				= identification string for this control - name of the "DIV";
// sValues			= delimited value-pair list of radio buttons;
//						(pairs delimited with "~", items within pair delimited with "¬");
// sInitialValue	= initial selected radio *value* (not index) or a null for none;
// nTabIndex		= the control's tab index;
//
// The control created will be disabled.
//
// Returns a reference to the radio-control object,
//	which supports the following public methods:
//


function fnRadioCreate(oLocation,sLocation,sID,sValues,sInitialValue,nTabIndex,sValueRef,sColour,sStyle,nRQGno,bRQGMultiple)
{
	var oRadio=new Object();
	if((nRQGno==undefined)||((nRQGno==0)&&((bRQGMultiple==undefined)||(!bRQGMultiple))))
	{
		oRadio.oLocation=oLocation.all[sID];
		oRadio.oDataLocation=oLocation.all[sValueRef];
		oRadio.sLocation=sLocation+".all."+sID;
	}
	else
	{
		oRadio.oLocation=(oLocation.all[sID])[nRQGno];
		oRadio.oDataLocation=(oLocation.all[sValueRef])[nRQGno];
		oRadio.sLocation=sLocation+".all."+sID+"["+nRQGno+"]";
	}
	oRadio.sValueRef=sValueRef;
	oRadio.sID=sID;
	oRadio.sColour=sColour;
	oRadio.nTabIndex=nTabIndex;
	oRadio.bEnabled=false;
	oRadio.olValues=new Array();
	oRadio.olNames=new Array();
	slPairs=sValues.split("~");	// get a list of pairs
	oRadio.nOptionCnt=slPairs.length;
	oRadio.nIndex=null;	// *Index* of selected choice
	// Build up two arrays of values and names
	for(var nPair=0;nPair<oRadio.nOptionCnt;++nPair)
	{
		var slCouple=slPairs[nPair].split("¬");
		oRadio.olValues[nPair]=slCouple[0];
		oRadio.olNames[nPair]=slCouple[1];
		if(slCouple[0]==sInitialValue)
		{
			oRadio.nIndex=nPair;
		}
	}
	oRadio.sValue=(oRadio.nIndex===null?null:sInitialValue);	// *Value* of selected choice
	oRadio.nHighlighted=null;	// Start with nothing highlighted
	oRadio.bFocussed=false;
	oRadio.sFGCol="000000";
	oRadio.sHiBGCol="FFFFFF";
	oRadio.sUnhiBGCol="FFFFFF";
	oRadio.nNoteComment=0; //0=none,1=note,2=comment,3=both
	oRadio.nRQGno=nRQGno;
	oRadio.sDisabledFontCol="888888";
	if((sStyle==undefined)||(sStyle===null))
	{
		oRadio.sStyle="";
	}
	else
	{
		oRadio.sStyle=sStyle;
	}

//
// Assign the object's methods - see functions called for detailed descriptions
//
	oRadio.getIndex=function()
	{
		return fnGetIndex(this);
	}
	oRadio.getValue=function()
	{
		return fnGetValue(this);
	}
	oRadio.getText=function()
	{
		return fnGetText(this)
	}
	oRadio.setIndex=function(nIndex,bNoPropogate,bForce)
	{
		bForce=(bForce!=undefined)? bForce:false;
		return fnSetIndex(this,nIndex,bNoPropogate,bForce);
	}
	oRadio.setValue=function(svalue,bNoPropogate,bForce)
	{
		bForce=(bForce!=undefined)? bForce:false;
		return fnSetValue(this,svalue,bNoPropogate,bForce);
	}
	oRadio.enable=function(bEnabled,sFGCol,sHiBGCol,sUnhiBGCol)
	{
		return fnEnable(this,bEnabled,sFGCol,sHiBGCol,sUnhiBGCol);
	}
	oRadio.colour=function(sFGCol,sHiBGCol,sUnhiBGCol)
	{
		return fnColour(this,sFGCol,sHiBGCol,sUnhiBGCol);
	}
	oRadio.focus=function(bNoPropogate)
	{
		return fnFocus(this,bNoPropogate);
	}
	oRadio.onfocus=function()
	{
		return fnOnfocus(this);
	}
	oRadio.blur=function(bNoPropogate,bRedraw)
	{
		return fnBlur(this,bNoPropogate,bRedraw);
	}
	oRadio.onblur=function()
	{
		return fnOnblur(this);
	}
	oRadio.onchange=function()
	{
		return fnOnchange(this);
	}
	oRadio.RadioNoteStatus=function(bNoteComment)
	{
		return fnRadioNoteStatus(this,bNoteComment);
	}
	fnCreateStructures(oRadio);
	fnRedraw(oRadio);
	return oRadio;
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

var nFakeIndex="-1";
var nFocusDelay=200;

//
// Get currently selected index, or null if none
//
function fnGetIndex(oRadio)
{
	return oRadio.nIndex;
}

//
// Get currently selected value, or null if none
//
function fnGetValue(oRadio)
{
	if (oRadio.nIndex===null)
	{
		return null;
	}
	else
	{
		return oRadio.olValues[oRadio.nIndex];
	}
}

//
// Get currently selected text, or null if none
//

function fnGetText(oRadio)
{
	if (oRadio.nIndex===null)
	{
		return null;
	}
	else
	{
		return oRadio.olNames[oRadio.nIndex];
	}
}

//
// Set the selection by index
// null will deselect whatever the selected value was
// If bNoPropogate is not true, the onchange() event will be called
// Has no effect if the control is disabled
//
function fnSetIndex(oRadio,nIndex,bNoPropogate,bForce)
{
//ic 12/06/2003 
// commented out condition, allow disabled radios to have values set - caters for derived radios
//	if((oRadio.bEnabled)||bForce)
//	{
		oRadio.nIndex=nIndex;
		if(nIndex!=null)
		{
			oRadio.oDataLocation.value=oRadio.olValues[nIndex];
		}
		else
		{
			oRadio.oDataLocation.value="";
		}
		fnRedraw(oRadio);
		if((!bNoPropogate)&&(nIndex!=null))
//		if(!bNoPropogate)
		{
			oRadio.onfocus();
		}
//	}
}

//
// Set the selection by value
// null will deselect whatever the selected value was.
// If bNoPropogate is not true, the onchange() event will be called
//
function fnSetValue(oRadio,sValue,bNoPropogate,bForce)
{
	var nIndex=null;
	var nCode1;
	var nCode2;
	for(var nOpt=0;nOpt<oRadio.nOptionCnt;++nOpt)
	{
		//change to same case before comparing 
		nCode1=oRadio.olValues[nOpt]+"";
		nCode1=nCode1.toLowerCase();
		nCode2=sValue+"";
		nCode2=nCode2.toLowerCase()
		if(nCode1==nCode2)
		{
			nIndex=nOpt;
			oRadio.oDataLocation.value=sValue;
		}
	}
	fnSetIndex(oRadio,nIndex,bNoPropogate,bForce);
}

//
// Enable or disable the control.
// Needs the foreground colour
// and the background colours for the selected and unselected states
//
function fnEnable(oRadio,bEnabled,sFGCol,sHiBGCol,sUnhiBGCol)
{
	oRadio.bEnabled=bEnabled;
	if(bEnabled)
	{
		oRadio.oLocation.all[oRadio.sID+"_dummy"].tabIndex=oRadio.nTabIndex;
	}
	else
	{
		oRadio.oLocation.all[oRadio.sID+"_dummy"].tabIndex=-1;
	}
	fnColour(oRadio,sFGCol,sHiBGCol,sUnhiBGCol);
}

//
// Sets the three colours for the control
//
function fnColour(oRadio,sFGCol,sHiBGCol,sUnhiBGCol)
{
	if(sFGCol!=null)
	{
		oRadio.sFGCol=sFGCol;
	}
	if(sHiBGCol!=null)
	{
		oRadio.sHiBGCol=sHiBGCol;
	}
	if(sUnhiBGCol!=null)
	{
		oRadio.sUnhiBGCol=sUnhiBGCol;
	}
	fnRedraw(oRadio);
}

//
// Apply focus to the control
// If bNoPropogate is not true, the onfocus() event will be called
//
function fnFocus(oRadio,bNoPropogate)
{
	if(oRadio.bEnabled)
	{
		oRadio.nHighlighted=oRadio.nIndex;
		if((oRadio.nHighlighted<0)||(oRadio.nHighlighted>=oRadio.nOptionCnt)||(oRadio.nHighlighted===null))
		{
			oRadio.nHighlighted=oRadio.nIndex!=null?oRadio.nIndex:0;
		}
		fnRedraw(oRadio);
		if(!bNoPropogate)
		{
			oRadio.onfocus();
		}
		fnFocusOnDummy(oRadio);
	}
}

//
// The control has gained focus. This can be overridden by a user-defined method.
//
function fnOnfocus(oRadio)
{
	// No default behaviour
}

//
// Lose focus from the control
// If bNoPropogate is true, the onblur() event will be called
//
function fnBlur(oRadio,bNoPropogate,bRedraw)
{
	oRadio.nHighlighted=null;
	oRadio.bFocussed=false;
	if (bRedraw)
	{
		fnRedraw(oRadio);
	}
	if(!bNoPropogate)
	{
		oRadio.onblur();
	}
}

//
// The control has lost focus. This can be overridden by a user-defined method.
//
function fnOnblur(oRadio)
{
	// No default behaviour
}

//
// The control's value has changed. This can be overridden by a user-defined method.
// Note: this event fires when the change occurs, not when the focus shifts after a change.
//
function fnOnchange(oRadio)
{
	// No default behaviour
}

///////////////////////////////////////////////////////////
// Private functions - used by the private methods above
///////////////////////////////////////////////////////////
//
// Control got focus
//
function fnRadioFocus(sRadio)
{
	var oRadio=eval(sRadio);
	oRadio.bFocussed=true;
	
	if((oRadio.nHighlighted<0)||(oRadio.nHighlighted>=oRadio.nOptionCnt)||(oRadio.nHighlighted===null))
	{
		oRadio.nHighlighted=oRadio.nIndex!=null?oRadio.nIndex:0;
	}
	fnRedraw(oRadio);
	fnFocusOnDummy(oRadio);
}

//
// An internal blur() event fired to get us here
//
function fnRadioDummyBlur(sRadio)
{
}

//
// An internal onclick() event fired to get us here
//
function fnRadioClick(sRadio,nOpt)
{
	var oRadio=eval(sRadio);
	if(oRadio.bEnabled)
	{
		oRadio.bFocussed=true;
		if(nOpt!=oRadio.nIndex)
		{
			oRadio.nIndex=nOpt;
			if(nOpt!=null)
			{
				oRadio.oDataLocation.value=oRadio.olValues[nOpt];
			}
			else
			{
				oRadio.oDataLocation.value="";
			}
			oRadio.onchange();
		}
		oRadio.nHighlighted=nOpt;
		fnRedraw(oRadio);
		fnSetFocus(oRadio);
	}
}

//
// An internal onkeydown() event fired to get us here
//
function fnRadioKeyPress(sRadio)
{
	var oRadio=eval(sRadio);
	if(oRadio.bEnabled)
	{
		var nKey=window.event.keyCode;
		var bKeyUsed=false;
		var bDataChanged=false;
		switch (nKey)
		{
			// Enter and Space both select the highlighted option
			case kEnter:
			case kSpace:
				if((oRadio.nHighlighted!=null)
					&&(oRadio.nHighlighted>=0)
					&&(oRadio.nHighlighted<oRadio.nOptionCnt))
				{
					oRadio.nIndex=oRadio.nHighlighted;
					bDataChanged=true;
					bKeyUsed=true;
				}
			break;
			// Move up a selection - wrap if required
			case kUp:
				if(oRadio.nHighlighted===null)
				{
					oRadio.nHighlighted=oRadio.nIndex;
				}
				if(oRadio.nHighlighted>0)
				{
					--oRadio.nHighlighted;
				}
				else
				{
					oRadio.nHighlighted=oRadio.nOptionCnt-1;
				}
				bKeyUsed=true;
			break;
			// Move down a selection - wrap if required
			case kDown:
				if(oRadio.nHighlighted===null)
				{
					oRadio.nHighlighted=oRadio.nIndex;
				}
				else if(oRadio.nHighlighted<oRadio.nOptionCnt-1)
				{
					++oRadio.nHighlighted;
				}
				else
				{
					oRadio.nHighlighted=0;
				}
				bKeyUsed=true;
			break;
		}
		if(bKeyUsed)
		{
			// We acted on a key-press - do the necessary screen re-draw and fire some events
			oRadio.bFocussed=true;

			var	sFGCol=oRadio.sFGCol;
			var	sHiBGCol=oRadio.sHiBGCol;
			var	sUnhiBGCol=oRadio.sUnhiBGCol;
			//ic 13/05/2002
			if(bDataChanged)
			{
				oRadio.oDataLocation.value=oRadio.olValues[oRadio.nIndex];
				oRadio.onchange();
			}

			fnRedraw(oRadio);
			window.event.returnValue=false;
			return false;
		}
		else
		{
			// Not acted on - allow the browser to see the key-press
			return true;
		}
	}
	else
	{
		// Not acted on - allow the browser to see the key-press
		return true;
	}
}

//
// Set the browser's focus to the radio button's master controller
//
function fnSetFocus(oRadio)
{
	//DPH 28/11/2002 oLocation.all sometimes empty, check is OK before using
	if((oRadio.oLocation.all[oRadio.sID+"_dummy"]!=undefined)&&(oRadio.oLocation.all[oRadio.sID+"_dummy"]!=null))
	{
		oRadio.oLocation.all[oRadio.sID+"_dummy"].focus();
	}
}

//ic 21/08/2002
//function updates radio note/comment status propery, redraws radio
//
function fnRadioNoteStatus(oRadio,nNoteComment)
{
	//0=none,1=note,2=comment,3=both
	oRadio.nNoteComment=nNoteComment;
	fnRedraw(oRadio);
}


//
// Set the browser's focus to the radio button's dummy controller.
//	This has no on-focus event, but is otherwise the same as the master control.
//	This allows the fnSetFocus function to redraw the table
//	(which overwrites the mastercontrol), then place focus to somewhere without
//	an onfocus() event to start the whole thing over again!
//
function fnSetDummyFocus(oRadio,bFocus)
{
	oRadio.onfocus();
	//DPH 12/12/2002 oLocation.all sometimes empty, check is OK before using
	if((oRadio.oLocation.all[oRadio.sID+"_dummy"]!=undefined)&&(oRadio.oLocation.all[oRadio.sID+"_dummy"]!=null))
	{
		if(!fnLoading()) oRadio.oLocation.all[oRadio.sID+"_dummy"].focus();
	}
}

function fnCreateStructures(oRadio)
{
	var sDummyIndex=oRadio.nTabIndex;
	var sControlIndex=nFakeIndex;
	var sHTML='';
	var osHTML=new Array();
	var sQID="";
	var nRQGno=0;
	// Draw master control
	sHTML='<a href="#1" id="';
	osHTML.push(sHTML);	
	osHTML.push(oRadio.sID);	
	sHTML='_control"';
	osHTML.push(sHTML);	
	sHTML='></a> ';
	osHTML.push(sHTML);	
	// End of master control

	// Draw "dummy" control.
	// This is needed as the above "onfocus" needs to redraw the table, but appear to keep focus on it.
	sHTML='<a href="#2" id="';
	osHTML.push(sHTML);	
	osHTML.push(oRadio.sID);	
	sHTML='_dummy"';
	osHTML.push(sHTML);	
	sHTML='onfocus="fnRadioFocus(\'';
	osHTML.push(sHTML);	
	osHTML.push(oRadio.sLocation);	
	sHTML='.oRadio\');window.status+=\'\';"';
	osHTML.push(sHTML);	
	sHTML='onkeydown="fnRadioKeyPress(\'';
	osHTML.push(sHTML);	
	osHTML.push(oRadio.sLocation);	
	sHTML='.oRadio\');"';
	osHTML.push(sHTML);	
	sHTML='tabindex="';
	osHTML.push(sHTML);	
	osHTML.push(sDummyIndex);	
	sHTML='"';
	osHTML.push(sHTML);	
	sHTML='></a> ';
	osHTML.push(sHTML);
	// End of dummy control

	// Start the changable DIV
	sHTML='<div id="';
	osHTML.push(sHTML);	
	osHTML.push(oRadio.sID);	
	sHTML='_dyn" style="position:relative;z-index:2;"';
	osHTML.push(sHTML);
	// remove _inpDIV to get question ID
	sQID=oRadio.sID.substring(0,(oRadio.sID.length-7));
	if(oRadio.nRQGno!=undefined)
	{
		nRQGno=oRadio.nRQGno;
	}
	sHTML=' onMouseOver="fnSetTooltip(\''+sQID+'\','+nRQGno+',0);"';
	osHTML.push(sHTML);
	sHTML='></div>';
	osHTML.push(sHTML);
	oRadio.oLocation.oRadio=oRadio;
	oRadio.oLocation.innerHTML=osHTML.join('');
	delete osHTML;
}

//ic 21/08/2002
//added style code to handle sdv control border, icon code to handle note/comment icon
//added extra table around radio, first cell for radio control, second cell for note/comment icon
function fnRedraw(oRadio)
{
	var sIcon="";

	var sHiBGCol;
	var sHTML='';
	var osHTML=new Array();
	var sFontCol;
	
	if(!oRadio.bEnabled)
	{
		sHiBGCol=oRadio.sUnhiBGCol;
		oRadio.bFocussed=false;
		oRadio.nHighlighted=-1;
		sFontCol=oRadio.sDisabledFontCol;
	}
	else
	{
		sFontCol=oRadio.sColour;
		//sFontCol=oRadio.sFGCol;
	}
	if(oRadio.bFocussed)
	{
		sHiBGCol=oRadio.sHiBGCol;
	}
	else
	{
		sHiBGCol=oRadio.sUnhiBGCol;
	}
	if((oRadio.nHighlighted<0)
		||(oRadio.nHighlighted>=oRadio.nOptionCnt)
		||(oRadio.nHighlighted===null))
	{
		oRadio.nHighlighted=oRadio.nIndex!=null?oRadio.nIndex:0;
	}
	
	switch (oRadio.nNoteComment)
	{
		case eNote:
			sIcon='<img src="../img/ico_note.gif">'
			break;
		case eComment:
			sIcon='<img src="../img/ico_comment.gif">'
			break;
		case eCommentNote:
			sIcon='<img src="../img/ico_note_comment.gif">'
			break;
		default:
	}

	
	sHTML='<table cellpadding="0" cellspacing="0">';
	osHTML.push(sHTML);	
	sHTML='<tr><td>';
	osHTML.push(sHTML);
	sHTML='<table id="';
	osHTML.push(sHTML);
	osHTML.push(oRadio.sID);	
	sHTML='_table" border="0" rules="none">';
	osHTML.push(sHTML);
	
	for(var nOpt=0;nOpt<oRadio.nOptionCnt;++nOpt)
	{
		if(nOpt===oRadio.nIndex)
		{
			// Selected option
			if((nOpt===oRadio.nHighlighted)&&(oRadio.bFocussed))
			{
				sHTML='<tr bgcolor="';
				osHTML.push(sHTML);	
				osHTML.push(sHiBGCol);	
				sHTML='" ';
				osHTML.push(sHTML);	
				if(oRadio.bEnabled)
				{
					sHTML='onclick="fnRadioClick(';
					osHTML.push(sHTML);	
					osHTML.push(oRadio.sLocation);	
					sHTML='.oRadio,';
					osHTML.push(sHTML);	
					osHTML.push(nOpt);	
					sHTML=');"';
					osHTML.push(sHTML);	
				}
				sHTML='><td><img src="../img/RadioSelHi';
				osHTML.push(sHTML);
			}
			else
			{
				sHTML='<tr bgcolor="';
				osHTML.push(sHTML);	
				osHTML.push(oRadio.sUnhiBGCol);	
				sHTML='" ';
				osHTML.push(sHTML);	
				if(oRadio.bEnabled)
				{
					sHTML='onclick="fnRadioClick(';
					osHTML.push(sHTML);	
					osHTML.push(oRadio.sLocation);	
					sHTML='.oRadio,';
					osHTML.push(sHTML);	
					osHTML.push(nOpt);	
					sHTML=');"';
					osHTML.push(sHTML);
				}
				sHTML='><td><img src="../img/RadioSel';
				osHTML.push(sHTML);
			}
		}
		else
		{
			// Not the selected option
			if((nOpt===oRadio.nHighlighted)&&(oRadio.bFocussed))
			{
				sHTML='<tr bgcolor="';
				osHTML.push(sHTML);	
				osHTML.push(sHiBGCol);	
				sHTML='" ';
				osHTML.push(sHTML);
				if(oRadio.bEnabled)
				{
					sHTML='onclick="fnRadioClick(';
					osHTML.push(sHTML);	
					osHTML.push(oRadio.sLocation);	
					sHTML='.oRadio,';
					osHTML.push(sHTML);	
					osHTML.push(nOpt);	
					sHTML=');"';
					osHTML.push(sHTML);
				}
				sHTML='><td><img src="../img/RadioUnselHi';
				osHTML.push(sHTML);
			}
			else
			{
				sHTML='<tr bgcolor="';
				osHTML.push(sHTML);	
				osHTML.push(oRadio.sUnhiBGCol);	
				sHTML='" ';
				osHTML.push(sHTML);
				if(oRadio.bEnabled)
				{
					sHTML='onclick="fnRadioClick(';
					osHTML.push(sHTML);	
					osHTML.push(oRadio.sLocation);	
					sHTML='.oRadio,';
					osHTML.push(sHTML);	
					osHTML.push(nOpt);	
					sHTML=');"';
					osHTML.push(sHTML);
				}
				sHTML='><td><img src="../img/RadioUnsel';
				osHTML.push(sHTML);	
			}
		}
		sHTML='.gif"></td><td>';
		osHTML.push(sHTML);
		//sHTML='<font color="';
		sHTML='<font style="color:#';
		osHTML.push(sHTML);	
		osHTML.push(sFontCol);
		sHTML=';';
		osHTML.push(sHTML);
		osHTML.push(oRadio.sStyle);	
		sHTML='">';
		osHTML.push(sHTML);	
		osHTML.push(oRadio.olNames[nOpt]);	
		sHTML='</font>';
		osHTML.push(sHTML);	
		sHTML='</td></tr>';
		osHTML.push(sHTML);	
	}
	sHTML='</table> ';
	osHTML.push(sHTML);
	sHTML='</td><td valign="top" align="right" width="16">';
	osHTML.push(sHTML);	
	osHTML.push(sIcon);	
	sHTML='</td></tr></table>';
	osHTML.push(sHTML);
	
	//DPH 28/11/2002 oLocation.all sometimes empty, check is OK before using
	if((oRadio.oLocation.all[oRadio.sID+"_dyn"]!=undefined)&&(oRadio.oLocation.all[oRadio.sID+"_dyn"]!=null))
	{
		oRadio.oLocation.all[oRadio.sID+"_dyn"].innerHTML=osHTML.join('');
	}
	delete osHTML;
}

//
// Function to shift the focus from the visible (master) control to the invisible one
//
function fnFocusOnDummy(oRadio)
{
//	setTimeout("fnSetDummyFocus("+oRadio.sLocation+".oRadio,true);",nFocusDelay);
	fnSetDummyFocus(oRadio,true);
}

//
// Function to set the tooltip on the visible DIV
//
function fnSetRadioTooltip(oRadio,sText)
{
	if((oRadio.oLocation.all[oRadio.sID+"_dyn"]!=undefined)&&(oRadio.oLocation.all[oRadio.sID+"_dyn"]!=null))
	{
		oRadio.oLocation.all[oRadio.sID+"_dyn"].title=sText;
	}
}