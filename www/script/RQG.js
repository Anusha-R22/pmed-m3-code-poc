////////////////////
////////////////////
// Public methods
////////////////////
////////////////////

//
// Create a Repeating Question Group object.
//
//	oLocation=object *containing* the "DIV" to hold the RQG;
//	sLocation=DOM textual reference to the above object;
//	sID=identification string for this control - name of the "DIV";
//	bEnabled=enabled status;
//	nQGroupId=QGroupId;
//	nInitRows=initial number of rows;
//	nDisplayRows=number of display rows;
//	nMinRepeats=minimum number of repeats;
//	nMaxRepeats=maximum number of repeats;
//	bBorder=RQG border flag;
//	lMaxHeight=Maximum height RQG DIV can be
//	lTabIndex=Tab Index start number
//
// ic 13/01/2004 added mandatory flag
function fnCreateRQG(oLocation,sLocation,sID,bEnabled,nQGroupId,
					nInitRows,nDisplayRows,nMinRepeats,nMaxRepeats,bBorder,
					lMaxHeight,bMandatory,lTabIndex)
{
	if(oForm.olQGroups==null)
	{
		oForm.olQGroups=new Array();
	}

	if(oForm.olQGroups[sID]==null)
	{
		oForm.olQGroups[sID]=new Object();
	}

	var	oRQG=oForm.olQGroups[sID];
	sgDIV=sID+"_RQGDiv";
	oRQG.oLocation=oLocation.all[sgDIV];
	oRQG.sLocation=sLocation+".all."+sgDIV;
	oRQG.sID=sID;
	oRQG.nQGroupId=nQGroupId;
	oRQG.nInitRows=nInitRows;
	oRQG.nDisplayRows=nDisplayRows;
	oRQG.nMinRepeats=nMinRepeats;
	oRQG.nMaxRepeats=nMaxRepeats;
	oRQG.bBorder=bBorder;
	oRQG.lMaxHeight=lMaxHeight;
	oRQG.bMandatory=bMandatory;
	oRQG.lTabIndex=lTabIndex;
	sgDIV=sID+"_RQGHeadDiv";
	oRQG.oHeadingLoc=oLocation.all[sgDIV];

//
// Assign the object's methods - see functions called for detailed descriptions
//
	oRQG.addquestion=function(sQuestionID,nOrder)
	{
		return fnAddToRQG(this,sQuestionID,nOrder);
	}
	
	oRQG.redraw=function()
	{
		return fnRedrawRQG(this);
	}

	oRQG.redrawheaders=function()
	{
		return fnRedrawRQGHeaders(this);
	}
	
	oRQG.sethandles=function()
	{
		return fnSetHandles(this);
	}
	
	oRQG.enablefields=function()
	{
		return fnEnableFields(this);
	}

	oRQG.createnewrow=function()
	{
		return fnNewRQGRow(this);
	}
	
	oRQG.checkfornewrow=function()
	{
		return fnRQGNeedsNewRow(this);
	}
	
	oRQG.setscroll=function()
	{
		return fnSetRQGScrollFunction(this);
	}
	
	oRQG.resizeRQGDIV=function()
	{
		return fnResizeRQGDIV(this);
	}
	
	oRQG.setallskips=function(sExpression)
	{
		return fnSetRQGSkip(this,sExpression);
	}

	return oRQG;
}

//
// Create a Repeating Question Group object.
//
//	oRQG=RQG object to have question added to it;
//	sQuestionID=sID of the question added to RQG;
//	nOrder=Order of question within RQG;
function fnAddToRQG(oRQG,sQuestionID,nOrder)
{
	if (oRQG.slQuestion==null)
	{
		oRQG.slQuestion=new Array();
	}
	oRQG.slQuestion[nOrder]=sQuestionID;
}

// Other Functions
// Grow
// 

function fnRedrawRQG(oRQG)
{
	// DPH 09/01/2002 - Changed to use array then join...
	var sHTML='';
	var bRadioExists=false;
	var sCatString='';
	var sBackColour=oForm.sDisabledColour;
	var lTabIndex=oRQG.lTabIndex;
	var sFIDCommon='';
	var sOptionValue='';
	var osHTML=new Array();
	var nRQGQs=oRQG.slQuestion.length;
	
	osHTML.push('<table border=');
	//sHTML='<table border=';
	//border
	if(oRQG.bBorder==true)
	{
		// make sure RQG is not all hidden
		if(!fnAllHiddenRQG(oRQG))
		{
			osHTML.push('1');
		}
		else
		{
			osHTML.push('0');
		}
	}
	else
	{
		osHTML.push('0');
	}
	osHTML.push(' id="'+oRQG.sID+'_bordertable" ');
	osHTML.push(' cellspacing=0 cellpadding=0><tr><td>');
	osHTML.push('<table cellspacing=0 cellpadding=2 ');
	osHTML.push(' cols=');
	//sHTML+=' id="'+oRQG.sID+'_bordertable" ';
	//sHTML+=' cellspacing=0 cellpadding=0><tr><td>';
	//sHTML+='<table cellspacing=0 cellpadding=2 '
	//sHTML+=' cols='
	//column no
	if (oRQG.slQuestion==null)
	{
		return false;
	}
	osHTML.push(nRQGQs);
	osHTML.push(' id="'+oRQG.sID+'_table">');
	//sHTML+=oRQG.slQuestion.length;
	//sHTML+=' id="'+oRQG.sID+'_table">';
	// data rows
	nDispRows=oRQG.nDisplayRows;
	nInitRows=oRQG.nInitRows;
	nDataRows=oForm.olQuestion[oRQG.slQuestion[0]].olRepeat.length; // from 1st q
	if(nInitRows>nDataRows){nRowCount=nInitRows}else{nRowCount=nDataRows;}
	
	for(i=0;i<nRowCount;i++)
	{
		osHTML.push('<tr>');
		//sHTML+='<tr>';
		for(j=0;j<nRQGQs;j++)
		{
			var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[j]];
			var sFieldID=oFieldTemp.sID;
			if(!oFieldTemp.bHidden)
			{
				if(oFieldTemp.olRepeat[i]!=null)
				{
					var oFieldIns=oFieldTemp.olRepeat[i];
					sFIDCommon=sFieldID.substring(1,sFieldID.length);
					osHTML.push('<td>');
					//sHTML+='<td>'
					// Data type
					switch(oFieldTemp.nType)
					{
						case etText :
						case etIntegerNumber:
						case etRealNumber:
						case etDateTime:
						case etLabTest:
							// all basic text types
							//<input type="hidden" name="af_initials"><input type="text" disabled tabindex="2" name="f_initials" Style="background-color:#a9a9a9; " size="3" maxlength="3" onfocus="o1.fnGotFocus('f_initials');">&nbsp;
							// input + images in a table
							sHTML='<input type="hidden" name="a';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='"><table cellpadding=\'0\' cellspacing=\'0\'><tr><td onMouseOver="';
							osHTML.push(sHTML);
							sHTML='fnSetTooltip(\''+sFieldID+'\','+i+',1);">';
							osHTML.push(sHTML);						
							sHTML='<input type="text" disabled tabindex="';
							osHTML.push(sHTML);
							sHTML=lTabIndex;
							osHTML.push(sHTML);
							sHTML='" name="';
							osHTML.push(sHTML);
							sHTML=sFieldID
							osHTML.push(sHTML);
							sHTML='" Style="background-color:';
							osHTML.push(sHTML);
							osHTML.push(sBackColour);
							sHTML='; " size="';
							osHTML.push(sHTML);
							osHTML.push(oFieldTemp.nDisplayLength);
							sHTML='" maxlength="';
							osHTML.push(sHTML);
							osHTML.push(oFieldTemp.nLength);
							sHTML='" idx="';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML='" onfocus="o1.fnGotFocus(this);"';
							osHTML.push(sHTML);
							// set scroll if required for notes/comments
							if(oFieldTemp.nType==etText)
							{
								sHTML=' onscroll="o1.fnDisplayNoteStatusScroll(this);"';
								osHTML.push(sHTML);
								sHTML=' onchange="o1.fnDisplayNoteStatusScroll(this);"';
								osHTML.push(sHTML);
							}
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='</td>';
							osHTML.push(sHTML);
							sHTML='<a onMouseup=\'javascript:fnPopMenu(\"';
							osHTML.push(sHTML);						
							osHTML.push(sFieldID);						
							sHTML='\",';
							osHTML.push(sHTML);						
							osHTML.push(i);						
							sHTML=',event);\'>';
							osHTML.push(sHTML);
							sHTML='<td><img src="';
							osHTML.push(sHTML);
							osHTML.push(oTopImages["blank"].src);
							sHTML='" name="imgc';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='</td><td><table cellpadding=\'0\' cellspacing=\'0\'><tr><td>';
							osHTML.push(sHTML);
							sHTML='<img src="';
							osHTML.push(sHTML);
							osHTML.push(oTopImages["blank"].src);
							sHTML='" name="img';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='"';
							osHTML.push(sHTML);
							sHTML=' onMouseOver="fnSetTooltip(\''+sFieldID+'\','+i+',2);"';
							osHTML.push(sHTML);
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='</td></tr><tr><td>';
							osHTML.push(sHTML);
							sHTML='<img src="';
							osHTML.push(sHTML);
							osHTML.push(oTopImages["blank"].src);
							sHTML='" name="imgs';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='</td></tr></table>';
							osHTML.push(sHTML);
							sHTML='</td></a>';
							osHTML.push(sHTML);
							sHTML='<td valign="top" id="'+sFieldID+'_tdCTC"></td>';
							osHTML.push(sHTML);
							sHTML='</tr></table>';
							osHTML.push(sHTML);
							//store imagenames
							oFieldIns.sImageName='img'+sFIDCommon;
							oFieldIns.sImageCName='imgc'+sFIDCommon;
							oFieldIns.sImageSName='imgs'+sFIDCommon;
							break;
						case etMultimedia:
							//file
							//<input type="hidden" name="af_picture"><input type="file" disabled tabindex="5" name="f_picture" style="background-color:#a9a9a9; " onfocus="o1.fnGotFocus('f_picture');">&nbsp;<img src="../../img/blank.gif" name="imgc_picture">&nbsp;<img src="../../img/blank.gif" name="img_picture">
							sHTML='<input type="hidden" name="a';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='<input type="file" disabled tabindex="';
							osHTML.push(sHTML);
							osHTML.push(lTabIndex);
							sHTML='" name="';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='" idx="';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML='" style="background-color:#';
							osHTML.push(sHTML);
							osHTML.push(sBackColour);
							sHTML='; " onfocus="o1.fnGotFocus(this);"';
							osHTML.push(sHTML);
							sHTML=' onMouseOver="fnSetTooltip(\''+sFieldID+'\','+i+',1);"';
							osHTML.push(sHTML);
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='<a onMouseup=\'javascript:fnPopMenu(\"';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='\",';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML=',event);\'>';
							osHTML.push(sHTML);
							sHTML='<img src="../img/blank.gif" name="imgc';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">&nbsp;<img src="../img/blank.gif" name="img';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='"';
							osHTML.push(sHTML);
							sHTML=' onMouseOver="fnSetTooltip(\''+sFieldID+'\','+i+',2);"';
							osHTML.push(sHTML);
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='<img src="';
							osHTML.push(sHTML);
							osHTML.push(oTopImages["blank"].src);
							sHTML='" name="imgs';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='<div id="'+sFieldID+'_tdCTC"></div>';
							osHTML.push(sHTML);
							//store imagenames
							oFieldIns.sImageName='img'+sFIDCommon;
							oFieldIns.sImageCName='imgc'+sFIDCommon;
							oFieldIns.sImageSName='imgs'+sFIDCommon;
							break;
						case etCategory:
							//radio
							//<input type="hidden" name="af_sex">
							//<input type="hidden" name="f_sex">
							//<table><tr><td><div id="f_sex_inpDiv" style="position:relative;z-index:2;"></div></td><td valign="top">&nbsp;<img src="../img/blank.gif" name="imgc_sex">&nbsp;<img src="../../img/blank.gif" name="img_sex"></td></tr></table>
							sHTML='<input type="hidden" name="a';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='"><input type="hidden" name="';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='" idx="';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML='" ';
							osHTML.push(sHTML);
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='<table cellpadding=\'0\' cellspacing=\'0\'><tr><td>';
							osHTML.push(sHTML);
							sHTML='<div id="';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='_inpDiv" idx="';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML='" style="position:relative;z-index:2;"></div>';
							osHTML.push(sHTML);
							sHTML='</td><td valign="top">';
							osHTML.push(sHTML);
							sHTML='<img src="../img/blank.gif" name="imgc';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='</td><td valign="top"><table cellpadding=\'0\' cellspacing=\'0\'><tr>';
							osHTML.push(sHTML);
							sHTML='<a onMouseup=\'javascript:fnPopMenu(\"';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='\",';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML=',event);\'>';
							osHTML.push(sHTML);
							sHTML='<td>';
							osHTML.push(sHTML);
							sHTML='<img src="../img/blank.gif" name="img';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='"';
							osHTML.push(sHTML);
							sHTML=' onMouseOver="fnSetTooltip(\''+sFieldID+'\','+i+',2);"';
							osHTML.push(sHTML);
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='</td></a><tr><tr><td>';
							osHTML.push(sHTML);
							sHTML='<img src="../img/blank.gif" name="imgs';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='</td></tr></table></td>';
							osHTML.push(sHTML);
							sHTML='<td valign="top" id="'+sFieldID+'_tdCTC"></td>';
							osHTML.push(sHTML);
							sHTML='</tr></table>';
							osHTML.push(sHTML);
							bRadioExists=true;
							oFieldIns.sImageName='img'+sFIDCommon;
							oFieldIns.sImageCName='imgc'+sFIDCommon;
							oFieldIns.sImageSName='imgs'+sFIDCommon;
							break;
						case etCatSelect:
							//select
							//<input type="hidden" name="af_race"><table><tr><td id="tbl_race"><select tabindex="9" disabled name="f_race" Style="background-color:#a9a9a9; " onfocus="o1.fnGotFocus('f_race');" onchange="o1.fnLostFocus('f_race');"><option value=""> </option><option value="0">Caucasian/white</option><option value="1">black</option><option value="2">oriental</option><option value="3">other</option></select></td><td>&nbsp;<img src="../../img/blank.gif" name="imgn_race">&nbsp;<img src="../../img/blank.gif" name="imgc_race">&nbsp;<img src="../../img/blank.gif" name="img_race"></td></tr></table>
							sHTML='<input type="hidden" name="a';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='<table><tr><td id="tbl';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='" onMouseOver="fnSetTooltip(\''+sFieldID+'\','+i+',1);"';
							osHTML.push(sHTML);
							sHTML='><select tabindex="';
							osHTML.push(sHTML);
							osHTML.push(lTabIndex);
							sHTML='" disabled name="';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='" idx="';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML='" Style="background-color:';
							osHTML.push(sHTML);
							osHTML.push(sBackColour);
							sHTML='; " onfocus="o1.fnGotFocus(this);" onchange="o1.fnLostFocus(this);">';
							osHTML.push(sHTML);
							// Option values
							sHTML='<option value=""> </option>';
							osHTML.push(sHTML);
							for(var k in oFieldTemp.olCatValue)
							{ 
								sHTML='<option value="';
								osHTML.push(sHTML);
								osHTML.push(k);
								sHTML='">';
								osHTML.push(sHTML);
								osHTML.push(oFieldTemp.olCatValue[k].sCatText);
								sHTML='</option>';
								osHTML.push(sHTML);
							}
							sHTML='</select></td>';
							osHTML.push(sHTML);
							sHTML='<a onMouseup=\'javascript:fnPopMenu(\"';
							osHTML.push(sHTML);
							osHTML.push(sFieldID);
							sHTML='\",';
							osHTML.push(sHTML);
							osHTML.push(i);
							sHTML=',event);\'>';
							osHTML.push(sHTML);
							sHTML='<td><img src="../img/blank.gif" name="imgn';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='"><img src="../img/blank.gif" name="imgc';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='</td><td><table cellpadding=\'0\' cellspacing=\'0\'><tr><td>';
							osHTML.push(sHTML);
							sHTML='<img src="../img/blank.gif" name="img';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='"';
							osHTML.push(sHTML);
							sHTML=' onMouseOver="fnSetTooltip(\''+sFieldID+'\','+i+',2);"';
							osHTML.push(sHTML);
							sHTML='>';
							osHTML.push(sHTML);
							sHTML='</td></tr><tr><td>';
							osHTML.push(sHTML);
							sHTML='<img src="';
							osHTML.push(sHTML);
							osHTML.push(oTopImages["blank"].src);
							sHTML='" name="imgs';
							osHTML.push(sHTML);
							osHTML.push(sFIDCommon);
							sHTML='">';
							osHTML.push(sHTML);
							sHTML='</td></tr></table>';
							osHTML.push(sHTML);
							sHTML='</td></a>';
							osHTML.push(sHTML);
							sHTML='<td valign="top" id="'+sFieldID+'_tdCTC"></td>';
							osHTML.push(sHTML);
							sHTML='</tr></table>';
							osHTML.push(sHTML);
							//store imagenames
							oFieldIns.sImageName='img'+sFIDCommon;
							oFieldIns.sImageCName='imgc'+sFIDCommon;
							oFieldIns.sSelectNoteImage='imgn'+sFIDCommon;
							oFieldIns.sImageSName='imgs'+sFIDCommon;
							break;
						default :
					}
					sHTML+='</td>';
				}
			}
			else
			{
				sHTML='<input type="hidden" name="' + sFieldID + '">';
				osHTML.push(sHTML);
				sHTML='<input type="hidden" name="a' + sFieldID + '">';
				osHTML.push(sHTML);
			}
			lTabIndex++;
		}
		osHTML.push('</tr>');
	}
	osHTML.push('</table>');
	osHTML.push('</td></tr></table>');
	oRQG.oLocation.innerHTML=osHTML.join('');

	delete osHTML;

	// Initialise Radio Objects separately
	if(bRadioExists)
	{
		// Loop through RQG questions and create any radio questions
		for(i=0;i<nRowCount;i++)
		{
			// Reset Tabindex
			lTabIndex=(oRQG.lTabIndex+(i*nRQGQs));
			for(j=0;j<nRQGQs;j++)
			{
				var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[j]];
				var sFieldID=oFieldTemp.sID;
				lTabIndex++;
				
				// Data type
				if(oFieldTemp.nType==etCategory)
				{
					if(aRadioList==null)
					{
						aRadioList=new Array();
					}

					if(aRadioList[sFieldID]==null)
					{
						aRadioList[sFieldID]=new Object();
					}
	
					// Create Repeat Array if need be
					if(aRadioList[sFieldID].olRepeat==null)
					{
						aRadioList[sFieldID].olRepeat=new Array();
					}

					sCatString='';
					for(k in oFieldTemp.olCatValue)
					{
						if(sCatString.length>0)
						{
							sCatString+='~';
						}
						sCatString+=k+'¬'+oFieldTemp.olCatValue[k].sCatText;
					}

					if(oFieldTemp.olRepeat[i]!=null)
					{
						var oFieldIns=oFieldTemp.olRepeat[i];
						var bMultipleRadios=false;
						//function fnRadioCreate(oLocation,sLocation,sID,sValues,sInitialValue,
						//							nTabIndex,sValueRef,sColour,bRQG)
						//aRadioList["f_sex"] = fnRadioCreate(frm2,"window.parent.frames[1].document.deFrm","f_sex_inpDiv","1¬Female~2¬Male",null,4,"f_sex","FONT-SIZE: 8 pt; FONT-FAMILY: Tahoma;");
						//aRadioList["f_sex"].onfocus=function(){frm1.fnGotFocus('f_sex');};
						//aRadioList["f_sex"].onchange=function(){frm1.fnLostFocus('f_sex');};

						if(nRowCount>1){bMultipleRadios=true;}
						aRadioList[sFieldID].olRepeat[i] =new Object();
						//aRadioList[oFieldTemp.sID].olRepeat[i] = fnRadioCreate(o2,"window.document.FormDE",oFieldTemp.sID+'_inpDiv',sCatString,null,lTabIndex,oFieldTemp.sID,'FONT-SIZE: 8 pt; FONT-FAMILY: Tahoma;',i,bMultipleRadios);
						aRadioList[sFieldID].olRepeat[i] = fnRadioCreate(o2,"window.document.FormDE",sFieldID+'_inpDiv',sCatString,null,lTabIndex,sFieldID,oFieldTemp.sColour,oFieldTemp.sFontStyle,i,bMultipleRadios);
						if(nRowCount>1)
						{
							var oFieldLoc=eval('o2.document.all["'+sFieldID+'"]['+i+']');
						}
						else
						{
							var oFieldLoc=eval('o2.document.all["'+sFieldID+'"]');
						}
						aRadioList[sFieldID].olRepeat[i].oValueRef=oFieldLoc;
						aRadioList[sFieldID].olRepeat[i].onfocus=function(){o1.fnGotFocus(this);};
						//aRadioList[oFieldTemp.sID].olRepeat[i].onchange=function(){o1.fnLostFocus(oFieldLoc);};
						aRadioList[sFieldID].olRepeat[i].onchange=function(){o1.fnLostFocus(this);};
						// set oRadio handle
						oFieldIns.oRadio=aRadioList[sFieldID].olRepeat[i];
					}
				}
			}
		}
	}
	fnSetHandles(oRQG);
	return true;
}

// draw headings & write to DIV
function fnRedrawRQGHeaders(oRQG)
{
//	if (oRQG.oLocation.readyState != "complete" && oRQG.oLocation.readyState != 4)
//	{
//	    window.setTimeout("fnRedrawRQGHeaders(oRQG);", 250);
//	    return;
//	}

	var sHTML='';
	var nScrollOffset=0;
	var nOffset=0;
	var osHTML=new Array();
	var nHeadHeight=0;
	var nQBlockTop=0;
	var nCell=0;
	var nRQGQs=oRQG.slQuestion.length;
	
	if(oRQG.oLocation.scrollWidth>oRQG.oLocation.clientWidth)
	{
		nScrollOffset=oRQG.oLocation.offsetWidth-oRQG.oLocation.clientWidth;
	}
	
	sHTML='<table border=';
	osHTML.push(sHTML);
	//border
	if(oRQG.bBorder==true)
	{
		// make sure RQG is not all hidden
		if(!fnAllHiddenRQG(oRQG))
		{
			osHTML.push('1');
		}
		else
		{
			osHTML.push('0');
		}
	}
	else
	{
		osHTML.push('0');
	}
	sHTML=' width=';
	osHTML.push(sHTML);
	osHTML.push(((oRQG.oLocation.all(oRQG.sID+"_bordertable").clientWidth)));
	sHTML=' cellspacing=0 cellpadding=0><tr><td>';
	osHTML.push(sHTML);
	sHTML='<table cellspacing=0 cellpadding=2';
	osHTML.push(sHTML);
	sHTML=' cols=';
	osHTML.push(sHTML);
	//column no
	if (oRQG.slQuestion==null)
	{
		return false;
	}
	osHTML.push(nRQGQs);
	sHTML=' id="';
	osHTML.push(sHTML);
	osHTML.push(oRQG.sID);
	sHTML='_tablehead"';
	osHTML.push(sHTML);
	sHTML=' width=';
	osHTML.push(sHTML);
	osHTML.push(((oRQG.oLocation.all(oRQG.sID+"_table").clientWidth)));
	sHTML='>';
	osHTML.push(sHTML);
	//column headings
	sHTML='<tr>';
	osHTML.push(sHTML);
	for(i=0;i<nRQGQs;i++)
	{
		var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[i]];
		if(!oFieldTemp.bHidden)
		{
			sHTML='<td valign="top" width=';
			osHTML.push(sHTML);
			osHTML.push(((oRQG.oLocation.all(oRQG.sID+"_table").cells[nCell].clientWidth)-4));
			sHTML='px';
			osHTML.push(sHTML);
			sHTML='>';
			osHTML.push(sHTML);
			osHTML.push(oFieldTemp.sCaptionText);
			sHTML='</td>';
			osHTML.push(sHTML);
			nCell++;
		}
	}
	sHTML='</tr>';
	osHTML.push(sHTML);
	sHTML='</table>';
	osHTML.push(sHTML);
	sHTML='</td></tr></table>';
	osHTML.push(sHTML);
	oRQG.oHeadingLoc.innerHTML=osHTML.join('');
	delete osHTML;
	oRQG.oHeadingLoc.style.width=oRQG.oLocation.style.width;
	
	// attempt to resize header height if necessary & move question DIV (+4 cellpadding)
	nHeadHeight=oRQG.oHeadingLoc.all(oRQG.sID+"_tablehead").clientHeight+4;
	// if height > start value
	if(nHeadHeight>25)
	{
		oRQG.oHeadingLoc.style.height=nHeadHeight;
		// calculate amount to 'move down' RQG question DIV
		nQBlockTop=oRQG.oHeadingLoc.style.pixelTop+nHeadHeight;
		oRQG.oLocation.style.pixelTop=nQBlockTop;
	}
}

// Update Handles to Questions
function fnSetHandles(oRQG)
{
	// data rows
	nDispRows=oRQG.nDisplayRows;
	nInitRows=oRQG.nInitRows;
	nDataRows=oForm.olQuestion[oRQG.slQuestion[0]].olRepeat.length; // from 1st q
	if(nInitRows>nDataRows){nRowCount=nInitRows}else{nRowCount=nDataRows;}
	var nRQGQs=oRQG.slQuestion.length;

	// Loop through RQG questions and associate handles
	for(i=0;i<nRowCount;i++)
	{
		for(j=0;j<nRQGQs;j++)
		{
			var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[j]];
			if(oFieldTemp.olRepeat[i]!=null)
			{
				var oFieldIns=oFieldTemp.olRepeat[i];
				if(oFieldTemp.olRepeat.length>1)
				{
					// use array
					oFieldIns.oHandle=(o2[oFieldTemp.sID])[i];
				}
				else
				{
					// use object directly
					oFieldIns.oHandle=o2[oFieldTemp.sID];
				}
				// Set hidden field value
				//o2["a"+oFieldTemp.sID].value=(oFieldIns.AIValue!=undefined)?oFieldIns.AIValue:"";
				// Set additional Field Info Property
				if(oFieldTemp.olRepeat.length>1)
				{
					(o2["a"+oFieldTemp.sID][i]).value=(oFieldIns.AIValue!=undefined)?oFieldIns.AIValue:"";
					o1.fnFd(oFieldTemp.sID,"oAIHandle",o2["a"+oFieldTemp.sID][i],i);
				}
				else
				{
					o2["a"+oFieldTemp.sID].value=(oFieldIns.AIValue!=undefined)?oFieldIns.AIValue:"";
					o1.fnFd(oFieldTemp.sID,"oAIHandle",o2["a"+oFieldTemp.sID],i);
				}
			}
		}
	}	
}

function fnEnableFields(oRQG)
{
	// data rows
	nDispRows=oRQG.nDisplayRows;
	nInitRows=oRQG.nInitRows;
	nDataRows=oForm.olQuestion[oRQG.slQuestion[0]].olRepeat.length; // from 1st q
	if(nInitRows>nDataRows){nRowCount=nInitRows}else{nRowCount=nDataRows;}
	var nRQGQs=oRQG.slQuestion.length;

	// Loop through RQG questions and enable fields
	for(i=0;i<nRowCount;i++)
	{
		for(j=0;j<nRQGQs;j++)
		{
			var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[j]];
			if((!oFieldTemp.bHidden)&&(oFieldTemp.olRepeat[i]!=null))
			{
				var oFieldIns=oFieldTemp.olRepeat[i];
				setJSValue(oFieldTemp.sID,oFieldIns.getFormatted(),true,i);
				if(oFieldIns.bEnabled)
				{
					setFieldEnabled(oFieldTemp.sID,i,true);
				}
				// Draw Normal Range / CTC
				fnDrawNRCTC(oFieldTemp.sID,i);
			}
		}
	}
}

function fnNewRQGRow(oRQG)
{
	// check if should be adding row
	var bAddRow=oRQG.checkfornewrow();
	if(bAddRow)
	{
		// add new question response to list then redraw
		// checking if less than max firstly
		var bRedraw=false;
		var nRQGQs=oRQG.slQuestion.length;
		for(i=0;i<nRQGQs;i++)
		{
			var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[i]];
			var nCurrentRows=oFieldTemp.olRepeat.length;
			if(nCurrentRows<oRQG.nMaxRepeats)
			{
				bRedraw=true;
				// create new instance
				//function fnCreateFieldInstance(sFieldID,nRepeatNo,vValue,bEnabled,nStatus,
				//		oHandle,oCapHandle,sImageName,nLockStatus,
				//		nDiscrepancyStatus,nSDVStatus,bNote,bComment,oRadio,
				//		sImageSName,sSelectNoteImage,nChanges,sImageCName,
				//		sComments,sRFC,sUserFull,sNRCTC,bReVal,sRFO)

				fnCI(oFieldTemp.sID,(nCurrentRows+1),"",true,-10,
										null,null,"",0,
										0,0,false,false,null,
										"","",0,"",
										"","","","",false,"");
			}
		}

		if(bRedraw==true)
		{
			// now redraw
			oRQG.redraw();
				
			// enable fields
			oRQG.enablefields();
			
			// draw icons
			fnApplyRulesForRQG(oRQG.sID);

			// set scroll
			oRQG.setscroll();
			
			// Resize DIV
			oRQG.resizeRQGDIV();
					
			// redraw headings
			oRQG.redrawheaders();
			
		}
	}
	return bAddRow;
}

function fnResizeRQGDIV(oRQG)
{
	// DPH 27/02/2004 - check that all RQG questions are not all hidden
	// if they are there is no need to resize
	if(!fnAllHiddenRQG(oRQG))
	{
		// firstly resize columns
		fnResizeColumnsByFactor(oRQG);
		var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[0]];
		var bWideRQG=false;
		// changed order to calculate width firstly
		// Sort out width by getting width of border table and adding size of scrollbar
		// tablewidth is actual size of table
		// oRQG.oLocation.clientWidth is displayable area size
		// calculate if is 'wide' rqg needing scrollbar
		var lContentWidth=oRQG.oLocation.all(oRQG.sID+"_bordertable").clientWidth;
		if(lContentWidth>oRQG.oLocation.clientWidth)
		{
			// table is wider than screen display area - store bWideRQG
			bWideRQG=true;
		}
		var lTableWidth=(oRQG.oLocation.all(oRQG.sID+"_bordertable").clientWidth)+18;
		//if(lTableWidth<oRQG.oLocation.clientWidth)
		if(lTableWidth<oRQG.oLocation.offsetWidth)
		{
			oRQG.oLocation.style.width=lTableWidth;
		}
	
		// Get height of first question element of table
		var lRowHeight=(oRQG.oLocation.all(oRQG.sID+"_table").cells[0].clientHeight);
		// plus padding per row
		// pad if not a wide RQG
		if(!bWideRQG)
		{
			lRowHeight+=4;
		}
		// multiply by no of display rows
		var lDispRowHeight=lRowHeight*oRQG.nDisplayRows;
		// no of data rows
		var lDataRowHeight=lRowHeight*oFieldTemp.olRepeat.length;
		if(lDataRowHeight<lDispRowHeight)
		{
			lRowHeight=lDataRowHeight;
		}
		else
		{
			lRowHeight=lDispRowHeight;
		}
		// if a wide RQG need horizontal scrollbars so add 20px to height
		if(bWideRQG)
		{
			lRowHeight+=20;
		}
		// make sure RQG height does not exceed maximum
		if(lRowHeight>oRQG.lMaxHeight)
		{
			lRowHeight=oRQG.lMaxHeight;
		}
		oRQG.oLocation.style.height=lRowHeight;
	}
	// move eForm next button...
	fnSetBtnNextPos();
}

// resize columns by a set factor (1.1) to attempt to match windows
function fnResizeColumnsByFactor(oRQG)
{
	var nFactor=1.1;
	// resize only first row of cells as others should resize to same width!?!
	var nRQGQs=oRQG.slQuestion.length;
	var nCellsToResize=0;
	var lWidth=0;
	// loop through cells to calc cellsno to resize
	for(var nI=0;nI<nRQGQs;nI++)
	{
		// check question is not hidden
		if(!oForm.olQuestion[oRQG.slQuestion[nI]].bHidden)
		{
			nCellsToResize++;
		}	
	}
	
	// loop through cells and multiply width by factor
	for(var nI=0;nI<nCellsToResize;nI++)
	{
		lWidth=(oRQG.oLocation.all(oRQG.sID+"_table").cells[nI].clientWidth)*nFactor;
		oRQG.oLocation.all(oRQG.sID+"_table").cells[nI].style.width=lWidth;
	}
}

// check if RQG needs a new row
function fnRQGNeedsNewRow(oRQG)
{
	//if already max rows quit
	// get 1st question name & check its max repeat
	var sQName=oRQG.slQuestion[0];
	var nDataRows=oForm.olQuestion[sQName].olRepeat.length; // from 1st q
	var nInitRows=oRQG.nInitRows;
	var bNewRow=true;
	var bDeriv=false;
	var bBlank;
	var bComplete;
	var bMand;
	var bMandEmpty;
	var bThisMandOK;
	if(nDataRows==oRQG.nMaxRepeats)
	{
		//already max repeats
		return false;
	}

	var nRQGQs=oRQG.slQuestion.length;

	// Loop through RQG questions and enable fields
	for(i=0;i<nDataRows;i++)
	{
		bBlank=true;
		bComplete=false;
		bMandEmpty=false;
		for(j=0;j<nRQGQs;j++)
		{
			var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[j]];
			if(!oFieldTemp.bHidden)
			{
				// mark if value is derived
				bDeriv=false;
				if(oFieldTemp.olDerivation!=null)
				{
					if(oFieldTemp.olDerivation.length>0)
					{
						bDeriv=true;
					}
				}
				// Mark if mandatory
				bMand=oFieldTemp.bMandatory;
				if(oFieldTemp.olRepeat[i]!=null)
				{
					var oFieldIns=oFieldTemp.olRepeat[i];
					// Check if Blank (not derived and "")
					if(!bDeriv)
					{
						if(oFieldIns.getFormatted()!="")
						{
							bBlank=false;
						}
					}
					// if question mandatory must have value
					if(bMand)
					{
						bThisMandOK=false;
						if(oFieldIns.getFormatted()!="")
						{
							bComplete=true;
							bThisMandOK=true;
						}
						// check status
						switch(getFieldStatus(oFieldTemp.sID,i))
						{
							case 10:
							case -5:
							case -8:
							case -10:
								break
							default:
								{
									bComplete=true;
									bThisMandOK=true;
								}
						}
						if(!bThisMandOK)
						{
							bMandEmpty=true;
						}
					}
					else
					{
						bComplete=true;
					}
				}
			}
		}
		// if whole row not blank / complete then OK for new row
		if((bBlank)||(!bComplete)||(bMandEmpty))
		{
			bNewRow=false;
		}
	}
	return bNewRow;
}

function fnSetRQGScrollFunction(oRQG)
{
//	if(oRQG.oLocation.scrollWidth>oRQG.oLocation.clientWidth)
//	{
		addScrollSynchronization(oRQG.oHeadingLoc,oRQG.oLocation,"horizontal");
//	}
//	else
//	{
//		removeScrollSynchronization(oRQG.oHeadingLoc);
//	}
}

// This is a function that returns a function that is used
// in the event listener
function getOnScrollFunction(oElement) {
	return function () {
		if (oElement._scrollSyncDirection == "horizontal" || oElement._scrollSyncDirection == "both")
			oElement.scrollLeft = event.srcElement.scrollLeft;
		if (oElement._scrollSyncDirection == "vertical" || oElement._scrollSyncDirection == "both")
			oElement.scrollTop = event.srcElement.scrollTop;
	};

}
// This function adds scroll syncronization for the fromElement to the toElement
// this means that the fromElement will be updated when the toElement is scrolled
function addScrollSynchronization(fromElement, toElement, direction) {
	removeScrollSynchronization(fromElement);
	
	fromElement._syncScroll = getOnScrollFunction(fromElement);
	fromElement._scrollSyncDirection = direction;
	fromElement._syncTo = toElement;
	toElement.attachEvent("onscroll", fromElement._syncScroll);
}

// removes the scroll synchronization for an element
function removeScrollSynchronization(fromElement) {
	if (fromElement._syncTo != null)
		fromElement._syncTo.detachEvent("onscroll", fromElement._syncScroll);

	fromElement._syncTo = null;
	fromElement._syncScroll = null;
	fromElement._scrollSyncDirection = null;
}

// TODO
// functions to:
// show values
// detect new row
//
// Checks RQG for a new row
function fnRQGNewRowCheck(sFieldID)
{
	// Get RQG object & call function to check / add row
	var oFieldTemp=oForm.olQuestion[sFieldID];
	if(oFieldTemp.bRQG)
	{
		var oRQG=aRQG[oFieldTemp.sRQG];
		var bNewRQGRow=oRQG.createnewrow();
		if(bNewRQGRow)
		{
			return true;
		}
	}
	return false;
}

// Set skip condition on all questions within RQG
function fnSetRQGSkip(oRQG,sExpression)
{
	// loop through questions
	var nRQGQs=oRQG.slQuestion.length;
	for(i=0;i<nRQGQs;i++)
	{
		var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[i]];
		// set condition
		if(oFieldTemp.olSkip==null)
		{
			oFieldTemp.olSkip=new Array();
		}
		nSkipCount=oFieldTemp.olSkip.length;
		oFieldTemp.olSkip[nSkipCount]=new Object();
		oFieldTemp.olSkip[nSkipCount].nOrder=nSkipCount;
		oFieldTemp.olSkip[nSkipCount].sExpression=sExpression;
		// Add the field to the relevant dependency lists
		fnCreateDependencies(oFieldTemp.sID,"S",sExpression);
	}
}

// check if RQG consists of all 'hidden' fields
function fnAllHiddenRQG(oRQG)
{
	var bHidden=true;
	var nRQGQs=oRQG.slQuestion.length;
	// loop through all fields
	for(i=0;i<nRQGQs;i++)
	{
		var	oFieldTemp=oForm.olQuestion[oRQG.slQuestion[i]];
		if(!oFieldTemp.bHidden)
		{
			bHidden=false;
		}
	}
	return bHidden;
}