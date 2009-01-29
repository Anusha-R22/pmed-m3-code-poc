///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2004
//
//	functions for handling eform window
//
///////////////////////////////////////////////////////////////////////////////////////////
var sDel1 = "`"; //major delimiter
var sDel2 = "|"; //minor delimiter
var sDel3 = "~";
var oEP; //eform permissions object
var o1;	//validation engine frame
var o2; //de form object
var o3; //div object list
var o4; //de frame
var s1; //de form (string)
var aRadioList=new Array(); //array of radios on eform
var aRQG=new Array(); // array of RQGs on eform
var oDepWin=null;
var beFormChanged=false; // boolean to track any change to eform data
var bDerivChanged=false; // check if derivations changed on eform
var bBtnNext=false;  // on button next
var bEformLoadCheck=false; //gets set to true when pageloaded() has completed
var bErrorReported=false; //has an error been sent to the server

function fnKeyDown()
{
	var bUseEnter=true;
	if (((event.keyCode<112)||(event.keyCode>123))&&(event.keyCode!=13)) return;
	
	if (fnLoading()) 
	{
		event.returnValue=false;
		event.keyCode=0;
		return;
	}
	
	switch (event.keyCode)
	{
		case 13:
			// enter
			if ((oCurrentFieldID!=null)&&(!bBtnNext))
			{
				//set focus target to current field (to stop multiple validations)
				oFocusTarget=oCurrentFieldID;
				// check current field
				//fnCheckCurrentField();
				//bUseEnter=false;
				// tab to next field - windows behaviour
				event.returnValue=true;
				event.keyCode=9; //tab
				return;
			}
			if(bUseEnter)
			{
				// do normal enter keystroke
				event.returnValue=true;
				event.keyCode=13;
				return;
			}
			break;
		case 114:
			//f3 previous
			fnSave("m1");break;
		case 115:
			//f4 next
			fnSave("m2");break;
		case 116:
			//f5 print
			fnPrint();break;
		case 117:
			//f6 return
			fnSave("m3");break;
		case 118:
			//f7 save
			fnSave("m0");break;
		case 119:
			//f8 cancel
			fnCancel();break;
		case 120:
			//f9 clear
			if (oCurrentFieldID!=null)
			{
				var bEnterable=fnEnterable(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx));
				if (oEP.bChangeData&&bEnterable)
				{
					fnPopMenuClick(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx),0)
				}
			}
			break;
		case 121:
			//f10 question menu
			if (oCurrentFieldID!=null)
			{
				fnPopMenu(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx),undefined,((window.document.body.clientWidth/3)+document.body.scrollLeft),((window.document.body.clientHeight/3)+document.body.scrollTop));
			}
			break;
		case 122:
			//f11 add comment
			if (oCurrentFieldID!=null)
			{
				var bHasResponse=fnHasResponse(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx));
				var bLockedOrFrozen=fnLockedOrFrozen(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx));
				var bEformReadOnly=fnFieldEformIsReadOnly(oCurrentFieldID.name)
				if (oEP.bAddComment&&bHasResponse&&(!bLockedOrFrozen)&&(!bEformReadOnly)&&oEP.bChangeData)
				{
					fnPopMenuClick(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx),1)
				}
				else
				{
					alert("You cannot add comments to this field")
				}
			}
			break;
		case 123:
			//f12 delete comments
			if (oCurrentFieldID!=null)
			{
				var bHasResponse=fnHasResponse(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx));
				var bLockedOrFrozen=fnLockedOrFrozen(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx));
				var bEformReadOnly=fnFieldEformIsReadOnly(oCurrentFieldID.name)
				if (oEP.bAddComment&&bHasResponse&&(!bLockedOrFrozen)&&(!bEformReadOnly)&&oEP.bChangeData)
				{
					fnPopMenuClick(oCurrentFieldID.name,DefaultRepeatNo(oCurrentFieldID.idx),2)
				}
				else
				{
					alert("You cannot delete comments from this field")
				}
			}
			break;
		default:
	}
	event.returnValue=false;
	event.keyCode=0;
}
//function moves back to previous page state
function fnBack()
{
	//if there is a backstate and the app isnt in split mode
	//only problem here is that if in split mode, cant use this function to go back to schedule
	if ((window.parent.window.parent.sBackState!=undefined)&&(window.parent.window.parent.sBackState!="")&&(!window.parent.window.parent.fnIsSplit()))
	{
		fnSave('m6','fnSetWinState("5|'+window.parent.window.parent.sBackState+'")');
	}
	else
	{
		fnSave("m1");
	}
}
function fnEformIsReadOnly()
{
	if ((o2==null)||(o2==undefined)) return true;
	if ((o2.readonly==null)||(o2.readonly==undefined)) return true;
	var sRW=o2.readonly.value;
	var aRW=sRW.split(sDel1);
	
	//if ((visit eform is writeable) or (user eform is writeable))
	if((aRW[0]==0)||(aRW[1]==0))
	{
		//data is writeable
		return false;
	}
	else
	{
		//data is read-only
		return true;
	}
}
//eform is readonly OR user doesnt have change data rights
function fnDataIsReadOnly()
{
	return (fnEformIsReadOnly()||(oEP.bChangeData!=1));
}
function fnCheckCurrentField()
{
	return fnLeaveFieldExceptCats(oCurrentFieldID);
}
function fnCancel()
{
	if(fnEformIsReadOnly())
	{
		//eform was opened readonly, 
		//no need to check for changes or realease db locks
		if ((window.parent.window.parent.sCancelState!=undefined)&&(window.parent.window.parent.sCancelState!="")&&(!window.parent.window.parent.fnIsSplit()))
		{
			eval("window.parent.window.parent.fnSetWinState('5|"+window.parent.window.parent.sCancelState+"')");
		}
		else
		{
			eval("window.parent.window.parent."+fnGetScheduleUrl());
		}
	}
	else
	{
		//eform was opened writeable,
		//if the user has made changes, confirm they will be lost
		//note: a user without change data rights can still save mimessages etc
		if(fnUserHasChangedEForm())
		{
			if (!confirm("The eForm has changed, are you sure you wish to close without saving?"))
			{
				return;
			}
		}
		
		//return to cancelstate
		if ((window.parent.window.parent.sCancelState!=undefined)&&(window.parent.window.parent.sCancelState!="")&&(!window.parent.window.parent.fnIsSplit()))
		{
			fnSave('m4','fnSetWinState("5|'+window.parent.window.parent.sCancelState+'")');
		}
		else
		{
			//failsafe, return to schedule
			fnSave('m4',fnGetScheduleUrl());
		}
	}
}
function fnGetScheduleUrl()
{
	var sSi=fnGetFormProperty("sSite");
	var sSt=fnGetFormProperty("sStudyId");
	var sSj=fnGetFormProperty("sSubject");
	return 'fnScheduleUrl("'+sSt+'","'+sSi+'","'+sSj+'")';
}
function fnConfirmGoto(jsfn)
{
	if(fnEformIsReadOnly())
	{
		//eform was opened readonly, no need to realease db locks
		eval("window.parent.window.parent."+jsfn);
	}
	else
	{
		//eform was opened writeable,
		if(fnUserHasChangedEForm())
		{
			//if the user has made changes, confirm save or cancel
			//note: a user without change data rights can still save mimessages etc
			if (confirm("Save eForm before continuing? Click OK to save, Cancel to continue without saving"))
			{
				fnSave("m3",jsfn);
			}
			else
			{
				fnSave("m4",jsfn);
			}
		}
		else
		{
			fnSave("m4",jsfn);
		}
	}
}
//m0=save&same, m1=save&prev, m2=save&next, m3=save&return, m4=cancel&return, m5=autoload, m6=save&back
function fnSave(id,jsfn)
{
	var nxt="";
	var nSave=0;
	var bFieldOk=true;
	var bPageOk=true;
	
	
	//null lastfocus variable as otherwise a gotfocus() can fire during or shortly 
	//after the following validation process resulting in multiple rfc dialogs. set 
	//it back if we exit without saving
	var oTempLastFocusID=oLastFocusID;
	oLastFocusID=null
	
	if(id!="m0")
	{
		//set parent window 'back' variable to this eform
		window.parent.window.parent.sBackState=window.parent.sWinState;
	}
	
	//set bOK = (id not equal to cancel & data is writeable) ? result of page revalidation check : true
	if(id!="m4")
	{
		bFieldOk=fnCheckCurrentField();
		if(bFieldOk)
		{
			bPageOk=fnRevalidatePage(false);
		}
	}
			
	if(bFieldOk&&bPageOk)
	{
		//this condition calculates whether to attempt to save the eform or not
		//it sets the nSave variable to '1' or leaves it at '0'. this will be 
		//submitted to the IO layer as part of the 'next' form field value 
		//constructed later in the function
		if((id!="m4")&&(!fnEformIsReadOnly()))
		{
			if(window.VisitDateOk())
			{
				if(window.FormDateOk())
				{
					if(oEP.bChangeData==1)
					{
						//user can change data
						if(fnUserHasChangedEForm()||fnDerivationHasChangedEForm()||fnNewEForm()||(id=="m0")||(id=="m3"))
						{
							//(form and visit date are ok) - either entered or not required
							//(the eform has changed by user action or derivation) or (this is a new eform)
							//or (user explicitly requested save) : SAVE
							nSave=1;
						}
					}
					else
					{
						//user cannot change data
						if(fnUserHasChangedEForm())
						{
							//(form and visit date are ok) - either entered or not required
							//(the eform has changed by user action: note/discrepancy/sdv) : SAVE
							nSave=1;
						}
					}
				}
				else
				{
					if(fnUserHasChangedEForm()||(id=="m0")||(id=="m3"))
					{
						//(form date is rejected) and (user has changed something or explicitly requested save)
						//alert the user to the missing form date, return
						alert("The eForm cannot be saved because no eForm date has been specified.");
						//set lastfocus back
						oLastFocusID=oTempLastFocusID;
						return;
					}
				}
			}
			else
			{
				if(fnUserHasChangedEForm()||(id=="m0")||(id=="m3"))
				{
					//(visit date is rejected) and (user has changed something or explicitly requested save)
					//alert the user to the missing visit date, return
					alert("The eForm cannot be saved because no visit date has been specified.");
					//set lastfocus back
					oLastFocusID=oTempLastFocusID;
					return;
				}
			}
		}
		
		//if explicit save was requested, but a save is not allowed, return
		//this might happen if a user hits f7 on a readonly eform
		if((id=="m0")&&(nSave==0)) 
		{
			//set lastfocus back
			oLastFocusID=oTempLastFocusID;
			return;
		}
		
		if((id=="m3")||(id=="m6"))
		{
			//set to return
			id="m3";
			
			//if save&return, return to cancelstate
			//this gets last visited page from parent window so can return to it
			//this last page state is concatenated with nxt and a delimiter
			if (jsfn!=undefined)
			{
				nxt=id+sDel3+jsfn;
			}
			else if ((window.parent.window.parent.sCancelState!=undefined)&&(window.parent.window.parent.sCancelState!="")&&(!window.parent.window.parent.fnIsSplit()))
			{
				nxt=id+sDel3+'fnSetWinState("5|'+window.parent.window.parent.sCancelState+'")';
			}
			else
			{
				//failsafe, return to schedule
				nxt=id+sDel3+fnGetScheduleUrl();
			}
		}
		else
		{
			//if jsfn is not passed, make it an empty string
			//concatenate it with menu id
			nxt=id+sDel3+((jsfn==undefined)?"":jsfn);
		}
		//concatenate the save flag
		nxt+=sDel3+nSave;
		
	
		//set next field value of form
		document.FormDE["next"].value=nxt;
		
		if(nSave==1)
		{
			
			if(o1.fnFinalisePage(window,o2,(id=="m0"),(id=="m3")))
			{
				//show saver,hide clickable buttons
				fnShowLoader("Contacting server, uploading data");
				fnDisableEformVisit();							
				o2.submit();
				return;
			}
		}
		else
		{
			//show saver,hide clickable buttons
			fnShowLoader("Contacting server, clearing locks");
			fnDisableEformVisit();
			o2.submit();
			return;
		}
	}
	//set lastfocus back
	oLastFocusID=oTempLastFocusID;
}
function fnPrint()
{
	//null lastfocus variable as otherwise a gotfocus() can fire during or shortly 
	//after the following validation process resulting in multiple rfc dialogs. set 
	//it back when we exit
	var oTempLastFocusID=oLastFocusID;
	oLastFocusID=null

	//validate current field, if any
	if (fnCheckCurrentField())
	{
		//check for eform changes. these must be saved before printing
		if (fnUserHasChangedEForm()||fnDerivationHasChangedEForm())
		{
			if(confirm("Unsaved responses on the eForm must be saved before printing."))
			{
				fnSave("m0");
			}
			//set lastfocus back
			oLastFocusID=oTempLastFocusID;
			//user doesnt want to save changes
			return;
		}
		//no changes to save - allow print
		var sSi=fnGetFormProperty("sSite");
		var sSt=fnGetFormProperty("sStudyId");
		var sSj=fnGetFormProperty("sSubject");
		var sVi=fnGetFormProperty("sVisitID");
		var sVId=fnGetFormProperty("sVEformID");
		var sVTId=fnGetFormProperty("sVEformPageTaskID");
		sVId=(sVId==null)? "":sVId;
		sVTId=(sVTId==null)? "":sVTId;
		var sId=fnGetFormProperty("sEformID");
		var sTId=fnGetFormProperty("sEformPageTaskID");
		var	sResult=window.showModalDialog("DialogPrintEform.asp?fltSi="+sSi+"&fltSt="+sSt+"&fltSj="+sSj+"&fltVi="+sVi+"&fltVId="+sVId+"&fltVTId="+sVTId+"&fltId="+sId+"&fltTId="+sTId,"","dialogHeight:450px; dialogWidth:500px; center:yes; resizable:1; status:0; dependent,scrollbars");
	}
	//set lastfocus back
	oLastFocusID=oTempLastFocusID;
}

//a mimessage has had its status changed in another window of the app
function fnUpdateStatusChange(sSite,sStudy,sSubject,sPageTaskID,sFieldID,nRepeat,nType,sStatus,bRefreshMIMessage)
{
	//get this eforms details
	var sSi=fnGetFormProperty("sSite");
	var sSt=fnGetFormProperty("sStudyId");
	var sSj=fnGetFormProperty("sSubject");
	var sVTID=fnGetFormProperty("sVEformPageTaskID");
	var sTID=fnGetFormProperty("sEformPageTaskID");
	
	//compare this eforms details with the details of the mimessage update
	if((sSi==sSite)&&(sSt==sStudy)&&(sSj==sSubject))
	{
		//same subject, is it the same eform?
		if ((sPageTaskID==sVTID)||(sPageTaskID==sTID))
		{
			//if field is on eform set new discrepancy/sdv status
			if(IsFieldOnForm(sFieldID,nRepeat))
			{
				switch (nType)
				{
				case eDiscrepancyType:
					fnSetDiscrepancyStatus(sFieldID,sStatus,nRepeat);
					break;
				case eSDVType:
					fnSetSDVStatus(sFieldID,sStatus,nRepeat);
					break;
				}
			}
		}
	}
	if(bRefreshMIMessage)
	{
		//if in split mode refresh mimessage window if necessary
		window.parent.window.parent.fnRefreshMIMessageWindow();
	}
}
function fnLoading()
{
	return(window.parent.bLoading);
}
function fnDisableFKeys(b)
{
	window.parent.bLoading=b;
}
function fnHideLoader()
{
	document.all.divMsgBox.style.visibility='hidden';
	if (document.all.wholeDiv!=undefined)
	{
		document.all.wholeDiv.style.visibility='visible';
	}
	fnDisableFKeys(false);
}
function fnShowLoader(sMsg)
{
	fnDisableFKeys(true);
	if (document.all.wholeDiv!=undefined)
	{
		document.all.wholeDiv.style.visibility='hidden';
	}
	document.all["tdMsg"].innerHTML="please wait<br><br><img src='../img/clock.gif'>&nbsp;&nbsp;"+sMsg+"...";
	document.all["divMsgBox"].style.pixelLeft=document.body.scrollLeft+50;
	document.all["divMsgBox"].style.pixelTop=document.body.scrollTop+50;
	document.all["divMsgBox"].style.visibility="visible";
}
function fnSetEformPermissions(sUserName,bChangeData,bAddComment,bViewComment,bViewAudit,bCreateDisc,bCreateSDV,bViewDisc,bViewSDV,bOverruleWarn,bMonitor)
{
	oEP=new Object();
	oEP.sUserName=sUserName;
	oEP.bChangeData=bChangeData;
	oEP.bAddComment=bAddComment;
	oEP.bViewComment=bViewComment;
	oEP.bViewAudit=bViewAudit;
	oEP.bCreateDisc=bCreateDisc;
	oEP.bCreateSDV=bCreateSDV;
	oEP.bViewDisc=bViewDisc;
	oEP.bViewSDV=bViewSDV;
	oEP.bOverruleWarn=bOverruleWarn;
	oEP.bMonitor=bMonitor;
}
function fnOverruleWarnings(sFieldID)
{
	return (oEP.bChangeData&&oEP.bOverruleWarn&&(!fnFieldEformIsReadOnly(sFieldID)));
}
function fnHasResponse(sFieldID,nRepeat)
{
	//return true if field already has response saved (status not requested) or
	//status is requested, but user has changedata permissions
	if (!IsFieldOnForm(sFieldID,nRepeat)) return false;
	return ((getFieldStatus(sFieldID,nRepeat)!=-10)||((getFieldStatus(sFieldID,nRepeat)==-10)&&oEP.bChangeData))
}
function fnHasSDV(sFieldID,nRepeat)
{
	//return true if field already has SDV
	if (!IsFieldOnForm(sFieldID,nRepeat)) return false;
	return (fnGetSDVStatus(sFieldID,nRepeat)!=0)
}
function fnPopMenu(sFieldID,nRepeat,oEvent,nX,nY)
{
	var bEnterable=fnEnterable(sFieldID,nRepeat);
	var bHasResponse=fnHasResponse(sFieldID,nRepeat);
	var bLockedOrFrozen=fnLockedOrFrozen(sFieldID,nRepeat);
	var bEformReadOnly=fnFieldEformIsReadOnly(sFieldID);
	var bHasNote=fnGetNoteStatus(sFieldID,nRepeat);
	var bHasDiscrepancy=fnGetDiscrepancyStatus(sFieldID,nRepeat)!=0;
	var bHasSDV=fnHasSDV(sFieldID,nRepeat);
	
	var sHtml="";
	sHtml+="<table width='180'>";
	// View question details...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(true)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",6);'>View question details...</a>":"View question details...";
	sHtml+="</td></tr>";
	// View question definition...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+="<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",8);'>View question definition...</a>";
	sHtml+="</td></tr>";
	// View audit trail...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bViewAudit&&bHasResponse)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",7);'>View audit trail...</a>":"View audit trail...";
	sHtml+="</td></tr>";
	// Separator
	sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	// View warning...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	// OKWarning = 25 Warning = 30
	sHtml+=(bHasResponse&&((fnGetResponseStatus(sFieldID,nRepeat)==25)||(fnGetResponseStatus(sFieldID,nRepeat)==30)))?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",10);'>View warning...</a>":"View warning...";
	sHtml+="</td></tr>";
	// View inform message... - only show to monitors
	if(oEP.bMonitor)
	{
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		// Inform = 20 
		sHtml+=(bHasResponse&&(fnGetResponseStatus(sFieldID,nRepeat)==20))?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",11);'>View inform message...</a>":"View inform message...";
		sHtml+="</td></tr>";
	}
	// Separator
	sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	// New comment...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bAddComment&&bHasResponse&&(!bLockedOrFrozen)&&(!bEformReadOnly)&&oEP.bChangeData)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",1);'>New comment...</a>":"New comment...";
	sHtml+="</td></tr>";
	// Remove comments...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bAddComment&&bHasResponse&&(!bLockedOrFrozen)&&(!bEformReadOnly)&&oEP.bChangeData)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",2);'>Remove comments</a>":"Remove comments...";
	sHtml+="</td></tr>";
	// Separator
	sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	// New note...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(bHasResponse&&(!bLockedOrFrozen)&&(!bEformReadOnly))?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",3);'>New note...</a>":"New note...";
	sHtml+="</td></tr>";
	// View notes...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(bHasResponse&&bHasNote)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",12);'>View notes...</a>":"View notes...";
	sHtml+="</td></tr>";
	// Separator
	sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	// New discrepancy...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bCreateDisc&&bHasResponse&&(!bLockedOrFrozen)&&(!bEformReadOnly))?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",4);'>New discrepancy...</a>":"New discrepancy...";
	sHtml+="</td></tr>";
	// View discrepancies...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bViewDisc&&bHasResponse&&bHasDiscrepancy)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",13);'>View discrepancies...</a>":"View discrepancies...";
	sHtml+="</td></tr>";
	// Separator
	sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	// New SDV mark...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bCreateSDV&&bHasResponse&&(!bLockedOrFrozen)&&(!bHasSDV)&&(!bEformReadOnly))?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",5);'>New SDV mark...</a>":"New SDV mark...";
	sHtml+="</td></tr>";
	// View SDV mark...
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bViewSDV&&bHasResponse&&bHasSDV)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",14);'>View SDV mark...</a>":"View SDV mark...";
	sHtml+="</td></tr>";
	// Separator
	sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	// Change status to unobtainable	
	if (((getFieldStatus(sFieldID,nRepeat)==10)||(getFieldStatus(sFieldID,nRepeat)==-10))&&bHasResponse&&bEnterable)
	{
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		sHtml+="<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",9);'>Change status to unobtainable</a>";
		sHtml+="</td></tr>";
		// Separator
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	}
	else if ((getFieldStatus(sFieldID,nRepeat)==-5)&&bHasResponse&&bEnterable)
	{
		sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
		sHtml+="<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",15);'>Change status to missing</a>";
		sHtml+="</td></tr>";
		// Separator
		sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
	}
	// Clear
	sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
	sHtml+=(oEP.bChangeData&&bEnterable)?"<a href='javascript:fnPopMenuClick(\""+sFieldID+"\","+nRepeat+",0);'>Clear</a>":"Clear";	
	sHtml+="</td></tr>";
	sHtml+="</table>";

	document.all["divPopMenu"].innerHTML=sHtml;
	fnPopMenuShow(document.all['divPopMenu']);
//	document.all["divPopMenu"].style.visibility='visible';
	if(nX!=undefined)
	{
//		document.all["divPopMenu"].style.pixelLeft=nX; 
//		document.all["divPopMenu"].style.pixelTop=nY;
		fnHideSelects(1,sFieldID,nRepeat,true);
	}
	else
	{
//		document.all["divPopMenu"].style.pixelLeft=document.body.scrollLeft+oEvent.clientX; 
//		document.all["divPopMenu"].style.pixelTop=document.body.scrollTop+oEvent.clientY;
		fnHideSelects(1,sFieldID,nRepeat);
	}
}
function fnPopMenuClick(sId,nRepeat,nItem)
{
	var bOK=true;
	if((nItem!=0)&&((oCurrentFieldID!=null)&&(oCurrentFieldID!=undefined)))
	{
		if (oCurrentFieldID.name==sId)
		{
			bOK=fnCheckCurrentField();
		}
	}
	if (bOK)
	{
		switch (nItem)
		{
			case 0: fnClearField(sId,nRepeat);break;
			case 1: var bRtn=fnReplaceAIBlock(sId,"c","","","",nRepeat);break;
			case 2: fnDeleteCommentBlock(sId,nRepeat);break;
			case 3: fnAddAIBlock(sId,"n",nRepeat);break;
			case 4: fnAddAIBlock(sId,"d",nRepeat);break;
			case 5: fnAddAIBlock(sId,"s",nRepeat);break;
			case 6: fnQuestion(sId,"v",nRepeat);break;
			case 7: fnQuestion(sId,"a",nRepeat);break;
			case 8: fnQuestion(sId,"d",nRepeat);break;
			case 9: setFieldNewStatus(sId,-5,nRepeat);break;
			case 10: // view warning
					fnShowValidationDialog(sId,nRepeat);
					break;
			case 11: // view inform
					fnShowValidationDialog(sId,nRepeat);
					break;
			case 12: // view notes
					fnMIMessageWindow(sId,nRepeat,"2");
					break;
			case 13: // view discrepancies
					fnMIMessageWindow(sId,nRepeat,"0");
					break;
			case 14: // view SDV mark
					fnMIMessageWindow(sId,nRepeat,"3");
					break;
			case 15: setFieldNewStatus(sId,10,nRepeat);break;
		}
	}
}
function fnQuestion(sFieldID,sType,nRepeat)
{
	var sDb=fnGetFormProperty("sDatabase");
	var sSi=fnGetFormProperty("sSite");
	var sSt=fnGetFormProperty("sStudyId");
	var sSj=fnGetFormProperty("sSubject");
	if (fnIsOnUserEform(sFieldID))
	{
		var sEf=fnGetFormProperty("sEformID");
		var sId=fnGetFormProperty("sEformPageTaskID");
	}
	else
	{
		var sEf=fnGetFormProperty("sVEformID");
		var sId=fnGetFormProperty("sVEformPageTaskID");
	}
	var sFd=fnGetQuestionID(sFieldID);
	var sCaption = fnGetFieldProperty(sFieldID,'sCaptionText');
	switch (sType)
	{
		case "v":
			var olArg=new Array();
			olArg[0]=fnGetQuestion(sFieldID)
			olArg[1]=nRepeat;
			var sRtn=window.showModalDialog('QuestionValue.htm',olArg,'dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
			break;
		case "d":
			var sRtn=window.showModalDialog('../sites/'+sDb+'/'+sSt+'/qd'+sFd+'.html','','dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
			break;
		case "a":
			var sRtn=window.showModalDialog('QuestionAudit.asp?fltSi='+sSi+'&fltSt='+sSt+'&fltSj='+sSj+'&fltId='+sId+'&fltFd='+sFd+'&caption='+sCaption+'&fltRp='+nRepeat,'','dialogHeight:300px; dialogWidth:500px; center:yes; resizable:1; status:0; dependent,scrollbars');
			break;
	}
}

//ic 07/12/2001
//function adds a mimessage block to an additional info field
function fnAddAIBlock(sFieldID,sType,nRepeat)
{
	var undefined;
	var sURL;
	
	if ((sFieldID=="")||(nRepeat==undefined))
	{
		alert("No field selected");
		return;
	}
			
	var oAIField=fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	var sCaption=fnGetFieldProperty(sFieldID,"sCaptionText");
	var sValue=fnGetFieldText(sFieldID,nRepeat);
	var sIn=oAIField.value;
	var aArgs=new Array();
	//question caption,question value
	aArgs[0]=sCaption;aArgs[1]=sValue;
	switch (sType)
	{
		case "d":sURL="DialogNewDiscrepancy.htm";break;
		case "s":sURL="DialogNewSDV.htm";break;
		default:sURL="DialogNewNote.htm";break;
	}
	sRtn=window.showModalDialog(sURL,aArgs,'dialogHeight:250px;dialogWidth:500px;center:yes;status:0;dependent,scrollbars');		
	
	if ((sRtn!=undefined)&&(sRtn!="")) 
	{
		oAIField.value=sIn+sDel1+sRtn;
		
		// store AIfield value for RQG redraws
		fnStoreAIInfo(sFieldID,nRepeat,oAIField.value);

		switch (sType)
		{
		case "d":
			fnSetDiscrepancyStatus(sFieldID,eDiscStatus.Raised,nRepeat);
			break;
		case "s":
			var nStatus;			
			switch (sRtn.split(sDel2)[2])
			{
				case "1": nStatus=eSDVStatus.Planned;break;
				case "2": nStatus=eSDVStatus.Queried;break;
				default: nStatus=eSDVStatus.Complete;break;
			}
			fnSetSDVStatus(sFieldID,nStatus,nRepeat);
			break;
		default:
			fnSetNoteStatus(sFieldID,"n",true,nRepeat);
		}
		fnSeteFormToChanged();
		//setFieldFocus(sFieldID,nRepeat);
	}
}
		
//ic 10/12/2001
//function replaces a type block in a question ai field
function fnReplaceAIBlock(sFieldID,sType,sPassword,sOverruleWarning,bUnobtainable,nRepeat,sRFC)
{
	var undefined;
	var sOut="";
	var sTxt="";
	var sBlock="";
	var aBlock;
	var aItm;
	var bOK=true;
	var sDb=fnGetFormProperty("sDatabase");
	var sSt=fnGetFormProperty("sStudyId");	
	var aPwd;
	var sUser;
				
	if (sFieldID=="")
	{
		alert("No field selected");
		return bOK;
	}
		
	var oAIField = fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	var sCaption = fnGetFieldProperty(sFieldID,"sCaptionText");
				
	//get the info field value from the form, split on major delimiter
	var aAdd=oAIField.value.split(sDel1);
				
	//extract the existing type block (if any) from the info field
	for (var n=0;n<aAdd.length;n++)
	{
		//split block on minor delimeter, check first item denoting 'type'
		//[c]comment, [d]discrepancy, [n]note, [s]sdv, [p]password, [r]reason for change
		aItm = aAdd[n].split(sDel2);
		if ((aItm[0]!=sType)&&(aItm[0]!=""))
		{
			//build sOut by concatenating all the blocks not of search type
			sOut+=aAdd[n]+sDel1;
		}
		else
		{
			//keep extracted search block (only required for comments as
			//passwords and rfc are discarded and replaced, whereas comments
			//are added to)
			sBlock=aAdd[n];
		}
	}
				
	//if the last char is a delimeter, remove it
	if (sOut.charAt(sOut.length-1)==sDel1) sOut=sOut.substring(0,sOut.length-1)
			
						
	switch (sType)
	{
	case "c":
		//comment blocks have 3 code blocks preceding the comment text
		//[0] type (c)
		//[1] delete all flag (1 or 0)
		//[2] total length of comments including those already saved
		aBlock=sBlock.split(sDel2);
				
		//get delete all flag, comment length
		var bDeleteComments=aBlock[1];
		var nLength=aBlock[2]*1;
			
		//get comment text by omitting first 3 items
		for (n=3;n<aBlock.length;n++)
		{
			sTxt+=aBlock[n]+sDel2;
		}
						
		//remove final delimeter
		sTxt=sTxt.substring(0,sTxt.length-1);
		
		var aArgs=new Array();
		//action,question caption,mimessage text
		aArgs[0]=sCaption;aArgs[1]=nLength;aArgs[2]=oEP.sUserName;	
					
		//get the new comment text to be appended
		var sRtn=window.showModalDialog('DialogNewComment.htm',aArgs,'dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
		
		if ((sRtn!=undefined)&&(sRtn!=""))
		{
			//add a major delimiter
			sOut+=(sOut!="")? sDel1:"";
					
			//add new comment block
			var sNewComment="c"+sDel2+bDeleteComments+sDel2+(nLength+sRtn.length+3)+sDel2+sRtn+sTxt;			
			sOut+=sNewComment;

			//update additional info field with new value
			oAIField.value=sOut;
			fnSetNoteStatus(sFieldID,"c",true,nRepeat);
//			var sStr=fnRemoveDelimiters(sRtn)+fnRemoveDelimiters(sTxt);
//			fnSetComments(sFieldID,nRepeat,sStr);
			fnSetComments(sFieldID,nRepeat,sNewComment);
		}
					
		break;
					
	case "p":	
		//add password to the sOut blocks
		sOut+=(sOut!="")? sDel1:"";
		sOut+="p"+sDel2+sPassword;
		oAIField.value=sOut;
		//get username
		aPwd=sPassword.split(sDel2);
		sUser=aPwd[0];
		fnSetFullName(sFieldID,nRepeat,sUser);
				
		break;
					
	case "r":
		//add rfc to the sOut blocks
		sOut+=(sOut!="")? sDel1:"";
		sOut+="r"+sDel2+sRFC;
		oAIField.value=sOut;

		break;
				
	case "o":
		//store overrule warning
		sOut+=(sOut!="")? sDel1:"";
		sOut+="o"+sDel2+sOverruleWarning;
		oAIField.value=sOut;
		
		break;
					
	case "u":
		//set to unobtainable/missing
		sOut+=(sOut!="")? sDel1:"";
		sOut+="u"+sDel2+bUnobtainable;
		oAIField.value=sOut;
					
		break;
		
	case "t":
		// rs 02/10/2002: Add current timestamp
		sOut+=(sOut!="")? sDel1:"";
		var dt = new Date();
		var sTimestamp = dt.getDate() + '|' + (dt.getMonth()+1) + '|' + dt.getFullYear() + '|' + dt.getHours() + '|' + dt.getMinutes() + '|' + dt.getSeconds() + '|' + dt.getTimezoneOffset();
		sOut+="t"+sDel2+sTimestamp;
		oAIField.value=sOut;
		
		break;				
	default:
	}

	// store AIfield value for RQG redraws
	fnStoreAIInfo(sFieldID,nRepeat,oAIField.value);
						
	return bOK;	
}

//function fnRemoveDelimiters(sStr)
//{
//	return sStr.replace(/[|]/g,"<br>")
//}
		
//ic 10/12/2001
//function deletes a type block from a question ai field
function fnDeleteCommentBlock(sFieldID,nRepeat)
{
	if ((sFieldID=="")||(nRepeat==undefined)) return;
	
	var sOut="";
	var aItm;
	var oAIField = fnGetFieldProperty(sFieldID,"oAIHandle",nRepeat);
	var sCaption = fnGetFieldProperty(sFieldID,"sCaptionText");
	var sIn=oAIField.value;
	var aAdd=sIn.split(sDel1);
	
	if(confirm('Are you sure you wish to delete all comments for the question "'+sCaption+'"?'))
	{
		for (var n=0;n<aAdd.length;n++)
		{
			aItm = aAdd[n].split(sDel2);
			if (aItm[0]!="c")
			{
				sOut+=aAdd[n]+sDel1;
			}
		}
		var sDeletedComment="c"+sDel2+"1"+sDel2+"0"
		oAIField.value=sOut+sDeletedComment;
		fnSetNoteStatus(sFieldID,"c",false,nRepeat);
		fnSetComments(sFieldID,nRepeat,sDeletedComment);
	}
}

// function to open discrepancy/sdv/note window
function fnMIMessageWindow(sFieldID,nRepeat,sType)
{
	//null lastfocus variable as otherwise a gotfocus() can fire during or shortly 
	//after the following validation process resulting in multiple rfc dialogs. set 
	//it back when we exit
	var oTempLastFocusID=oLastFocusID;
	oLastFocusID=null

	//validate current field, if any
	if (fnCheckCurrentField())
	{
		//check for eform changes. these must be saved before viewing mimessages
		if (fnUserHasChangedEForm()||fnDerivationHasChangedEForm())
		{
			if(confirm("Unsaved responses on the eForm must be saved before viewing messages."))
			{
				fnSave("m0");
			}
			//set lastfocus back
			oLastFocusID=oTempLastFocusID;
			//user doesnt want to save changes
			return;
		}

		//get required info
		var sSi=fnGetFormProperty("sSite");
		var sSt=fnGetFormProperty("sStudyId");
		var sSj=fnGetFormProperty("sSubject");
		var sFd=fnGetQuestionID(sFieldID);
		var sVi=fnGetFormProperty("sVisitID");
		var sVRpt=fnGetFormProperty("sVisitCycle");
		
		//get eformid/taskid for either user or visit eform
		if (fnIsOnUserEform(sFieldID))
		{
			var sEf=fnGetFormProperty("sEformID");
			var sERpt=fnGetFormProperty("sEformCycle");
			var sId=fnGetFormProperty("sEformPageTaskID");
		}
		else
		{
			var sEf=fnGetFormProperty("sVEformID");
			var sERpt=fnGetFormProperty("sVEformCycle");
			var sId=fnGetFormProperty("sVEformPageTaskID");
		}

		var aArgs=new Array();
		aArgs[0]=window;
		var sMIMessageURL="MIMessage.asp?fltSt="+sSt+"&fltSi="+sSi+"&fltSj="+sSj+"&fltVi="+sVi+"&fltVRpt="+sVRpt+"&fltEf="+sEf+"&fltERpt="+sERpt+"&fltQu="+sFd+"&type="+sType+"&newwindow=1&fltQRpt="+(nRepeat+1);		
		var	sResult=window.showModalDialog(sMIMessageURL,aArgs,"dialogHeight:600px; dialogWidth:750px; center:yes; resizable:1; status:0; dependent,scrollbars");
	}
	//set lastfocus back
	oLastFocusID=oTempLastFocusID;
}

// function to show tooltips on data entry fields
function fnSetTooltip(sFieldID,nRepeat,nType)
{
	var sText="";
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];
	
	if(nType<2)
	{
		sText=fnToolTipText(sFieldID,nRepeat);
		if(nType==1)
		{
			oField.oHandle.title=sText;
		}
		else
		{
			// radio button image
			fnSetRadioTooltip(oField.oRadio,sText);
		}
	}
	else
	{
		// status icon image
		sText=fnStatusToolTipText(sFieldID,nRepeat);
		if(oFieldTemp.olRepeat.length>1)
		{
			// array
			var oImageHandle=oForm.oOtherPage[oField.sImageName][nRepeat];
		}
		else
		{
			// standard
			var oImageHandle=oForm.oOtherPage[oField.sImageName];
		}
		if(oImageHandle!=null)
		{
			oImageHandle.alt=sText;
		}
	}	
}

// Return the text for a question tooltip
function fnToolTipText(sFieldID,nRepeat)
{
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];
	var sTxt="Question: "+fnGetFieldProperty(sFieldID,"sCaptionText");
	//var sTxt="Question: "+oFieldTemp.sCaptionText;
	var sHlpTxt=fnGetFieldProperty(sFieldID,"sHelpText")
	if((oField.sComments==null)||(oField.sComments==undefined))
	{
		var sComments="";
	}
	else
	{
		var aC=oField.sComments.split("|")
		var sComments="";
		for(n=3;n<aC.length-1;n++)sComments+=aC[n]+"|";
		sComments=sComments.replace(/[|]/g," ");
	}
	if((sComments!="")&&(sComments!=null))
	{
		// if comments show them
		sTxt+=" - Comments: "+sComments;
	}
	else
	{
		if((sHlpTxt!="")&&(sHlpTxt!=null))
		{
			// show help text
			sTxt+=" - "+sHlpTxt;
		}
	}
	return sTxt;
}

// Return the text for a status icon tooltip
function fnStatusToolTipText(sFieldID,nRepeat)
{
	var	oFieldTemp=oForm.olQuestion[sFieldID];
	var oField=oFieldTemp.olRepeat[nRepeat];
	//var sTxt="Question: "+oFieldTemp.sCaptionText;
	var sTxt="Question: "+fnGetFieldProperty(sFieldID,"sCaptionText");
	var sHlpTxt=fnGetFieldProperty(sFieldID,"sHelpText")
	var nStatus=getFieldStatus(sFieldID,nRepeat);
	// TODO sValMessage / sStatusTxt
	var sValMessage=fnGetValidationMessage(sFieldID,nRepeat,nStatus);
	
	if(sValMessage!="")
	{
		var sStatusTxt=fnGetStatusString(nStatus);
		sTxt+=" - "+sStatusTxt+": "+sValMessage;
		// if an OKWarning status
		if(nStatus==25)
		{
			var sRFO=fnGetRFO();
			if(sRFO!="")
			{
				sTxt+=". Overrule Reason: "+sRFO+".";
			}
		}
	}
	return sTxt;
}

function fnPopMenuShow(oDiv)
{
	var nPopHeight=oDiv.clientHeight;
	var nPopWidth=oDiv.clientWidth;
	var nPopYPosition=window.event.clientY;
	var nPopXPosition=window.event.clientX;
	var nVisibleScreenHeight=document.body.clientHeight;
	var nVisibleScreenWidth=document.body.clientWidth;
	var nX=0;
	var nY=0;
	
	if((nVisibleScreenHeight-(nPopYPosition+nPopHeight)>0)||((nPopYPosition-nPopHeight)<0))
	{
		//if space below or not space above, display below
		nY=nPopYPosition;
	}
	else
	{
		//display above
		nY=nPopYPosition-nPopHeight;
	}
	if((nVisibleScreenWidth-(nPopXPosition+nPopWidth)>0)||((nPopXPosition-nPopWidth)<0))
	{
		//if space right or not space left, display right
		nX=nPopXPosition;
	}
	else
	{
		//display left
		nX=nPopXPosition-nPopWidth;
	}
	
  oDiv.style.pixelLeft=document.body.scrollLeft+nX;
  oDiv.style.pixelTop=document.body.scrollTop+nY;
  oDiv.style.visibility='visible';
}

function fnExecCompressed(sCompressedJS){
	var Sep='`';
	var sRep;
	var result=new RegExp('^(.*)'+Sep+'([^'+Sep+']*)$').exec(sCompressedJS);
	for(var i=result[2].valueOf();i>0;i--){
		result=new RegExp('^(.*)'+Sep+'([^'+Sep+']*)$').exec(result[1]);
		result[1]=result[1].replace(new RegExp(Sep+i+Sep,'g'),result[2]);
	}
	eval(result[1]);
}

function fnUserHasChangedEForm()
{
	return beFormChanged;
}

function fnSeteFormToChanged()
{
	beFormChanged=true;
}

function fnDerivationHasChangedEForm()
{
	return bDerivChanged;
}

function fnSetDeriveFormToChanged()
{
	bDerivChanged=true;
}

// look at all DIVs on page and set next button position
// set "page color" DIV size at same time
function fnSetBtnNextPos()
{
	var lMaxX=0;
	var lMaxY=0;
	var lPageX=0;
	var lPageY=0;
	var lRight;
	var lBottom;
	var aDivs = document.all.tags("DIV");
	for(var i=0; i<aDivs.length; i++)
	{
		if((aDivs[i].id!="btnNextDIV")&&(aDivs[i].id!="eFormColorDIV")&&(aDivs[i].id!="wholeDiv"))
		{
			// right most position
			lRight=aDivs[i].style.pixelLeft+aDivs[i].clientWidth;
		
			// bottom position
			lBottom=aDivs[i].style.pixelTop+aDivs[i].clientHeight;
		
			// change max values if appropriate
			if(lRight>lMaxX){lMaxX=lRight;}
			if(lBottom>lMaxY){lMaxY=lBottom;}
		}
	}
	// move button next DIV
	// position button 30 pix inside right edge last element 
	// will make page exact size
	lPageX=document.all("eFormHeaderTable").clientWidth;
	if(lPageX<lMaxX)
	{
		lPageX=lMaxX;
		// resize header - knock on effect to display formatting
		//document.all("eFormHeaderTable").width=lMaxX;
	}
	lPageY=lMaxY+45;
	lMaxX-=32;
	if(lMaxX<0){lMaxX=0;}
	lMaxY+=10;
	
	// set button
	document.all("btnNextDIV").style.pixelLeft=lMaxX;
	document.all("btnNextDIV").style.pixelTop=lMaxY;
	
	// set page color DIV
	document.all("eFormColorDIV").style.pixelWidth=lPageX;
	document.all("eFormColorDIV").style.pixelHeight=lPageY;	
}

function fnOnBtnNext()
{
	bBtnNext=true;
}

// disable eform / visit clickable labels
function fnDisableEformVisit()
{
	// Call disable function in EformTop & eFormLh
	if((window.parent.frames[0].fnDisable!=undefined)&&(window.parent.frames[0].fnDisable!=null))
	{
		window.parent.frames[0].fnDisable();
	}
	if((window.parent.frames[1].fnDisable!=undefined)&&(window.parent.frames[1].fnDisable!=null))
	{
		window.parent.frames[1].fnDisable();
	}
}
function fnError(sMsg,sUrl,sLine)
{
	if(!bErrorReported)
	{
		var olArg=new Array();
		olArg[0]=sMsg;
		olArg[1]=sUrl;
		olArg[2]=sLine;
		olArg[3]=window.document.body.outerHTML;
		olArg[4]=navigator.userAgent;
		var s=window.showModalDialog('DialogEformError.asp',olArg,'dialogHeight:300px; dialogWidth:500px; center:yes; status:0; dependent,scrollbars');
		bErrorReported=true;
	}
	return true;
}
//check a selection has been made
function fnSelection(oInput)
{
	var b=false;

	if (oInput.length==undefined)
	{
		b=oInput.checked;
	}
	else
	{
		for (var n=0;n<oInput.length;n++)
		{
			if (oInput[n].checked) b=true;
		}
	}

	return b;
}
var bVTries=0;
var bETries=0;
function fnLoadVisitBorder(sSubjectLabel,sVisitList,sSelectedVisitTaskId)
{
	//try to call the eform border load function in the lh frame for 10 seconds, then assume it cant load
	if(bVTries < 10)
	{
		if((window.parent.frames[0].fnLoadVisitBorder==undefined)||(window.parent.frames[0].fnLoadVisitBorder==null))
		{
			//the visit frame has not loaded yet
			bVTries++;
			window.setTimeout('fnLoadVisitBorder("'+sSubjectLabel+'","'+sVisitList+'","'+sSelectedVisitTaskId+'");',1000);
		}
		else
		{
			window.parent.frames[0].fnLoadVisitBorder(sSubjectLabel,sVisitList,sSelectedVisitTaskId);
		}
	}
	else
	{
		alert("MACRO was unable to load the visit border frame at the top. This could be due to a\nslow internet connection. Please try re-loading the eForm");
	}
}
function fnLoadEformBorder(sEformList,sSelectedEformTaskId)
{
	//try to call the eform border load function in the lh frame for 10 seconds, then assume it cant load
	if(bETries < 10)
	{
		if((window.parent.frames[1].fnLoadEformBorder==undefined)||(window.parent.frames[1].fnLoadEformBorder==null))
		{
			//the eform frame has not loaded yet
			bETries++;
			window.setTimeout('fnLoadEformBorder("'+sEformList+'","'+sSelectedEformTaskId+'");',1000);
		}
		else
		{
			window.parent.frames[1].fnLoadEformBorder(sEformList,sSelectedEformTaskId);
		}
	}
	else
	{
		alert("MACRO was unable to load the eForm border frame on the left. This could be due to a\nslow internet connection. Please try re-loading the eForm");
	}
}
document.onkeydown=fnKeyDown;
window.onerror=fnError;
