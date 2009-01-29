///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	functions for handling main window
//
// 
//
// RS 26/05/2005:  Changed fnEformURL to be able to open eForm from a MACRO Home Page Report
///////////////////////////////////////////////////////////////////////////////////////////
var undefined;
var sDel1="`";
var sDel2="|";
var sDel3="~";

function fnSetAppState()
{
	var sAppState=top.sAppState;
	top.sAppState="";
	fnSetWinState(sAppState);
}
function fnSetWinState(sState)
{		
	var bIsSplit=fnIsSplit();
	var aArg;
	
	if (sState!="")
	{
		var aWin=sState.split(sDel1);			
		//loop through window states
		for (n=0;n<aWin.length;n++)
		{	
			aArg=aWin[n].split(sDel2);
			//aArg[0] is frame number
			switch (aArg[0])
			{
			case "1":
				//code to handle set state for menu frame
				break;
			case "5":
			case "6":			
				if ((bIsSplit)||((!bIsSplit)&&(aArg[0]=="5")))
				{
					//aArg[0] is page loader function number
					switch (aArg[1])
					{
					case "0":
						fnNewSubjectUrl();
						break;
					case "1":
						fnSubjectListUrl(aArg[2],aArg[3],aArg[4],aArg[5],aArg[6]);
						break;
					case "3":
						fnMIMessageUrl(aArg[2],aArg[3],aArg[4],aArg[5],aArg[6],aArg[7],aArg[8],aArg[9],aArg[10],aArg[11],aArg[12],aArg[13])
						break;
					case "4":
						fnScheduleUrl(aArg[2],aArg[3],aArg[4])
						break;
					case "5":
						fnEformUrl(aArg[2],aArg[3],aArg[4],aArg[5],aArg[6])
						break;
					case "6":
						fnBrowserUrl(aArg[2],aArg[3],aArg[4],aArg[5],aArg[6],aArg[7],aArg[8],aArg[9],aArg[10],aArg[11],aArg[12],aArg[13],aArg[14],aArg[15],aArg[16],aArg[17])
						break;
					case "7":
						fnBufferUrl(aArg[2],aArg[3],aArg[4])
						break;
					case "2":
					case "undefined":
					default:
	//					fnHomeUrl();
					}
				}
			default:
			}
		}
	}
	else
	{
		//fnHomeUrl();
	}
}
var bDLATries=0;
function fnEnableDLA(b)
{
	//try to call the enable dla load function in the lh frame for 10 seconds, then assume it cant load
	if(bDLATries < 10)
	{
		if((window.parent.frames[2].fnEnableTaskListItem==undefined)||(window.parent.frames[2].fnEnableTaskListItem==null))
		{
			//the menu frame has not loaded yet
			bDLATries++;
			window.setTimeout('fnEnableDLA('+b+');',1000);
		}
		else
		{
			window.parent.frames[2].fnEnableTaskListItem("DLA",b);
		}
	}
	else
	{
		alert("MACRO was unable to enable the database lock administration option in the menu panel. This could be due to a\nslow internet connection. Please try re-loading the application");
	}
}
var bRegTries=0;
function fnEnableRegister(b)
{
	//try to call the enable dla load function in the lh frame for 10 seconds, then assume it cant load
	if(bRegTries < 10)
	{
		if((window.parent.frames[2].fnEnableTaskListItem==undefined)||(window.parent.frames[2].fnEnableTaskListItem==null))
		{
			//the menu frame has not loaded yet
			bRegTries++;
			window.setTimeout('fnEnableRegister('+b+');',1000);
		}
		else
		{
			window.parent.frames[2].fnEnableTaskListItem("RS",b);
		}
	}
	else
	{
		alert("MACRO was unable to enable the registration option in the menu panel. This could be due to a\nslow internet connection. Please try re-loading the application");
	}
}
function fnGetAppState(nWin)
{
	return window.frames[nWin].sWinState;
}
function fnSaveDataFirst(sFn)
{
	if (fnEformIsLoaded())
	{
		window.frames[0].window.frames[2].fnConfirmGoto(sFn);
	}
	else
	{
		eval(sFn);
	}
}
//called from mimessage window when not in new window
//only has any effect when in split mode with an eform loaded in top pane
function fnUpdateStatusOnEform(sSite,sStudy,sSubject,sPageTaskID,sFieldID,nRepeat,nType,sStatus)
{
	if (fnEformIsLoaded())
	{
		window.frames[0].window.frames[2].fnUpdateStatusChange(sSite,sStudy,sSubject,sPageTaskID,sFieldID,nRepeat,nType,sStatus,0)
	}
}
function fnRefreshMIMessageWindow()
{
	//if in split mode, and mimessages are loaded, refresh them
	if (fnIsSplit())
	{
		if(window.frames[1].sWinState!=undefined)
		{
			if(window.frames[1].sWinState.substring(0,1)=="3")
			{
				window.frames[1].location.reload(true);
			}
		}
	}
}
//set task list counter
function fnSTLC(sName,sNum,bZ)
{
	window.parent.frames[2].fnSetTaskListItemCounter(sName,sNum,bZ);
}
function fnEformIsLoaded()
{
	if (window.frames[0].sWinState==undefined) return false;
	if (window.frames[0].sWinState.substring(0,1)!="5") return false;
	if (window.frames[0].window.frames[2].bEformLoadCheck==undefined) return false;
	return window.frames[0].window.frames[2].bEformLoadCheck==true;
}
function fnIsSplit()
{
	return bSplit;
}
function fnAddToQuickList(sItem)
{
	window.parent.frames[2].fnAddToQuickList(sItem);
}
function fnRefreshMenu()
{
	//window.parent.frames[2].location.reload(true);
	window.parent.frames[2].navigate("AppMenuLh.asp");
}
//#0
function fnNewSubjectUrl(oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)? window.frames[0]:oTargetWindow;
	oTargetWindow.navigate("NewSubject.asp");
	oTargetWindow.parent.fnSetTitle("New Subject")
}
//#1
function fnSubjectListUrl(sSt,sSi,sSj,nOrderBy,nBookmark,oTargetWindow)
{
	sSt=(sSt==undefined)?"":sSt;
	sSi=(sSi==undefined)?"":sSi;
	sSj=(sSj==undefined)?"":sSj;
	nBookmark=(nBookmark==undefined)?0:nBookmark;
	nOrderBy=(nOrderBy==undefined)?-1:nOrderBy;
	oTargetWindow=(oTargetWindow==undefined)?window.frames[0]:oTargetWindow;
	oTargetWindow.navigate("ListSubject.asp?fltSt="+sSt+"&fltSi="+sSi+"&fltLb="+sSj+"&orderby="+nOrderBy+"&bookmark="+nBookmark);
	oTargetWindow.parent.fnSetTitle("Subject List")
}
//#2
function fnHomeUrl(oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)?window.frames[0]:oTargetWindow;
	oTargetWindow.navigate("Home.asp");
	oTargetWindow.parent.fnSetTitle("Home")
}

//#3
function fnMIMessageUrl(sType,sSt,sSi,sVi,sEf,sQu,sSj,sUs,sB4,sTm,sSs,sSc,nBookMark,oTargetWindow,sVRpt,sERpt,sQRpt)
{
	var nWinNdx=(fnIsSplit())? 1:0;
	nBookMark=(nBookMark==undefined)?"0":nBookMark;
	var sTitle;
	oTargetWindow=(oTargetWindow==undefined)?window.frames[nWinNdx]:oTargetWindow;
	oTargetWindow.navigate("MIMessage.asp?type="+sType+"&fltSt="+sSt+"&fltSi="+sSi+"&fltVi="+sVi+"&fltEf="+sEf+"&fltQu="+sQu+"&fltSjLb="+sSj+"&fltUs="+sUs+"&fltSs="+sSs+"&fltB4="+sB4+"&fltTm="+sTm+"&fltObj="+sSc+"&bookmark="+nBookMark+"&fltVRpt="+sVRpt+"&fltERpt="+sERpt+"&fltQRpt="+sQRpt);
	switch (sType)
	{
		case "0": sTitle="Discrepancies";break;
		case "1": sTitle="SDVs";break;
		default: sTitle="Notes";break;
	}
	oTargetWindow.parent.fnSetTitle(sTitle,nWinNdx)
}
function fnLogOutUrl(oTargetWindow)
{
	if (oTargetWindow==undefined) oTargetWindow=top;
	if(confirm('Are you sure you wish to logout?'))
	{
		oTargetWindow.navigate("Logout.asp");
	}
}
function fnHoldUrl(sAppState,oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)?top:oTargetWindow;
	if (confirm("This will put MACRO into standby mode."))
	{
		oTargetWindow.navigate("login.asp?"+sAppState);
	}
}
function fnSwitchUrl(sAppState,oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)?top:oTargetWindow;
	if (confirm("Log out and log in as a different user?"))
	{
		oTargetWindow.navigate("login.asp?"+sAppState);
	}
}
//#4
function fnScheduleUrl(sSt,sSi,sSj,bSameEform,oTargetWindow,bNew)
{	
	oTargetWindow=(oTargetWindow==undefined)?window.frames[0]:oTargetWindow;
	bNew=(bNew==undefined)?"0":"1";	
	if (bSameEform)
	{
		fnEformUrl(sSt,sSi,sSj,"","")
	}
	else
	{
		oTargetWindow.navigate("Schedule.asp?fltSt="+sSt+"&fltSi="+sSi+"&fltSj="+sSj+"&new="+bNew);
		oTargetWindow.parent.fnSetTitle("Subject")
	}
}
//#5
function fnEformUrl(sSt,sSi,sSj,sId,sXId,oTargetWindow,sCancelState)
{
	oTargetWindow=(oTargetWindow==undefined)?window.frames[0]:oTargetWindow;
	sCancelState=(sCancelState==undefined)? "":sCancelState;
	// RS 26/05/2005: Added ../app/ to relative path to enable use of the function in a MACRO Home Page report	
	oTargetWindow.navigate("../app/eformFrm.asp?fltSt="+sSt+"&fltSi="+sSi+"&FltSj="+sSj+"&fltId="+sId+"&fltXId="+sXId+"&cancelstate="+sCancelState)
	oTargetWindow.parent.fnSetTitle("Subject")
}
//#6
//function fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sGet,nBookMark,oTargetWindow)
//{
//	var nWinNdx=(fnIsSplit())? 1:0;
//	nBookMark=(nBookMark==undefined)?"0":nBookMark;
//	oTargetWindow=(oTargetWindow==undefined)?window.frames[nWinNdx]:oTargetWindow;
//	oTargetWindow.navigate("databrowser.asp?fltSt="+sSt+"&fltSi="+sSi+"&fltVi="+sVi+"&fltEf="+sEf+"&fltQu="+sQu+"&fltSjLb="+sSj+"&fltUs="+sUs+"&fltSs="+sSs+"&fltLk="+sLk+"&fltB4="+sB4+"&fltTm="+sTm+"&fltAd="+sCm+"&get="+sGet+"&bookmark="+nBookMark);
//	oTargetWindow.parent.fnSetTitle("Data Review",nWinNdx)
//}
//#6
// ATN 17/1/2003 -  added  sDi, sSD, sNo
function fnBrowserUrl(sSt,sSi,sVi,sEf,sQu,sSj,sUs,sSs,sLk,sB4,sTm,sCm,sDi,sSd,sNo,sGet,nBookMark,oTargetWindow)
{
	var nWinNdx=(fnIsSplit())? 1:0;
	nBookMark=(nBookMark==undefined)?"0":nBookMark;
	oTargetWindow=(oTargetWindow==undefined)?window.frames[nWinNdx]:oTargetWindow;
// ATN 17/1/2003 -  added  sDi, sSD, sNo
	oTargetWindow.navigate("databrowser.asp?st="+sSt+"&si="+sSi+"&vi="+sVi+"&ef="+sEf+"&qu="+sQu+"&sjlb="+sSj+"&us="+sUs+"&ss="+sSs+"&lk="+sLk+"&b4="+sB4+"&tm="+sTm+"&cm="+sCm+"&di="+sDi+"&sd="+sSd+"&no="+sNo+"&get="+sGet+"&bookmark="+nBookMark);
	oTargetWindow.parent.fnSetTitle("Data Review",nWinNdx)
}

function fnCreateSubjectUrl(sSt,sSi,oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)?window.frames[0]:oTargetWindow;
	oTargetWindow.navigate("CreateSubject.asp?fltSt="+sSt+"&fltSi="+sSi);
}
function fnPrint(oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)?window:oTargetWindow;
	oTargetWindow.focus();
	oTargetWindow.print();
}
// Buffer data browser url
function fnBufferUrl(sSt,sSi,sSj,oTargetWindow)
{
	oTargetWindow=(oTargetWindow==undefined)?window.frames[0]:oTargetWindow;
	oTargetWindow.navigate("BufferSummary.asp?fltSt="+sSt+"&fltSi="+sSi+"&fltSj="+sSj);
	oTargetWindow.parent.fnSetTitle("Buffer Summary")
}
