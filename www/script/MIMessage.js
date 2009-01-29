///////////////////////////////////////////////////////////////////////////////////////////
//
//	(c) InferMed 2003
//
//	MIMessage event handlers
//	fnLoadAudit(...) loads audit asp into audit frame
//	fnM(...) pops context menu
//	fnMClick(...) handles context menu events
//
///////////////////////////////////////////////////////////////////////////////////////////

var undefined;
var gsItem=""; //global clicked item's delimited list
var oU; //user permissions object

function fnInitUser(bChangeData,bCreateDiscrepancy,bCreateSDV)
{
	oU=new Object();
	oU.bChangeData=bChangeData;
	oU.bCreateDiscrepancy=bCreateDiscrepancy;
	oU.bCreateSDV=bCreateSDV;
}

function fnLoadAudit(sType,sItem)
{
    var sUrl;
    var aItem=sItem.split("`");
    
    sUrl= "MIMessageBot.asp";
    sUrl+="?mimessagetype="+sType;
    sUrl+="&study="+aItem[1];
//    sUrl+="&subject="+aItem[1];
//    sUrl+="&subjectid="+aItem[3];
//    sUrl+="&visit="+aItem[4];
//    sUrl+="&eform="+aItem[5];
//    sUrl+="&question="+aItem[8];
//    sUrl+="&status="+aItem[9];
//    sUrl+="&value=";
    sUrl+="&id="+aItem[10];
    sUrl+="&src="+aItem[11];
    sUrl+="&site="+aItem[2];
//    sUrl+="&priority="+aItem[12];
    
    window.parent.window.frames[1].location.replace(sUrl)
}
function fnM(sType,sItem,btn,bGoto,bCanEdit,bNewWin)
{
	gsItem=sItem;
    var aItem=sItem.split("`");
    
    var sHtml="";
        
    if (btn==1)
    {
		fnLoadAudit (sType,sItem)
    }
    else
	{sHtml+="<tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
		switch (sType)
		{
		case 0:
			sHtml+="<table width='160'><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnMClick(5,0)'><b>View message history</b></a>"
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(bGoto)?"<a href='javascript:fnMClick(4,0)'>Go to Eform</a>":"Go to Eform";
			sHtml+="</td></tr><tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bCreateDiscrepancy)&&(aItem[9]=="Responded"))?"<a href='javascript:fnMClick(1,0)'>Re-raise discrepancy</a>":"Re-raise discrepancy";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bChangeData)&&(aItem[9]=="Raised"))?"<a href='javascript:fnMClick(0,0)'>Respond to discrepancy</a>":"Respond to discrepancy";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(bCanEdit)?"<a href='javascript:fnMClick(3,0)'>Edit discrepancy</a>":"Edit discrepancy";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bCreateDiscrepancy)&&(aItem[9]=="Raised"||aItem[9]=="Responded"))?"<a href='javascript:fnMClick(2,0)'>Close discrepancy</a>":"Close discrepancy";
			sHtml+="</td></tr></table>";
			break;
		case 1:
			sHtml+="<table width='160'><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnMClick(5,1)'><b>View message history</b></a>"
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(bGoto)?"<a href='javascript:fnMClick(4,1)'>Go to Eform</a>":"Go to Eform";
			sHtml+="</td></tr>";
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(!bNewWin)?"<a href='javascript:fnMClick(9,1)'>Go to Schedule</a>":"Go to Schedule";
			sHtml+="</td></tr>";
			sHtml+="</td></tr><tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr>";
			// Planned Queried Cancelled Done Edit
			sHtml+="<tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bCreateSDV)&&((aItem[9]=="Queried")||(aItem[9]=="Cancelled")||(aItem[9]=="Done")))?"<a href='javascript:fnMClick(6,1)'>Set SDV to planned</a>":"Set SDV to planned";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bCreateSDV)&&((aItem[9]=="Planned")||(aItem[9]=="Cancelled")||(aItem[9]=="Done")))?"<a href='javascript:fnMClick(7,1)'>Set SDV to queried</a>":"Set SDV to queried";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bCreateSDV)&&((aItem[9]=="Planned")||(aItem[9]=="Queried")||(aItem[9]=="Done")))?"<a href='javascript:fnMClick(8,1)'>Set SDV to cancelled</a>":"Set SDV to cancelled";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=((oU.bCreateSDV)&&((aItem[9]=="Planned")||(aItem[9]=="Queried")||(aItem[9]=="Cancelled")))?"<a href='javascript:fnMClick(0,1)'>Set SDV to done</a>":"Set SDV to done";
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(bCanEdit)?"<a href='javascript:fnMClick(1,1)'>Edit SDV</a>":"Edit SDV";
			sHtml+="</td></tr></table>";
			break;
		case 2:
			sHtml+="<table width='160'><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+="<a href='javascript:fnMClick(5,2)'><b>View message history</b></a>"
			sHtml+="</td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(bGoto)?"<a href='javascript:fnMClick(4,2)'>Go to Eform</a>":"Go to Eform";
			sHtml+="</td></tr><tr height='1'><td align=left width='98%' bgcolor='#C0C0C0'></td></tr><tr height='15'><td class='clsPopMenuLinkText'>";
			sHtml+=(bCanEdit)?"<a href='javascript:fnMClick(0,2)'>Edit Note</a>":"Edit Note";
			sHtml+="</td></tr></table>";
			break;
		}
			
        document.all["divPopMenu"].innerHTML=sHtml;
        document.all["divPopMenu"].style.pixelLeft=document.body.scrollLeft+window.event.clientX;
        document.all["divPopMenu"].style.pixelTop=document.body.scrollTop+window.event.clientY;
		document.all["divPopMenu"].style.visibility='visible';
	}
}
function fnMClick(sAction,sType)
{
	var aItem=gsItem.split("`");
	var sURL;
	switch (sAction)
	{
	case 0:
	case 1:
	case 2:
	case 3:
	case 6:
	case 7:
	case 8:
	//UPDATE MIMESSAGE
		var aArgs=new Array();
		//action,question caption,mimessage text
		aArgs[0]=sAction;aArgs[1]=aItem[8];aArgs[2]=aItem[15];
		switch (sType)
		{
			case 0:sURL="DialogUpdateDiscrepancy.htm";break;
			case 1:sURL="DialogUpdateSDV.htm";break;
			default:sURL="DialogUpdateNote.htm";break;
		}
		var sRtn=window.showModalDialog(sURL,aArgs,"dialogHeight:300px;dialogWidth:500px;center:yes;status:0;dependent,scrollbars");
        //handle update, if any
        if ((sRtn!=undefined)&&(sRtn!=""))
		{
			document.FormMI.badd.value=sRtn;
		}
		else
		{	
			return;
		}
		document.FormMI.bidentifier.value=gsItem;
		document.FormMI.btype.value=sType;
		document.FormMI.submit();
		break;
	case 4:
	//GOTO EFORM
		window.parent.window.parent.fnEformUrl(aItem[1],aItem[2],aItem[3],aItem[7],"",undefined,window.parent.sWinState)
		break;
	case 5:
	//LOAD MIMESSAGE AUDIT 
		fnLoadAudit(sType,gsItem);
		break;
	case 9:
	//GOTO SCHEDULE
		window.parent.window.parent.fnScheduleUrl(aItem[1],aItem[2],aItem[3],false,undefined,false)
		break;
	}
}
function fnHideLoader()
{
	document.all["divMsgBox"].style.top=document.body.scrollTop+50;document.all["divMsgBox"].style.left=document.body.scrollLeft+50;
    document.all["divMsgBox"].style.visibility='hidden';
}