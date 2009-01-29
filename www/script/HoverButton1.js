function fnInitialiseButton(sName,clsHover,clsUnHover,clsSelected,bInitialSelect)
{
	var oClass = new Object();
	oClass.clsHover=clsHover;
	oClass.clsUnHover=clsUnHover;
	oClass.clsSelected=clsSelected;
	
	var aButtons=document.all[sName];
	for (var n=0;n<aButtons.length;n++)
	{
		aButtons[n].ndx=n;
		aButtons[n].bSelected=false;
		aButtons[n].oClass=oClass;
	}
	if (bInitialSelect)
	{
		for (n=0;n<aButtons.length;n++)
		{
			var sPattern = /clsHoverButtonDisabled/;
			if (sPattern.exec(aButtons[n].className)==null)
			{
				aButtons[n].bSelected=true;
				return;
			}
		}
	}
}
function fnGetSelectedButton(sName)
{
	var oDiv;
	var aButtons=document.all[sName];
	for (var n=0;n<aButtons.length;n++)
	{
		if (aButtons[n].bSelected==true)
		{
			return aButtons[n];
		}
	}
	return oDiv;
}
function fnSetButtonHover(oDiv)
{
	if (!oDiv.bSelected)
	{
		oDiv.className=oDiv.oClass.clsHover;
	}
}
function fnSetButtonUnHover(oDiv)
{
	if (!oDiv.bSelected)
	{
		oDiv.className=oDiv.oClass.clsUnHover;
	}
}
function fnSetButtonSelected(oDiv)
{
	//get currently selected button, deselect it
	var oSelDiv=fnGetSelectedButton(oDiv.id);
	if (oSelDiv!=undefined) fnSetButtonUnSelected(oSelDiv);
	//set passed button selected
	oDiv.bSelected=true
	oDiv.className=oDiv.oClass.clsSelected;
	//call onclick handler
	fnDoButtonClick(oDiv.ndx);
}
function fnSetButtonUnSelected(oDiv)
{
	oDiv.className=oDiv.oClass.clsUnHover;
	oDiv.bSelected=false;
}