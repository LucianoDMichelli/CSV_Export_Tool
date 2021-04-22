function SetFile(sFileIn)
{
	document.getElementById('fileIn').value					= sFileIn;
}
function GetFirstDataRow()
{
	return document.getElementById('firstRow').value;
}
function GetChkStatus(sType)
{
	if (document.getElementById(sType).checked == true)
	{
		return 1;
	}
	else
	{
		return 0;
	}
}
function GetFieldVal(sField)
{
	return document.getElementById(sField).value
}
function ExportOptChange()
{
	document.getElementById("utype").checked				= true;
	document.getElementById("unit").checked					= true;
	document.getElementById("person").checked				= true;
	document.getElementById("tenant").checked				= true;
	document.getElementById("leasecharge").checked			= true;
	document.getElementById("demographics").checked			= true;
	document.getElementById("secdeps").checked				= true;
	document.getElementById("roommates").checked			= true;
}
function ResetExportOpts(sSelf)
{
	/*var	OptionExport							= document.getElementById(sSelf).checked;
	
	alert(sSelf + " begins as " + OptionExport)
	
	if (OptionExport == true)
	{
		document.getElementById(sSelf).checked	= false;
		alert (sSelf + " is now FALSE");
	}
	else
	{
		document.getElementById(sSelf).checked	= true;
		alert (sSelf + " is now TRUE");
	}*/
}
function ProgressBarInc(i)
{
	var	pbiName								= "pbar" + i;
	
	document.getElementById(pbiName).style.background		= "#0000FF";
}
function Updater(sUpdate)
{
	document.getElementById("statusupdater").innerHTML		= sUpdate;
}
function ResetAll()
{
	var	i;
	var	pbiName;
	var	sUpdate								= "";
	for(i=1; i<=18; i++)
	{
		pbiName								= "pbar" + i;
		document.getElementById(pbiName).style.background	= "#FFFFFF";
	}
	
	Updater(sUpdate)
}