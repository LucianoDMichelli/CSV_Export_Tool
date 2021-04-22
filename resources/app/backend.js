// PACKAGES
var dialog = require('electron').remote.dialog;
var fs = require('fs');   
var xlsx = require('node-xlsx'); // https://www.npmjs.com/package/node-xlsx

function GetInFile()
{
    // // Cycles through each possible drive letter to find location of New Acquisitions folder
    // var dPath = "C:\\"
    // var pathNames = ["A:\\New Acquisitions", "B:\\New Acquisitions", "C:\\New Acquisitions", "D:\\New Acquisitions", "E:\\New Acquisitions", "F:\\New Acquisitions", "G:\\New Acquisitions", "H:\\New Acquisitions", "I:\\New Acquisitions", "J:\\New Acquisitions", "K:\\New Acquisitions", "L:\\New Acquisitions", "M:\\New Acquisitions", "N:\\New Acquisitions", "O:\\New Acquisitions", "P:\\New Acquisitions", "Q:\\New Acquisitions", "R:\\New Acquisitions", "S:\\New Acquisitions", "T:\\New Acquisitions", "U:\\New Acquisitions", "V:\\New Acquisitions", "W:\\New Acquisitions", "X:\\New Acquisitions", "Y:\\New Acquisitions", "Z:\\New Acquisitions"]
    // for (let path of pathNames) 
    // {
    //     if (fs.existsSync(path))
    //     {
    //         dPath = path;
    //         break;
    //     }
    // }
    dialog.showOpenDialog( 
        {
            filters: 
                    [  
                        {name: "Excel Files",   extensions: ['xls','xlsx']}, 
                        {name: "All Files",     extensions: ['*']}
                    ],
            properties: ['openFile'],
            defaultPath: __dirname.substring(0,3) + "New Acquisitions", //dPath,
            
        }).then(result => {
            const filePath = result.filePaths.pop().toString();
            const fileDirectory = filePath.substring(0, filePath.lastIndexOf("\\"));
            const fileExtension = filePath.substring(filePath.lastIndexOf(".")+1);
            
            document.getElementById('fileIn').value = filePath;
            document.getElementById('fileExtension').value = fileExtension;
            document.getElementById('fileDirectory').value = fileDirectory;
            
        }).catch(err => {
            console.log(err)
        });
}

function GetPropNum(sPropCode)
{
    return sPropCode.substring(sPropCode.indexOf('-') + 1);
}

function FullScript()
{
    //VARIABLE DECLARATIONS

	//CONSTANTS
    //const textCompare = 1; // Why was this here?
    
    //OBJECTS
    var	oXL;
    var	oWkbk;
    var	oDict;
    var	oResCodes;
    var	oRoomies;
    var	oResRows;
    var	oUnits;
    var oChargeFields;
    
    //STRINGS
    var	sUpdater;
    //var	sType;
    //var	sField;
    var	sLine;
    var	sFileIn;
    var	sFilePath;
    var	sFileOut;
    var	sResCode;
    var	sResZero;
    var	sName;
    var	sFName;
    var	sLName;
    var	sStatus;
    var	sPersonType;
    var	sFedID;
    var	sPropCode;
    var	sPropNum;
    //var	sResStatus;
    var	sLessee;
    //var	sCell;
    //var	sKey;
    var	sRentalType;
    var	sUnitCode;
    var	sUtilCode;
    // var	sRRent;
    // var	sRPark;
    // var	sRMTM;
    // var	sRPetFee;
    // var	sRConces;
    // var	sRRubs;
    var	sRoomCode;
    var	sSex;
    var	sMarital;
    var	sEmployer;
    var sUnitType;
    //var sOccupation;
    
    //INTEGERS
    var	iProgress;
    //var	iResult;
    var	iOverWrite;
    var	iRow;
    var	iFirstRow;
    var	iResCode;
    //var	iRoomCode;
    var	iResCodeFirst;
    //var	iRoomCodeFirst;
    var	iResZero;
    var	iPersonType;
    var iOccupantType;
    //var	iCurrRow;
    // var	iComma;
    var overwriteFlag;
    
    //DECIMALS
    var	dRConces;
    var	dAnnual;
    //var	dMonthly; 
    
    //DATES
    var	dtMoveIn;
    var	dtMoveOut;
    var	dtLeaseFrom;
    var	dtLeaseTo;
    var	dtDOB;

    //Unknown
    //var unitSqft;
    var secDeposit;
    var petDeposit;

    
    //CLEAN THE SCREEN IF NECESSARY
    ResetAll();

    // Dictionary containing all codes as keys for their corresponding field names (ex. {"sStatus": "Status"})
    var fileExtension = GetFieldVal('fileExtension');
    var fieldNames = new Map();

    if (fileExtension === "xls") // xls uses /n for line breaks, xlsx uses /r/n
    {
        fieldNames
                .set("sUnitCode", "Unit")
                .set("sUnitType", "Unit\nType")
                .set("unitSqft", "Unit\nSQFT")
                .set("sFName", "First Name")
                .set("sLName", "Last Name")
                .set("sStatus", "Status")
                .set("marketRent", "Market\nRent")
                // .set("sRRent", "Lease\nRent")
                .set("dtMoveIn", "Move\nin")
                .set("dtLeaseFrom", "Lease\nfrom")
                .set("dtLeaseTo", "Lease\nExp\nDate")
                .set("dtMoveOut", "Vacate Date")
                .set("secDeposit", "Security\nDeposit")
                .set("petDeposit", "Pet\nDeposit")
                // .set("sRPark", "Park")
                // .set("sRMTM", "MTM")
                // .set("sRPetFee", "Pet\nFee")
                // .set("sRConces", "Conces")
                // .set("sRRubs", "rubs")
                .set("sLessee", "Leaseholder Y/N")
                .set("sSex", "Sex\n(M/F)")
                .set("sMarital", "Marital\nStatus\n(S/M)")
                .set("dtDOB", "Date\nof\nBirth\n(MM/DD/YYYY)")
                .set("dAnnual", "Annual Salary")
                .set("sEmployer", "Employer")
                .set("sFedID", "Social\nSecurity\nNumber") 
                .set("sOccupation", "Occupation")
                .set("sEmail", "Email Address") 
                .set("sPhoneNum0", "Telephone Number") 
                .set("sRelationship", "Roommate Type\n(Guarantor/\nOccupant/\nOther/\nRoommate/\nSpouse)")
                .set("iOccupantType", "Adult or Child? (A/C)");
    }
    else
    {
        fieldNames
                .set("sUnitCode", "Unit")
                .set("sUnitType", "Unit\r\nType")
                .set("unitSqft", "Unit\r\nSQFT")
                .set("sFName", "First Name")
                .set("sLName", "Last Name")
                .set("sStatus", "Status")
                .set("marketRent", "Market\r\nRent")
                // .set("sRRent", "Lease\r\nRent")
                .set("dtMoveIn", "Move\r\nin")
                .set("dtLeaseFrom", "Lease\r\nfrom")
                .set("dtLeaseTo", "Lease\r\nExp\r\nDate")
                .set("dtMoveOut", "Vacate Date")
                .set("secDeposit", "Security\r\nDeposit")
                .set("petDeposit", "Pet\r\nDeposit")
                // .set("sRPark", "Park")
                // .set("sRMTM", "MTM")
                // .set("sRPetFee", "Pet\r\nFee")
                // .set("sRConces", "Conces")
                // .set("sRRubs", "rubs")
                .set("sLessee", "Leaseholder Y/N")
                .set("sSex", "Sex\r\n(M/F)")
                .set("sMarital", "Marital\r\nStatus\r\n(S/M)")
                .set("dtDOB", "Date\r\nof\r\nBirth\r\n(MM/DD/YYYY)")
                .set("dAnnual", "Annual Salary")
                .set("sEmployer", "Employer")
                .set("sFedID", "Social\r\nSecurity\r\nNumber") 
                .set("sOccupation", "Occupation")
                .set("sEmail", "Email Address") 
                .set("sPhoneNum0", "Telephone Number") 
                .set("sRelationship", "Roommate Type\r\n(Guarantor/\r\nOccupant/\r\nOther/\r\nRoommate/\r\nSpouse)")
                .set("iOccupantType", "Adult or Child? (A/C)");
    }

    // Looks through header row to find value corresponding to desired field
    function GetVal(valName, iRow, data = oXL, headers = headerLine, dict = fieldNames)
        {              
                if (data[iRow] === undefined)
                {
                    return "";
                }

                var returnVal = data[iRow][headers.indexOf(fieldNames.get(valName))]

                if (returnVal != undefined)
                {
                    if (typeof returnVal === "string") 
                    {
                        return returnVal.trim().replaceAll(",", "");
                    }
                    else
                    {
                        if (valName.substring(0,2) === "dt")
                        {
                            let date = new Date((returnVal - (25567 + 1))*86400*1000); // https://stackoverflow.com/questions/31343129/convert-excel-datevalue-to-javascript-date
                            return date.getMonth()+1 + "/" + date.getDate() + "/" + date.getFullYear()
                        }
                        return returnVal;
                    }
                }
                else
                {
                    console.log("Value for " + valName + " at row" + iRow + " not found")
                    return "";
                }

        }
    
    //INITIALIZATION
    iProgress		= 1
    iResCodeFirst	= 1			
    sPropCode		= GetFieldVal("propCode")
    if (sPropCode === "")
    {
        
        dialog.showErrorBox("","You must enter a valid property code!");
        return;
    }
    
    sPropNum		= GetPropNum(sPropCode)
    
    sFileIn			= GetFieldVal("fileIn")
    if (sFileIn === "")
    {
        dialog.showErrorBox("","You must enter a valid Excel file to import from!")
        return;
    }

    sFilePath = GetFieldVal("fileDirectory")

    oWkbk = xlsx.parse(sFileIn);
    oXL = oWkbk.values().next().value.data;

    // Dictionaries
    oDict = new Map();
    oResCodes = new Map();
    oRoomies = new Map();
    oResRows = new Map();
    oUnits = new Map();
    oChargeFields = new Map();

    iFirstRow = parseInt(GetFirstDataRow())-1; // because arrays start at 0
    const headerLine = oXL[iFirstRow-1];


    if (GetChkStatus("person") === 1 ||
        GetChkStatus("tenant") === 1 ||
		GetChkStatus("leasecharge") === 1 ||
		GetChkStatus("demographics") === 1 ||
		GetChkStatus("roommates") === 1 ||
        GetChkStatus("secdeps") === 1 ||
        GetChkStatus("unit") === 1 ||
        GetChkStatus("utype") === 1)
        {
			iRow			= iFirstRow
			iResCode		= iResCodeFirst
            sStatus         = GetVal('sStatus',iRow).substring(0,1).toLowerCase();

            while (sStatus != "") 
            {
                if (sStatus === "c" ||
                    sStatus === "n" ||
                    sStatus === "f" ||
                    sStatus === "r")
                    {
                        sResZero = "";
                        sResCode = iResCode.toString();
                        iResZero = sResCode.length;
                        while (iResZero <= 3) 
                        {
                            sResZero += "0";
                            iResZero++;
                        }

                        if (sStatus === "r") 
                        {
							sPersonType	= "o";	// '**Was r, but caused repeating sCode issues
                            iPersonType	= 93;
                        }
                        else 
                        {
							sPersonType	= "n";	// '**Was r, but caused repeating sCode issues
                            iPersonType	= 1;
                        }
                        
                        sResCode = sPropNum + sPersonType + sResZero + sResCode;

                        iResCode++;
                    }
                else 
                {
                    sResCode = "VACANT";
                }

                oResCodes.set(iRow, sResCode);

                iRow++;
                sStatus = GetVal('sStatus',iRow).substring(0,1).toLowerCase();
                
            }
        }

        else
        {
            dialog.showErrorBox("","You must select at least one export option!");
            return;
        }

    iRow = iFirstRow;
    sStatus = GetVal('sStatus',iRow).substring(0,1).toLowerCase();

    while (sStatus != "") 
    {
        if (sStatus === "r") 
        {
            oRoomies.set(iRow, iRow);
        }
        else 
        {
            oResRows.set(GetVal('sUnitCode', iRow).toUpperCase(), iRow);
        } 
    
        iRow++;
        sStatus = GetVal('sStatus',iRow).substring(0,1).toLowerCase();
    }

    iRow = iFirstRow;
    sUnitCode = GetVal('sUnitCode', iRow).toUpperCase();

    while (sUnitCode != "")
    {
        if (!oUnits.has(sUnitCode))
        {
            oUnits.set(sUnitCode, iRow);
        }
        
        iRow++;
        sUnitCode = GetVal('sUnitCode', iRow).toUpperCase();
    }

    ProgressBarInc(iProgress);
    iProgress++;
    ProgressBarInc(iProgress);
    iProgress++;
    
    sUpdater = "<table>";

    // CREATE ALL_PEOPLE.CSV

    const allPeople = "\\All_People";
    var multiFlag = 0;
    var allPeopleCounter = 1;
    sFileOut = sFilePath + allPeople + allPeopleCounter + ".csv";

    while (fs.existsSync(sFileOut))
    {
        multiFlag = 1;
        sFileOut = sFilePath + allPeople + allPeopleCounter + ".csv";
        allPeopleCounter++;
    }
    if (multiFlag === 1)
    {
        allPeopleCounter--;
    }
    sUpdater += "<tr><td>Creating All_People" + allPeopleCounter + ".csv...</td>";
    Updater(sUpdater);

    fs.writeFileSync(sFileOut);
    var allPeopleWriter = fs.createWriteStream(sFileOut);

    iRow = iFirstRow;
    sName   = GetVal('sFName', iRow)
            + " "
            + GetVal('sLName', iRow);

    while (sName.trim() != "")
    {

        sResCode        = oResCodes.get(iRow);
        console.log(oResCodes);
        console.log(iRow)
               
        sLine		    = "'" + GetVal('sUnitCode', iRow) + "'" + "," 
                        + "'" + sResCode + "'" + ","
                        + "'" + GetVal('sFName', iRow) + "'" + ","
                        + "'" + GetVal('sLName', iRow) + "'" + ","
        sLine = sLine.replaceAll("''", "")
        allPeopleWriter.write(sLine + "\n");
        iRow++;
        sName   = GetVal('sFName', iRow)
                + " "
                + GetVal('sLName', iRow);
    }
    allPeopleWriter.close();

    sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
    Updater(sUpdater);

    // CREATE PERSON.CSV
    if (GetChkStatus("person") === 1)
    {
        sFileOut = sFilePath + "\\Person_List.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Person_List.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var personWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            iResCode    = iResCodeFirst;
            // if (iResCode.toString() === "")  // iResCodeFirst is set to 1 at the start so this will never happen
            // {
            //     dialog.showMessageBoxSync("You must enter a valid resident code!")
            //     return;
            // }

            while (GetVal('sFName', iRow) + GetVal('sLName', iRow) != "")
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase();
                
                if (!(sStatus === "v"
                    || oResCodes.get(iRow) === "VACANT"))
                {

                    sResCode        = oResCodes.get(iRow);
                    sFName          = GetVal('sFName', iRow);
                    sLName          = GetVal('sLName', iRow);
                    sFedID          = GetVal('sFedID', iRow);

                    if (sFedID.length > 0)
                    {
                        sFedID      = sFedID.replaceAll("-", "");
                    }
                    else
                    {
                        sFedID      = "";
                    }

                    if (sStatus === "r")
                    {
                        iPersonType = 93;
                    }
                    else
                    {
                        iPersonType = 1;
                    }
                    sLine   = "'" + sResCode + "'" + ","
                            + "'" + sLName + "'" + ","
                            + "'" + sFName + "'" + ","
                            + iPersonType + ","
                            + "'" + sLName.toUpperCase() + "'" + ","
                            + "0,"
                            + "'" + sFedID + "'" + ","
                            + "'" + GetVal('sEmail', iRow) + "'" + "," 
                            + "'" + GetVal('sPhoneNum0', iRow) + "'"
                    sLine = sLine.replaceAll("''", "")
                    personWriter.write(sLine + "\n");

                }

                iRow++;
            }
            personWriter.close();
        }

        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }

    }

    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;
    
    // CREATE UNIT_TYPES.CSV
    if (GetChkStatus("utype") === 1)
    {
        sFileOut = sFilePath + "\\Unit_Type_List.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Unit_Type_List.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var unitTypeWriter = fs.createWriteStream(sFileOut);

            iRow = iFirstRow;
            sUnitType = GetVal('sUnitType', iRow);

            while (GetVal('sUnitCode', iRow) != "")
            {
                if (sUnitType != "")
                {
                    if (!oDict.has(sUnitType))
                    {
                        oDict.set(sUnitType, iRow);
                    }
                }

                iRow++;
                sUnitType   = GetVal('sUnitType', iRow);
            }
            
            for (const unitType of oDict) {
                iRow        = unitType[1]

                sLine       = "'" + sPropNum + "-" + unitType[0] + "'" + ","
                            + ","
                            + GetVal('marketRent', iRow) + ","
                            + "0.00,"
                            + GetVal('unitSqft', iRow) + ","
                            + "'" + sPropCode + "'"
                sLine = sLine.replaceAll("''", "")
                unitTypeWriter.write(sLine + "\n");
            }

            unitTypeWriter.close();
        }
        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }
    }
    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;

    // CREATE UNIT.CSV
    if (GetChkStatus("unit") === 1)
    {
        sFileOut = sFilePath + "\\Unit_List.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Unit_List.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var unitWriter = fs.createWriteStream(sFileOut);

            for (const unit of oUnits) {
                iRow        = unit[1];
                sRentalType = "Residential";
                sUnitCode   = GetVal('sUnitCode', iRow);

                sLine       = "'" + sPropCode + "'" + ","
                            + "'" + sPropNum + "-" + GetVal('sUnitType', iRow) + "'" + ","
                            + "'" + sUnitCode + "'" + ","
                            + GetVal('marketRent', iRow) + ","
                            + GetVal('unitSqft', iRow) + ","
                            + "'" + sRentalType + "'" + ","
                            + "0,"
                            + "0,"
                            + "'" + sPropCode + sUnitCode + "'"
                sLine = sLine.replaceAll("''", "")
                unitWriter.write(sLine + "\n");
            }

            unitWriter.close();
        }
        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }
    }
    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;

    // CREATE TENANTS.CSV
    if (GetChkStatus("tenant") === 1)
    {
        sFileOut = sFilePath + "\\Tenant_List.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Tenant_List.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var tenantWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            iResCode    = iResCodeFirst;
            
            while (GetVal('sFName', iRow) + GetVal('sLName', iRow) != "") 
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase();
                

                if (!(sStatus === "r"
                    || sStatus === "v"
                    || oResCodes.get(iRow) === "VACANT"))
                {

                    sResCode        = oResCodes.get(iRow);  
                    sFName          = GetVal('sFName', iRow);
                    sLName          = GetVal('sLName', iRow);
                    
                    dtMoveIn        = GetVal('dtMoveIn', iRow)                      
                    dtMoveOut       = GetVal('dtMoveOut', iRow)                                    
                    dtLeaseFrom     = GetVal('dtLeaseFrom', iRow)                                    
                    dtLeaseTo       = GetVal('dtLeaseTo', iRow)

                    // if (dtMoveIn === "''")
                    // {
                    //     dtMoveIn    = "";
                    // }
                    // if (dtMoveOut === "''")
                    // {
                    //     dtMoveOut   = "";
                    // }
                    // if (dtLeaseFrom === "''")
                    // {
                    //     dtLeaseFrom = "";
                    // }
                    // if (dtLeaseTo === "''")
                    // {
                    //     dtLeaseTo   = "";
                    // }

                    if (sStatus === "c")
                    {
                        sStatus = "Current";
                    }
                    else if (sStatus === "n")
                    {
                        sStatus = "Notice";
                    }
                    else if (sStatus === "f")
                    {
                        sStatus = "Future";
                    }
                    else
                    {
                        sStatus = "";
                    }

                    sUnitCode   =  GetVal('sUnitCode', iRow);
                    sFedID      = GetVal('sFedID', iRow);

                    if (sFedID.length > 0)
                    {
                        sFedID  = sFedID.replaceAll("-", "");
                    }
                    else
                    {
                        sFedID  = "";
                    }
                    
                    sLine   = "'" + sResCode + "'" + ","                                       
                            + "'" + sLName + "'" + ","                                         
                            + "'" + sPropCode + "'" + ","                                      
                            +"'" + sUnitCode + "'" + "," 			
                            + "'" + sStatus + "'" + ","                                       
                            + "'" + sFName + "'" + ","                                         
                            + "'" + dtLeaseFrom + "'" + "," 		                        
                            + "'" + dtMoveIn + "'" + "," 		                        
                            + "'" + dtMoveOut + "'" + "," 		                        
                            + "'" + dtLeaseTo + "'" + "," 		                        
                            + "0," 							
                            + GetVal('secDeposit', iRow) + "," 			
                            + "'" + dtMoveIn + "'" + "," // dtsigndate
                            + "'" + sResCode + "'" + ","                                       
                            +"'" + sUnitCode + "'" + "," 
                            + "'" + sFedID + "'" + ","
                            + "'" + GetVal('sEmail', iRow) + "'" + "," 
                            + "'" + GetVal('sPhoneNum0', iRow) + "'"     
                    sLine = sLine.replaceAll("''", "")
                    tenantWriter.write(sLine + "\n");

                    //iResCode++; // This is incremented but never used?
                }
                iRow++;
            }
            tenantWriter.close();
        }
        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }

    }

    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;

    // CREATE LEASE_CHARGES_LIST.CSV
    if (GetChkStatus("leasecharge") === 1)
    {

        sFileOut = sFilePath + "\\Lease_Charges_List.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Lease_Charges_List.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var leaseChargeWriter = fs.createWriteStream(sFileOut);

            iRow          = iFirstRow;
            iResCode      = iResCodeFirst;

            for (let fieldName of headerLine)
            {
                let index = headerLine.indexOf(fieldName);
                fieldName = fieldName.toLowerCase();

                if (fieldName.substring(0,2) === "r-")
                {

                    oChargeFields.set(index, fieldName.replaceAll("\n", "").replaceAll("\r", ""))
                }
            }

            sUtilCode	= GetFieldVal("utilOpt")
            if (sUtilCode === "r-util")
            {
                oChargeFields.set(headerLine.indexOf("r-rubs"), "r-util")
            }

            
            while (GetVal('sFName', iRow) + GetVal('sLName', iRow) != "") 
            {
                
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase();
                
                if (!(sStatus === "r"
                    || sStatus === "v"
                    || oResCodes.get(iRow) === "VACANT"))
                {
                    sResCode    = oResCodes.get(iRow);  

                    dtLeaseFrom = GetVal('dtLeaseFrom', iRow)
                    dtLeaseTo   = GetVal('dtLeaseTo', iRow)
                    

                    for (const chargeField of oChargeFields)
                    {
                        // console.log(chargeField)
                        // console.log(oXL[iRow])
                        
                        let chargeAmount    = oXL[iRow][chargeField[0]];
                        let chargeName      = chargeField[1].trim();

                        if (chargeAmount != undefined && chargeAmount != 0)
                        {

                            // Make sure concessions are stored as negative numbers
                            if (chargeName === "r-conces")
                            {
                                if (chargeAmount.toString().trim() === "")
                                {
                                    dRConces = 0.00;
                                }
                                else
                                {
                                    dRConces = parseFloat(chargeAmount);
                                }
                                
                                if (dRConces != 0.00)
                                {
                                    if (dRConces > 0.00)
                                    {
                                        chargeAmount = "-" + chargeAmount;
                                    }
                                }
                            }
                        
                            sLine	= "'" + chargeName + "'" + ","
                                    + "'" + sResCode + "'" + ","
                                    + "'" + dtLeaseFrom + "'" + ","	
                                    + "'" + dtLeaseTo + "'" + ","				
                                    + "2,"					
                                    + chargeAmount + ","				
                                    + "2"
                            sLine = sLine.replaceAll("''", "")
                            leaseChargeWriter.write(sLine + "\n");
                        }
                    }
                    //iResCode++; // This is incremented but never used?
                }
                iRow++;
            }
            leaseChargeWriter.close();
        }
        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }
    }
    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;

    // CREATE DEMOGRAPHICS.CSV
    if (GetChkStatus("demographics") === 1)
    {
        sFileOut = sFilePath + "\\Demographics.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Demographics.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var demographicsWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            sName       = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);

            while (sName.trim() != "")
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase();
                if (!(oResCodes.get(iRow) === "VACANT"
                    || sStatus === "v"))
                {

                    sResCode        = oResCodes.get(iRow);  
                    sRoomCode       = sResCode;
                    
                    if (sResCode.includes("o"))
                    {
                        sUnitCode	= GetVal('sUnitCode', iRow);
                        sUnitCode	= oUnits.get(sUnitCode.toUpperCase())
                        sResCode	= oResCodes.get(sUnitCode)
                    }

                    sLessee		= GetVal('sLessee', iRow).substring(0,1).toLowerCase();
                    sSex		= GetVal('sSex', iRow).substring(0,1).toLowerCase();
                    sMarital	= GetVal('sMarital', iRow).substring(0,1).toLowerCase();
                    dAnnual		= GetVal('dAnnual', iRow);

                    // // dMonthly never used? --> not used anymore
                    // if (dAnnual === "")
                    // {
                    //     dAnnual	 = 0.00
                    // }
                    // if (dAnnual === 0.00) 
                    // {
                    //     dMonthly = 0.00
                    // }
                    // else
                    // {
                    //     dMonthly = (Math.round((dAnnual / 12.0))).toString();
                    // }
                    
                    if (sLessee === "y") 
                    {
                        sLessee	 = "Y"
                    }
                    else
                    {
                        sLessee	 = "N"
                    }
                    
                    if (sSex === "m") 
                    {
                        sSex	 = "Male"
                    }
                    else if (sSex === "f") 
                    {
                        sSex	 = "Female"
                    }
                    else
                    {
                        sSex	 = ""
                    }
                    
                    if (sMarital === "m") 
                    {
                        sMarital = "Married"
                    }
                    else if (sMarital === "s") 
                    {
                        sMarital = "Single"
                    }
                    else
                    {
                        sMarital = ""
                    }

                    dtDOB           = GetVal('dtDOB', iRow)                                   
                    dtMoveOut       = GetVal('dtMoveOut', iRow)
                    
                    // if (dtDOB === "")
                    // {
                    //     dtDOB       = "01/01/1900";
                    // }

                    sEmployer       = GetVal('sEmployer', iRow);
                    
                    sLine		    = "'" + sResCode + "'" + "," 
                                    + "'" + sRoomCode + "'" + ","
                                    + "'" + sName + "'" + "," 
                                    + "'" + sLessee + "'" + "," 
                                    + "'" + sSex + "'" + "," 
                                    + "'" + sMarital + "'" + "," 
                                    + "'" + dtDOB + "'" + "," 
                                    + dAnnual + "," 
                                    + "'" + sEmployer + "'" + "," 
                                    + "'" + GetVal('sOccupation', iRow) + "'" + ","
                                    + "'" + dtMoveOut + "'"
                    sLine = sLine.replaceAll("''", "")
                    demographicsWriter.write(sLine + "\n");

                    //iResCode++; // iResCode isn't even in this if section?
                }
                iRow++;
                sName   = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);
            }
            demographicsWriter.close();
        }
        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }
    }
    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;

    // CREATE Roommates.CSV
    if (GetChkStatus("roommates") === 1)
    {

        sFileOut = sFilePath + "\\Roommates.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating Roommates.csv...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var roommatesWriter = fs.createWriteStream(sFileOut);

            for (const rowNum of oRoomies) {
                iRow        = rowNum[0]
                sResCode    = oResCodes.get(iRow);
                sRoomCode   = sResCode;

                if (sResCode.includes("o"))
                {
                        sUnitCode	= GetVal('sUnitCode', iRow);
                        iResCode	= oUnits.get(sUnitCode.toUpperCase())
                        sResCode	= oResCodes.get(iResCode) 
                };

                dtMoveIn        = GetVal('dtMoveIn', iRow)
                dtMoveOut       = GetVal('dtMoveOut', iRow)

                if (dtMoveIn === "")
                {
                    dtMoveIn    = GetVal('dtMoveIn', iResCode)
                }
                if (dtMoveOut === "")
                {
                    dtMoveOut   = GetVal('dtMoveOut', iResCode)
                }
                // if (dtMoveIn === "")
                // {
                //     dtMoveIn    = "01/01/1900";
                // }

                if (GetVal('iOccupantType', iRow) === "A")
                {
                    iOccupantType = 1;
                }
                else if (GetVal('iOccupantType', iRow) === "C")
                {
                    iOccupantType = 0;
                }
                else
                {
                    dtDOB = GetVal('dtDOB', iRow) != ""
                    let adultBDay = new Date(new Date().setFullYear(new Date().getFullYear() - 18));
                    let adultBDayToMMDDYYY = adultBDay.getMonth()+1 + "/" + adultBDay.getDate() + "/" + adultBDay.getFullYear();
                    if (dtDOB === "" && Date.parse(dtDOB) <= Date.parse(adultBDayToMMDDYYY))
                    {
                        iOccupantType = 1;           
                    }
                    else {
                        iOccupantType = 0;
                    }
                }

                let sRelationship = GetVal('sRelationship', iRow)
                if (sRelationship === "")
                {
                    sRelationship = "Other";
                }


                sLine       = "'" + sRoomCode + "'" + ","
                            + "'" + sResCode + "'" + ","
                            + "'" + dtMoveIn + "'" + ","
                            + "'" + sRelationship + "'" + ","
                            + iOccupantType + ","
                            + "'" + dtMoveOut + "'"
                sLine = sLine.replaceAll("''", "")
                roommatesWriter.write(sLine + "\n");
            }

            roommatesWriter.close();
        }
        if (overwriteFlag === 1)
        {
            sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
            Updater(sUpdater);
            overwriteFlag = 0;
        }
    }
    ProgressBarInc(iProgress)
    iProgress++;
    ProgressBarInc(iProgress)
    iProgress++;

    // CREATE SECDEPS.CSV
    if (GetChkStatus("secdeps") === 1)
    {
        // Create Security Deposit Charges

        sFileOut = sFilePath + "\\SecDepsCharges.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating SecDepsCharges.csv (Security Deposits)...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var secDepsWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            sName       = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);

            while (sName.trim() != "")
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase();
                secDeposit = GetVal('secDeposit', iRow);

                if (!(sStatus === "r"
                        || sStatus === "v"
                        || oResCodes.get(iRow) === "VACANT")
                    && secDeposit > 0.00)
                {
                    sResCode        = oResCodes.get(iRow); 

                    dtMoveIn        = GetVal('dtMoveIn', iRow);
                    if (dtMoveIn === "")
                    {
                        dtMoveIn    = "11/01/1997"
                    }
                    sLine		    = "'C'" + "," 
                                    + "," 
                                    + "'" + sResCode + "'" + "," 
                                    + "'" + sName + "'" + "," 
                                    + "'" + dtMoveIn + "'" + "," 
                                    + "'11/01/1997'" + "," 
                                    + "'Imp: Sec Dep'" + "," 
                                    + "'Imp: Sec Dep'" + "," 
                                    + "'" + sPropCode + "'" + "," 
                                    + secDeposit + "," 
                                    + "23300000" + "," 
                                    + "12100000" + "," 
                                    + "'r-secdep'" + "," 
                                    + ",,,,"
                    sLine = sLine.replaceAll("''", "")
                    secDepsWriter.write(sLine + "\n");

                }
                iRow++;
                sName   = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);
            }
            secDepsWriter.close();
        }

        // Create Pet Deposit Charges

        sFileOut = sFilePath + "\\PetDepsCharges.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating PetDepsCharges.csv (Pet Deposits)...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var petDepsWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            iResCode    = iResCodeFirst;
            sName       = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);

            while (sName.trim() != "")
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase(); 
                petDeposit = GetVal('petDeposit', iRow);

                if (!(sStatus === "r"
                        || sStatus === "v"
                        || oResCodes.get(iRow) === "VACANT")
                    && petDeposit > 0.00)
                {

                    sResCode        = oResCodes.get(iRow); 

                    dtMoveIn        = GetVal('dtMoveIn', iRow);
                    if (dtMoveIn === "")
                    {
                        dtMoveIn    = "11/01/1997"
                    }
                    
                    sLine		    = "'C'" + "," 
                                    + "," 
                                    + "'" + sResCode + "'" + "," 
                                    + "'" + sName + "'" + "," 
                                    + "'" + dtMoveIn + "'" + "," 
                                    + "'11/01/1997'" + ","
                                    + "'Imp: Pet Dep'" + "," 
                                    + "'Imp: Pet Dep'" + "," 
                                    + "'" + sPropCode + "'" + "," 
                                    + petDeposit + "," 
                                    + "23400000" + "," 
                                    + "12100000" + "," 
                                    + "'r-secpet'" + "," 
                                    + ",,,,"
                    sLine = sLine.replaceAll("''", "")
                    petDepsWriter.write(sLine + "\n");

                }
                iRow++;
                sName   = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);
            }
            petDepsWriter.close();
        }

        // Create Security Deposit Receipts

        sFileOut = sFilePath + "\\SecDepsReceipts.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating SecDepsReceipts.csv (Security Deposit Receipts)...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var secReceiptWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            iResCode    = iResCodeFirst;
            sName       = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);

            while (sName.trim() != "")
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase(); 
                secDeposit = GetVal('secDeposit', iRow);

                if (!(sStatus === "r"
                        || sStatus === "v"
                        || oResCodes.get(iRow) === "VACANT")
                    && secDeposit > 0.00)
                {

                    sResCode        = oResCodes.get(iRow); 

                    dtMoveIn        = GetVal('dtMoveIn', iRow);
                    if (dtMoveIn === "")
                    {
                        dtMoveIn    = "11/01/1997"
                    }
                    
                    sLine		    = "'R'" + "," 
                                    + "," 
                                    + "'" + sResCode + "'" + "," 
                                    + "'" + sName + "'" + "," 
                                    + "'" + dtMoveIn + "'" + "," 
                                    + "'11/01/1997'" + "," 
                                    + "'Imp: Sec Dep'" + "," 
                                    + "'Imp: Sec Dep'" + "," 
                                    + "'" + sPropCode + "'" + "," 
                                    + secDeposit + "," 
                                    + "23300000" + "," 
                                    + "12100000" + "," 
                                    + "'r-secdep'" + "," 
                                    + ",,,,"
                    sLine = sLine.replaceAll("''", "")
                    secReceiptWriter.write(sLine + "\n");
                }
                iRow++;
                sName   = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);
            }
            secReceiptWriter.close();
        }

        // Create Pet Deposit Receipts

        // sUpdater += "<tr><td>Creating PetDepsReceipts.csv (Pet Deposit Receipts)...</td>";
        // Updater(sUpdater);
        sFileOut = sFilePath + "\\PetDepsReceipts.csv";

        if (fs.existsSync(sFileOut))
        {
            iOverWrite = dialog.showMessageBoxSync( /* Returns 0 for yes, 1 for no */
                {
                    type: "question",
                    buttons: ["Yes", "No"],
                    message: "Overwite " + sFileOut + "?",
                    title: "File Already Exists"
                });
        }
        else 
        {
            iOverWrite  = 10;
        }

        if (iOverWrite === 0) 
        {
            fs.unlinkSync(sFileOut);
            iOverWrite  = 10;
        }

        if (iOverWrite === 10)
        {
            sUpdater += "<tr><td>Creating PetDepsReceipts.csv (Pet Deposit Receipts)...</td>";
            Updater(sUpdater);
            overwriteFlag = 1;

            fs.writeFileSync(sFileOut);
            var petReceiptWriter = fs.createWriteStream(sFileOut);

            iRow        = iFirstRow;
            iResCode    = iResCodeFirst;
            sName       = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);

            while (sName.trim() != "")
            {
                sStatus = GetVal('sStatus', iRow).substring(0,1).toLowerCase(); 
                petDeposit = GetVal('petDeposit', iRow);

                if (!(sStatus === "r"
                        || sStatus === "v"
                        || oResCodes.get(iRow) === "VACANT")
                    && petDeposit > 0.00)
                {

                    sResCode        = oResCodes.get(iRow); 

                    dtMoveIn        = GetVal('dtMoveIn', iRow);
                    if (dtMoveIn === "")
                    {
                        dtMoveIn    = "11/01/1997"
                    }
                    
                    sLine		    = "'R'" + "," 
                                    + "," 
                                    + "'" + sResCode + "'" + "," 
                                    + "'" + sName + "'" + "," 
                                    + "'" + dtMoveIn + "'" + "," 
                                    + "'11/01/1997'" + "," 
                                    + "'Imp: Pet Dep'" + "," 
                                    + "'Imp: Pet Dep'" + "," 
                                    + "'" + sPropCode + "'" + "," 
                                    + petDeposit + "," 
                                    + "23400000" + "," 
                                    + "12100000" + "," 
                                    + "'r-secpet'" + "," 
                                    + ",,,,"
                    sLine = sLine.replaceAll("''", "")
                    petReceiptWriter.write(sLine + "\n");

                }
                iRow++;
                sName   = GetVal('sFName', iRow)
                        + " "
                        + GetVal('sLName', iRow);
            }
            petReceiptWriter.close();
        }
        sUpdater += "<td>&nbsp;&nbsp;&nbsp;Done</td></tr>";
        Updater(sUpdater);
    }

    ProgressBarInc(iProgress);
    iProgress++;
    ProgressBarInc(iProgress);
    iProgress++; 

    sUpdater    += "<tr><td><b class='FilesDone'>File Extraction Complete</td></tr>"
                +  "<tr><td>Files saved to:</td></tr>"
                +  "<tr><td>" + sFilePath + "</td></tr></table>"
    Updater(sUpdater)

    return;
}