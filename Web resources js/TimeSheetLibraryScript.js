// ********** Main Function **********

TimeSheetLibraryScript = {
    ExportExcel: function(){
        var pageInput = {
            pageType: "webresource",
            webresourceName: "new_ExportExcel"
          };
          var navigationOptions = {
            target: 2,
            width: 450, // value specified in pixel
            height: 450, // value specified in pixel
            position: 1,
            title: "Monthly Report"
          };
          Xrm.Navigation.navigateTo(pageInput,navigationOptions)
    },


    PopupConfirmSubmit: async function(gridContext,executionContext){
        var selectedRecordIDs = executionContext;
        var confirmStrings = { text:"This is a confirmation.", title:"Confirmation Dialog", confirmButtonLabel:"Submit"};
        var confirmOptions = { height: {value: 50, unit:"%"}, width: {value: 50, unit:"%"} };

        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then( 
        async (success) => {
            if (success.confirmed) {
                Xrm.Utility.showProgressIndicator('Please waiting a moment...');
                try {
                    const updatePromises = [];
                    for (const IDRecord of selectedRecordIDs) {
                        updatePromises.push(retrieveRecord(IDRecord));
                    }
                    await Promise.all(updatePromises);
                } catch (error) {
                } finally {
                    // Reload view 
                    gridContext.refresh(true);
                    Xrm.Utility.closeProgressIndicator();
                }
            }
            else {
                console.log("Dialog closed using Cancel button or X.");
            }

        });
    },


    MainGridTimeSheetApproveBtn: async function (gridContext, executionContext){
        var selectedRecordIDs = executionContext;
        var confirmStrings = { text:"This is a confirmation.", title:"Confirmation Dialog", confirmButtonLabel:"Approve"};
        var confirmOptions = { height: {value: 50, unit:"%"}, width: {value: 50, unit:"%"} };

        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then( 
        async (success) => {
            if (success.confirmed) {
                Xrm.Utility.showProgressIndicator('Please waiting a moment...');
                try {
                    const updatePromises = [];
                    for (const IDRecord of selectedRecordIDs) {
                        const data =
                            {
                                "new_status": TimeSheetCurrentStatusEnum.Approved
                            }
                        updatePromises.push(updateRecord(IDRecord, data));
                    }
                    await Promise.all(updatePromises);
                } catch (error) {
                } finally {
                    // Reload view 
                    gridContext.refresh(true);
                    Xrm.Utility.closeProgressIndicator();
                }
            }
            else {
                console.log("Dialog closed using Cancel button or X.");
            }

        });
    },


    MainGridTimeSheetRejectBtn: function (executionContext){
        var selectedRecordIDs = executionContext;
        var pageInput = {
            pageType: "custom",
            name: "new_insertrejectreason_d6b15",
            recordId: selectedRecordIDs
        };
    
        var navigationOptions = {
            target: 2,
            position: 1,
            height: {value: 50, unit:"%"},
            width: {value: 50, unit:"%"},
            title: "Reject Confirmation "
        };
    
        Xrm.Navigation.navigateTo(pageInput, navigationOptions)
    },


    checkVisibleApproveRejectFlyOut: async function (executionContext){
        var selectedRecordIDs = executionContext;
        console.log(selectedRecordIDs);
        var isCheck = false;
        var userSettings = await Xrm.Utility.getGlobalContext().userSettings; // userSettings is an object with user information.
        var current_User_Id = String(userSettings.userId).toLowerCase().replace(/[{}]/g, ''); // The user's unique id
        console.log(current_User_Id);
        for (const IDRecord of selectedRecordIDs) {
            const result = await Xrm.WebApi.retrieveRecord("new_timesheet", IDRecord, "?$select=_ownerid_value,_createdby_value");
            let ownerIdRetrieve = result._ownerid_value;
            let createdByRetrieve = result._createdby_value; 
            console.log(current_User_Id);
            console.log(ownerIdRetrieve);
            console.log(current_User_Id === ownerIdRetrieve);
            if(ownerIdRetrieve != createdByRetrieve && current_User_Id === ownerIdRetrieve){
                isCheck = true;
            }else{
                isCheck = false
                break;
            }
        }
        console.log(isCheck);
        return isCheck;
    }
}

// ********** Sub Function **********

const TimeSheetCurrentStatusEnum = {
    Submitted: 0,
    Approved: 1,
    Rejected: 2,
    Draft: 3
};

const retrieveRecord = async (IDRecord) =>{
    const result = await Xrm.WebApi.retrieveRecord("new_timesheet", IDRecord, "?$select=_ownerid_value");
    const ownerIdRetrieve = result._ownerid_value;
    await shareUserToRecord(IDRecord, ownerIdRetrieve); 
    return changeOwner(IDRecord, ownerIdRetrieve);
}

const shareUserToRecord = async (IDRecord, ownerIdRetrieve) => {
    var target = {
        new_timesheetid: IDRecord, // Pass your respective Record id
        '@odata.type' : "Microsoft.Dynamics.CRM.new_timesheet" // Pass your respective Entity Logical name record type
    };
    // Principle to User
    var principalAccess = {
        Principal: {
            '@odata.type': "Microsoft.Dynamics.CRM.systemuser",
            systemuserid: ownerIdRetrieve
        },
        AccessMask: "ReadAccess"
        //Access Mask Sample "ReadAccess, WriteAccess, AppendAccess, AppendToAccess, CreateAccess, DeleteAccess, ShareAccess, AssignAccess"
    };

    var parameters = {
        Target: target,
        PrincipalAccess: principalAccess
    };

    var context;

    if (typeof GetGlobalContext === 'function') {
            context = GetGlobalContext();
        } else {
            context = Xrm.Page.context;
        }

    var req = new XMLHttpRequest();
    req.open("POST",context.getClientUrl() + "/api/data/v9.0/GrantAccess", true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
                //Success - No Return Data - Do Something
            } else {
                var errorText = this.responseText;
                //error response
            }
        }
    };
    req.send(JSON.stringify(parameters));
}

const changeOwner = async (IDRecord, ownerIdRetrieve) => {
    const result = await Xrm.WebApi.retrieveRecord("systemuser", ownerIdRetrieve, "?$select=_parentsystemuserid_value");
    const ownerIdRetrieveManager = result._parentsystemuserid_value;
    const data =
            {
                "ownerid@odata.bind": `systemusers(${ownerIdRetrieveManager})`,
                "new_status": TimeSheetCurrentStatusEnum.Submitted
            }

    await updateRecord(IDRecord, data)
}

const updateRecord = async (IDRecord,data) => {
    const result = await Xrm.WebApi.updateRecord("new_timesheet", IDRecord, data)
    return result
}

const GetReportExcel =  () => {

    var userSettings =  Xrm.Utility.getGlobalContext().userSettings; // userSettings is an object with user information.
    var current_User_Id = String(userSettings.userId).toLowerCase().replace(/[{}]/g, ''); // The user's unique id
    console.log(current_User_Id)

    Xrm.WebApi.retrieveMultipleRecords("new_timesheet", `?$filter=new_status eq ${TimeSheetCurrentStatusEnum.Approved}` ).then(
        function success(result) {
            let ManagerView = []
            let EmployeeView=[]
            for (var i = 0; i < result.entities.length; i++) {
                console.log(result.entities[i]);
                if(current_User_Id == result.entities[i]._ownerid_value ){
                    ManagerView.push(result.entities[i]);
                }
                else{
                    if(current_User_Id == result.entities[i]._createdby_value){
                        EmployeeView.push(result.entities[i]);
                    }
                }
            }


            
            let workbook = XLSX.utils.book_new();
            if (ManagerView != null) {
                let worksheet = XLSX.utils.json_to_sheet(ManagerView);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'ManagerView');
            }

            if (EmployeeView != null) {
                let worksheet = XLSX.utils.json_to_sheet(EmployeeView);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'EmployeeView');
            }

            XLSX.writeFile(workbook, 'output.xlsx');
            window.close();
            console.log("manager: ", ManagerView);
            console.log("employee: ", EmployeeView);
            
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}