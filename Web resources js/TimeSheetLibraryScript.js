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




function groupByDayAndCreatedBy(rawData, month, year) {
    var daysInMonth = new Date(year, month, 0).getDate();
    var result = [];

    // Tạo một object tạm để lưu trữ dữ liệu khi nhóm
    var temp = {};
    rawData.forEach(function(item) {
        // Lấy ngày từ trường 'new_date'
        var date = new Date(item.new_date);
        
        // Kiểm tra xem ngày này có thuộc tháng và năm đang xét hay không
        if (date.getMonth() + 1 === month && date.getFullYear() === year) {
            // Nếu chưa có dữ liệu cho 'createdby' này, khởi tạo một object mới
            if (!temp[item._createdby_value]) {
                temp[item._createdby_value] = { createdby: item._createdby_value };
                for (var i = 1; i <= daysInMonth; i++) {
                    temp[item._createdby_value][i] = null;
                }
            }

            // Cộng dồn 'new_duration' vào ngày tương ứng trong kết quả
            if (temp[item._createdby_value][date.getDate()]) {
                temp[item._createdby_value][date.getDate()] += item.new_duration;
            } else {
                temp[item._createdby_value][date.getDate()] = item.new_duration;
            }
        }
    });

    // Chuyển dữ liệu từ object tạm sang mảng kết quả
    for (var key in temp) {
        result.push(temp[key]);
    }

    return result;
}




// Export data from TimeSheet to Excel
const GetReportExcel =  async (val) => {
    console.log("Month and year: ",val);
    const VendorApi  = await(await Xrm.WebApi.retrieveMultipleRecords("new_vendor")) ;

    const VendorEmployeeApi= await Xrm.WebApi.retrieveMultipleRecords("new_vendoremployees");

    const TimeSheetApi = await Xrm.WebApi.retrieveMultipleRecords("new_timesheet", `?$filter=new_status eq ${TimeSheetCurrentStatusEnum.Approved}` );

    const TimeSheetData = [];
    const VendorEmployeeData = [];
    const VendorData = [];

    for (var i = 0; i < VendorApi.entities.length; i++) {
        TimeSheetData.push(VendorApi.entities[i])
    }

    for (var i = 0; i < VendorEmployeeApi.entities.length; i++) {
        VendorEmployeeData.push(VendorEmployeeApi.entities[i])
    }

    for (var i = 0; i < TimeSheetApi.entities.length; i++) {
        VendorData.push(TimeSheetApi.entities[i])
    }
  

    var jointabledata = TimeSheetData.map(ts => {
        var ve = VendorEmployeeData.find(ve => ve._new_employee_value === ts._createdby_value);
        var v = VendorData.find(v => v.new_vendorid === ve._new_vendor_value );
        return {...ts, ...ve, ...v};
    });

    console.log(JSON.stringify(jointabledata));

    var parts = val.split("-");
    var year = Number(parts[0]);
    var month = Number(parts[1]);

    var userSettings =  Xrm.Utility.getGlobalContext().userSettings; // userSettings is an object with user information.
    var current_User_Id = String(userSettings.userId).toLowerCase().replace(/[{}]/g, ''); // The user's unique id
    console.log(current_User_Id)

    const result = await Xrm.WebApi.retrieveMultipleRecords("new_timesheet", `?$filter=new_status eq ${TimeSheetCurrentStatusEnum.Approved}` );

            let ManagerView = []
            let EmployeeView = []

            for (var i = 0; i < result.entities.length; i++) {
                //console.log(result.entities[i]);
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

                let result = groupByDayAndCreatedBy(ManagerView,month,year)
                let worksheet = XLSX.utils.json_to_sheet(result);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'ManagerView');
            }

            if (EmployeeView != null) {
                let result1 = groupByDayAndCreatedBy(EmployeeView,month,year)
                let worksheet = XLSX.utils.json_to_sheet(result1);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'EmployeeView');
            }
            // Calculate Time Working

            XLSX.writeFile(workbook, 'output.xlsx');
            window.close();
           // console.log("manager: ", ManagerView);
            //console.log("employee: ", EmployeeView);


}

