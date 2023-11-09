// ********** Main Function **********

TimeSheetLibraryScript = {
    ExportExcel: function(){
        var pageInput = {
            pageType: "webresource",
            webresourceName: "new_MonthPicker"
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

const CurrenStatus = {
    Office: 0,
    Before22h: 1,
    After22h: 2

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


function OtHour(item){
	let value = 1;
	let date = new Date(item.new_date);
	let day = date.getDay();
	if(day >= 1 && day <= 5 && item.new_currentstatus_ == CurrenStatus.Office)
		value = item.new_duration * item.new_weekdayoffice;
	else if (day >= 1 && day <= 5 && item.new_currentstatus_ ==  CurrenStatus.Before22h)
		value = item.new_duration * item.new_weekdaybefore22h;
	else if (day >= 1 && day <= 5 && item.new_currentstatus_ ==  CurrenStatus.After22h)
		value = item.new_duration * item.new_weekdayafter22h;
	else if ((day === 0 || day === 6) && item.new_currentstatus_ ==  CurrenStatus.Before22h)
		value = item.new_duration * item.new_weekendbefore22h;
	else if ((day === 0 || day === 6) && item.new_currentstatus_ ==  CurrenStatus.After22h)
		value = item.new_duration * item.new_weekdayafter22h;
	return value;
}

function groupByDayAndCreatedBy(rawData, month, year) {
    var daysInMonth = new Date(year, month, 0).getDate();
    var totalOfficeHours_index = daysInMonth + 1
    var totalBefore22h_index = totalOfficeHours_index + 1
    var totalAfter22h_index = totalBefore22h_index + 1
    var result = [];

    // Tạo một object tạm để lưu trữ dữ liệu khi nhóm
    var temp = {};
    rawData.forEach(function (item) {
        // Lấy ngày từ trường 'new_date'
        var date = new Date(item.new_date);

        // Kiểm tra xem ngày này có thuộc tháng và năm đang xét hay không
        if (date.getMonth() + 1 === month && date.getFullYear() === year) {
            // Nếu chưa có dữ liệu cho '_createdby_value' này, khởi tạo một object mới
            if (!temp[item._createdby_value]) {
                temp[item._createdby_value] = {};
                for (var i = 1; i <= totalAfter22h_index; i++) {
                    temp[item._createdby_value][i] = null;
                }
            }

        // Cộng dồn 'new_duration' vào ngày tương ứng trong kết quả
            if (temp[item._createdby_value][date.getDate()]) {
                temp[item._createdby_value][date.getDate()] += item.new_duration;
            } else {
                temp[item._createdby_value][date.getDate()] = item.new_duration;
            }


            var status = item.new_currentstatus_;
            if(status == CurrenStatus.Office){
                if (temp[item._createdby_value][totalOfficeHours_index]) {
                    temp[item._createdby_value][totalOfficeHours_index] += OtHour(item);
                } else {
                    temp[item._createdby_value][totalOfficeHours_index] = OtHour(item);
                }
            }
            if(status == CurrenStatus.After22h){
                if (temp[item._createdby_value][totalAfter22h_index]) {
                    temp[item._createdby_value][totalAfter22h_index] += OtHour(item);
                } else {
                    temp[item._createdby_value][totalAfter22h_index] = OtHour(item);
                }
            }
            if(status == CurrenStatus.Before22h){
                if (temp[item._createdby_value][totalBefore22h_index]) {
                    temp[item._createdby_value][totalBefore22h_index] += OtHour(item);
                } else {
                    temp[item._createdby_value][totalBefore22h_index] = OtHour(item);
                }
            }
        }
    });

    //console.log(temp)
    // Chuyển dữ liệu từ object tạm sang mảng kết quả
    for (var key in temp) {
        var obj = { Employee: key };
        for (var i = 1; i <= daysInMonth; i++) {
            obj[i] = temp[key][i];
        }
        obj["TotalOfficeHours"] = temp[key][totalOfficeHours_index];
        obj["TotalBefore22h"] = temp[key][totalBefore22h_index];
        obj["TotalAfter22h"] = temp[key][totalAfter22h_index];
        obj["TotalWorkingHours"] = obj["TotalOfficeHours"] + obj["TotalBefore22h"] + obj["TotalAfter22h"]
        result.push(obj);
    }
    return result;
}





// Export data from TimeSheet to Excel
// Get the val parameter from ExportExcel.html web resource
const GetReportExcel =  async (val) => {
    console.log("Month and year: ",val);
    const date = new Date(val.toLocaleString());
    const month = Number(date.getMonth() + 1); // getMonth() returns month from 0 to 11, so we add 1 to get the correct month number
    const year = Number(date.getFullYear());
    console.log("month,year", month,year)
    const VendorApi  = await(await Xrm.WebApi.retrieveMultipleRecords("new_vendor")) ;

    const VendorEmployeeApi= await Xrm.WebApi.retrieveMultipleRecords("new_vendoremployees");

    const TimeSheetApi = await Xrm.WebApi.retrieveMultipleRecords("new_timesheet", `?$filter=new_status eq ${TimeSheetCurrentStatusEnum.Approved}`);

    const TimeSheetData = [];
    const VendorEmployeeData = [];
    const VendorData = [];

    for (var i = 0; i < TimeSheetApi.entities.length; i++) {
        TimeSheetData.push(TimeSheetApi.entities[i])
    }

    for (var i = 0; i < VendorEmployeeApi.entities.length; i++) {
        VendorEmployeeData.push(VendorEmployeeApi.entities[i])
    }

    for (var i = 0; i < VendorApi.entities.length; i++) {
        VendorData.push(VendorApi.entities[i])
    }

        // join 3 table TimeSheet, Vendor, VendorEmployee
        var jointabledata = TimeSheetData.map(ts => {
            var ve = VendorEmployeeData.find(ve => ve._new_employee_value === ts._createdby_value);
            var v = VendorData.find(v => v.new_vendorid === ve._new_vendor_value );
            return { ...ts, ...ve, ...v, _createdby_value: ts._createdby_value, _ownerid_value: ts._ownerid_value };
        });

    console.log(JSON.stringify(jointabledata));

    var userSettings =  Xrm.Utility.getGlobalContext().userSettings; // userSettings is an object with user information.
    var current_User_Id = String(userSettings.userId).toLowerCase().replace(/[{}]/g, ''); // The user's unique id
    console.log(current_User_Id)

        let ManagerView = []
        let EmployeeView = []

        // Tách join table ra làm 2 view : view employee và view manager
        for (var i = 0; i < jointabledata.length; i++) {

            if(current_User_Id == jointabledata[i]._ownerid_value ){
                ManagerView.push(jointabledata[i]);
            }
            else{
                if(current_User_Id == jointabledata[i]._createdby_value){
                    EmployeeView.push(jointabledata[i]);
                }
            }
        }
        let workbook = XLSX.utils.book_new();

        if (ManagerView.length > 0) {
            let result = groupByDayAndCreatedBy(ManagerView,month,year)
            let worksheet = XLSX.utils.json_to_sheet(result);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'ManagerView');
        }

        if (EmployeeView.length > 0) {
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

