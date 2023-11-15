const TimeSheetCurrentStatusEnum = {
    Draft: 1,
    Submitted: 2,
    Approved: 3,
    Rejected: 0,
    Paid: 4,
  };

TimeSheetLibraryScript = {
    CheckVisibleApproveRejectFlyOut: async function (executionContext){
        const selectedRecordIDs = executionContext;
        let isCheck = false;
        const userSettings = await Xrm.Utility.getGlobalContext().userSettings; // userSettings is an object with user information.
        const current_User_Id = String(userSettings.userId).toLowerCase().replace(/[{}]/g, ''); // The user's unique id
        const filterOption = `?$select=_ownerid_value,_createdby_value&$filter=Microsoft.Dynamics.CRM.In(PropertyName='dmo_dmotimesheetid',PropertyValues=${JSON.stringify(selectedRecordIDs)})`;
        const result = await Xrm.WebApi.retrieveMultipleRecords("dmo_dmotimesheet", filterOption);
        console.log(result);
        isCheck = result.entities.every(x => x._ownerid_value != x._createdby_value && current_User_Id === x._ownerid_value)
        console.log(isCheck)
        return isCheck;
    },

    CheckVisibleSubmit: async function(executionContext){
        const selectedRecordIDs = executionContext;
        let isCheck = false;
        const filterOption = `?$select=dmo_currentstatus,&$filter=Microsoft.Dynamics.CRM.In(PropertyName='dmo_dmotimesheetid',PropertyValues=${JSON.stringify(selectedRecordIDs)})`;
        const result = await Xrm.WebApi.retrieveMultipleRecords("dmo_dmotimesheet", filterOption);
        isCheck = result.entities.every(x => x.dmo_currentstatus === TimeSheetCurrentStatusEnum.Draft)
        return isCheck;
    }

}