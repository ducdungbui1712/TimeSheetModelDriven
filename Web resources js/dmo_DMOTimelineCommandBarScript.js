/**
 * The LookupType type
 * @typedef {Object} LookupType
 * @property {string} entityType
 * @property {string} id
 * @property {string} name
 */

/**
 * The TimeRange object type
 * @typedef {Object} TimeRange
 * @property {Date} start
 * @property {Date} end
 */

/**
 * The SelectedControlAllItemRef
 * @typedef {Object}  SelectedControlAllItemRef
 * @property {string} Id
 * @property {string} Name
 * @property {number} TypeCode
 * @property {string} TypeName
 */

var Sdk = window.Sdk || {};


//define send notify object.

var DMONotifty = window.DMONotifty || {};
DMONotifty.SendAppNotificationRequest = function (
  title,
  recipient,
  body,
  priority,
  iconType,
  toastType,
  expiry,
  overrideContent,
  actions) {
  this.Title = title;
  this.Recipient = recipient;
  this.Body = body;
  this.Priority = priority;
  this.IconType = iconType;
  this.ToastType = toastType;
  this.Expiry = expiry;
  this.OverrideContent = overrideContent;
  this.Actions = actions;
};

DMONotifty.SendAppNotificationRequest.prototype.getMetadata = function () {
  return {
    boundParameter: null,
    parameterTypes: {
      "Title": {
        "typeName": "Edm.String",
        "structuralProperty": 1
      },
      "Recipient": {
        "typeName": "mscrm.systemuser",
        "structuralProperty": 5
      },
      "Body": {
        "typeName": "Edm.String",
        "structuralProperty": 1
      },
      "Priority": {
        "typeName": "Edm.Int",
        "structuralProperty": 1
      },
      "IconType": {
        "typeName": "Edm.Int",
        "structuralProperty": 1
      },
      "ToastType": {
        "typeName": "Edm.Int",
        "structuralProperty": 1
      },
      "Expiry": {
        "typeName": "Edm.Int",
        "structuralProperty": 1
      },
      "OverrideContent": {
        "typeName": "mscrm.expando",
        "structuralProperty": 5
      },
      "Actions": {
        "typeName": "mscrm.expando",
        "structuralProperty": 5
      },
    },
    operationType: 0,
    operationName: "SendAppNotification",
  };
};

// A namespace defined for the sample code
// As a best practice, you should always define
// a unique namespace for your libraries
var DMOTimelineCommandBarScript = window.DMOTimelineCommandBarScript || {};
(function () {
  // Define some global variables
  // var currentUserName = Xrm.Utility.getGlobalContext().userSettings.userName; // get current user name
  const TimeshhetCurrentStatusEnum = {
    Draft: 1,
    Submitted: 2,
    Approved: 3,
    Rejected: 0,
    Paid: 4,
  };
  /**
   *
   * @param {SelectedControlAllItemRef} record
   * @param {string} loggedinUserId
   */

  const shareRecordToCurrentUserIsReadOnly = (record, loggedinUserId) => {
    var target = {
      [record.TypeName + 'id']: record.Id,
      //put <other record type>id and Guid of record to share here
      '@odata.type': `Microsoft.Dynamics.CRM.${record.TypeName}`,
      //replace account with other record type
    };

    var principalAccess = {
      Principal: {
        systemuserid: loggedinUserId,
        //put teamid here and Guid of team if you want to share with team
        '@odata.type': 'Microsoft.Dynamics.CRM.systemuser',
        //put team instead of systemuser if you want to share with team
      },
      AccessMask: 'ReadAccess',
      //full list of privileges is "ReadAccess, WriteAccess, AppendAccess, AppendToAccess, CreateAccess, DeleteAccess, ShareAccess, AssignAccess"
    };
    var parameters = {
      Target: target,
      PrincipalAccess: principalAccess,
    };

    var context = Xrm.Page.context;;

    // if (typeof GetGlobalContext === 'function') {
    //   context = GetGlobalContext();
    // } else {
    //   context = Xrm.Page.context;
    // }
    return new Promise((resolve, reject) => {
      var req = new XMLHttpRequest();
      req.open('POST', context.getClientUrl() + '/api/data/v9.0/GrantAccess', true);
      req.setRequestHeader('OData-MaxVersion', '4.0');
      req.setRequestHeader('OData-Version', '4.0');
      req.setRequestHeader('Accept', 'application/json');
      req.setRequestHeader('Content-Type', 'application/json; charset=utf-8');
      req.onreadystatechange = function () {
        if (this.readyState === 4) {
          req.onreadystatechange = null;
          if (this.status === 204) {
            //Success - No Return Data - Do Something
            resolve();
          } else {
            var errorText = this.responseText;
            //Error and errorText variable contains an error - do something with it
            reject(errorText);
          }
        }
      };
      req.send(JSON.stringify(parameters));
    });
  };

  /**
   *
   * @param {SelectedControlAllItemRef} record
   * @param {string} parentUserId
   */
  const changeOwnerRecordToManager = (record, parentUserId) => {
    return new Promise((resolve, reject) => {
      Xrm.WebApi.online
        .updateRecord(record.TypeName, record.Id, {
          'ownerid@odata.bind': `/systemusers(${parentUserId})`,
          dmo_currentstatus: TimeshhetCurrentStatusEnum.Submitted,
        })
        .then(
          function success(result) {
            console.log('ChangeOwnerRecordToManager sucess');
            // perform operations on record update
            resolve();
          },
          function (error) {
            console.log(error.message);
            // handle error conditions
            reject(error);
          }
        );
    });
  };

  /**
   *
   * @param {SelectedControlAllItemRef} record
   * @param {number} status
   */
  const changeRecordStatus = (record, status) => {
    return new Promise((resolve, reject) => {
      Xrm.WebApi.online
        .updateRecord(record.TypeName, record.Id, {
          dmo_currentstatus: status,
        })
        .then(
          function success(result) {
            // perform operations on record update
            resolve();
          },
          function (error) {
            console.log(error.message);
            // handle error conditions
            reject(error);
          }
        );
    });
  };
  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
//   const onConfirmSubmit = async (records) => {
//     const loggedinUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace('{', '').replace('}', '');
//     try {
//       const result = await Xrm.WebApi.online.retrieveRecord(
//         'systemuser',
//         loggedinUserId,
//         '?$select=systemuserid,domainname,parentsystemuserid,_parentsystemuserid_value'
//       );
//       const parentUserId = result['_parentsystemuserid_value'];
//       if (!parentUserId) {
//         Xrm.Navigation.openAlertDialog('You do not have a Reporting Manager to submit your timesheet, please contact the DMO admin for assistance.');
//         return;
//       }
//       const promisesShare = [];
//       for (const iterator of records) {
//         promisesShare.push(shareRecordToCurrentUserIsReadOnly(iterator, loggedinUserId));
//       }
//       await Promise.all(promisesShare);
//       const promiseChangeOwner = [];
//       for (const iterator of records) {
//         promiseChangeOwner.push(changeOwnerRecordToManager(iterator, parentUserId));
//       }
//       await Promise.all(promiseChangeOwner);
//       this.sendInAppNotification({
//         title: "Aprrove Pending Timesheets",
//         body: "You have a timesheet that needs to be approved, please check it.",
//         type: 'info',
//         targetUserId: parentUserId,
//         entityType: records[0].TypeName,
//         url: "?pagetype=entitylist&etn=dmo_dmotimesheet&viewid=7c873025-3e72-ee11-9ae7-00224808d9a4&viewType=1039"
//       });
//     } catch (err) {
//       Xrm.Navigation.openAlertDialog(err.message);
//     }
//   };

  const onConfirmSubmit = async (records) => {

    const loggedinUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace('{', '').replace('}', '');
    try {
      const result = await Xrm.WebApi.online.retrieveRecord(
        'systemuser',
        loggedinUserId,
        '?$select=systemuserid,domainname,parentsystemuserid,_parentsystemuserid_value'
      );
      const parentUserId = result['_parentsystemuserid_value'];

      // Retrieve dmo_date & dmo_name of multiple records id.
      const recordIdArray = records.map(x => x.Id);
      const filterOption = `?$select=dmo_date,dmo_name&$filter=Microsoft.Dynamics.CRM.In(PropertyName='dmo_dmotimesheetid',PropertyValues=${JSON.stringify(recordIdArray)})`;
      const result1 = await Xrm.WebApi.retrieveMultipleRecords("dmo_dmotimesheet", filterOption);
      
      // Calculate expire days
      let today = new Date();
      let expireDay = [];
      result1.entities.forEach(element => {
        let oldDay = new Date(element.dmo_date);
        let differenceInTime = today.getTime() - oldDay.getTime();
        let differenceInDays = differenceInTime / (1000 * 3600 * 24);
        if(differenceInDays > 5){
          expireDay.push(element.dmo_name)
        }
      });

      // Check condition for expire days
      if(expireDay.length > 0 ){
        console.log("aaaa",expireDay)
        var alertStrings = { confirmButtonLabel: "OK", text: `TimeSheets over 5 days ${expireDay} `, title: "Expire TimeSheet" };
        var alertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions)
      }else{

        if (!parentUserId) {
          Xrm.Navigation.openAlertDialog('You do not have a Reporting Manager to submit your timesheet, please contact the DMO admin for assistance.');
          return;
        }
        const promisesShare = [];
        for (const iterator of records) {
          promisesShare.push(shareRecordToCurrentUserIsReadOnly(iterator, loggedinUserId));
        }
        await Promise.all(promisesShare);
        const promiseChangeOwner = [];
        for (const iterator of records) {
          promiseChangeOwner.push(changeOwnerRecordToManager(iterator, parentUserId));
        }
  
        await Promise.all(promiseChangeOwner);
        this.sendInAppNotification({
          title: "Aprrove Pending Timesheets",
          body: "You have a timesheet that needs to be approved, please check it.",
          type: 'info',
          targetUserId: parentUserId,
          entityType: records[0].TypeName,
          url: "?pagetype=entitylist&etn=dmo_dmotimesheet&viewid=7c873025-3e72-ee11-9ae7-00224808d9a4&viewType=1039"
        });
      }
    } catch (err) {
      Xrm.Navigation.openAlertDialog(err.message);
    }
  };


  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
  const onApproveTimeSheet = async (records) => {
    const promisesApproves = [];
    try {
      for (const iterator of records) {
        promisesApproves.push(changeRecordStatus(iterator, TimeshhetCurrentStatusEnum.Approved));
      }
      await Promise.all(promisesApproves);

      const timeSheetIds = records.map(x => x.Id);
      const entityLogicalName = records[0].TypeName;
      const filterOption = `?$select=dmo_dmotimesheetid,_createdby_value&$filter=Microsoft.Dynamics.CRM.In(PropertyName=@p1,PropertyValues=@p2)&@p1='dmo_dmotimesheetid'&@p2=${JSON.stringify(timeSheetIds)}`;
      const result = await Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, filterOption);

      const userSendNotify = [...new Set(result.entities.map(x => x._createdby_value))];
      for (const iterator of userSendNotify) {
        this.sendInAppNotification({
          title: "Timesheet Approved",
          body: "Your timesheet submission has been approved. Thank you for submitting your timesheet on time.",
          type: 'success',
          targetUserId: iterator,
          entityType: records[0].TypeName,
          url: "?pagetype=entitylist&etn=dmo_dmotimesheet&viewid=d5ae8e45-0ae2-4651-83f8-53098b8ca511&viewType=1039"
        });
      }

    } catch (err) {
      Xrm.Navigation.openAlertDialog(err.message);
    }
  };

  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
  const onRejectTimeSheet = async (records) => {
    const promisesRejects = [];
    try {
      for (const iterator of records) {
        promisesRejects.push(changeRecordStatus(iterator, TimeshhetCurrentStatusEnum.Rejected));
      }
      await Promise.all(promisesRejects);
      // setTimeout(() => {
      //   Xrm.Utility.refreshParentGrid({ entityType: records[0].TypeName, id: records[0].Id });
      // }, 100);
    } catch (err) {
      Xrm.Navigation.openAlertDialog(err.message);
    }
  };

  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
  this.onClickSubmitCommand = async function (records, gridContext) {
    console.log('data', records);
    const confirmStrings = { text: 'Are you sure to submit this records', title: 'Confirmation submit timesheets' };
    const confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(async function (success) {
      if (success.confirmed) {
        Xrm.Utility.showProgressIndicator('Please waiting a moment...');
        try {
          await onConfirmSubmit(records);
        } catch (error) {
          Xrm.Navigation.openAlertDialog(error.message);
        } finally {
          gridContext?.refresh();
          gridContext?.refreshRibbon();
          Xrm.Utility.closeProgressIndicator();
        }
      } else console.log('Dialog closed using Cancel button or X.');
    });
  };


  this.onClickSubmitCommandForm = async function (formContext) {

    const confirmStrings = { text: 'Are you sure to submit this records', title: 'Confirmation submit timesheets' };
    const confirmOptions = { height: 200, width: 450 };
    /**
     * @type {SelectedControlAllItemRef[]}
     */
    const records = [];
    const id = formContext._data._entity._entityId.guid;
    const entityName = formContext._data._entity._entityType;
    records.push({
      Id: id,
      TypeName: entityName
    });
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(async function (success) {
      if (success.confirmed) {
        Xrm.Utility.showProgressIndicator('Please waiting a moment...');
        try {
          await onConfirmSubmit(records);
          formContext?.data.refresh()
        } catch (error) {
          Xrm.Navigation.openAlertDialog(error.message);
        } finally {
          Xrm.Utility.closeProgressIndicator();
        }
      } else console.log('Dialog closed using Cancel button or X.');
    });
  };

  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
  this.onClickApproveCommand = async function (records, gridContext) {
    console.log('data', records);
    const confirmStrings = { text: 'Are you sure to approve this records', title: 'Confirmation approve timesheets' };
    const confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(async function (success) {
      if (success.confirmed) {
        Xrm.Utility.showProgressIndicator('Please waiting a moment...');
        try {
          await onApproveTimeSheet(records);
          gridContext?.refresh();
          gridContext?.refreshRibbon();
        } catch (error) {
          Xrm.Navigation.openAlertDialog(error.message);
        } finally {
          Xrm.Utility.closeProgressIndicator();
        }
      } else console.log('Dialog closed using Cancel button or X.');
    });
  };
  this.onClickApproveCommandForm = async function (formContext) {
    const confirmStrings = { text: 'Are you sure to approve this records', title: 'Confirmation approve timesheets' };
    const confirmOptions = { height: 200, width: 450 };
    /**
     * @type {SelectedControlAllItemRef[]}
     */
    const records = [];
    const id = formContext._data._entity._entityId.guid;
    const entityName = formContext._data._entity._entityType;
    records.push({
      Id: id,
      TypeName: entityName
    });
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(async function (success) {
      if (success.confirmed) {
        Xrm.Utility.showProgressIndicator('Please waiting a moment...');
        try {
          await onApproveTimeSheet(records);
          formContext?.data.refresh()
        } catch (error) {
          Xrm.Navigation.openAlertDialog(error.message);
        } finally {
          Xrm.Utility.closeProgressIndicator();
        }
      } else console.log('Dialog closed using Cancel button or X.');
    });
  };


  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
  this.onClickRejectCommand = async function (records, gridContext) {
    const recordIds = records.map(x => x.Id).join(',')
    const pageInputOption = {
      pageType: 'custom',
      name: 'dmo_timesheetpopupreject_40e9b',
      recordId: recordIds
    };
    const confirmOptions = {
      target: 2,
      height: 400,
      width: 550,
      position: 1,
      title: "Reject submit Timesheets"
    };
    Xrm.Navigation.navigateTo(pageInputOption, confirmOptions).then(async function (success) {
      gridContext?.refresh();
      gridContext?.refreshRibbon();
    });
  };
  this.onClickRejectCommandForm = async function (formContext) {
    const recordId = formContext._data._entity._entityId.guid;
    const pageInputOption = {
      pageType: 'custom',
      name: 'dmo_timesheetpopupreject_40e9b',
      recordId: recordId
    };
    const confirmOptions = {
      target: 2,
      height: 400,
      width: 550,
      position: 1,
      title: "Reject submit Timesheets"
    };
    Xrm.Navigation.navigateTo(pageInputOption, confirmOptions).then(async function (success) {
      formContext?.data.refresh()
    });
  };
  /**
   * @typedef SendInAppNotifyOption
   * @property {string} title
   * @property {string} body
   * @property {string} entityType
   * @property {string?} url
   * @property { "info" | "error" | "success"} type
   * @property {string} targetUserId
   */

  const IconType = {
    info: 100000000,
    error: 100000002,
    success: 100000001
  }

  /**
   * 
   * @param {SendInAppNotifyOption} option 
   */
  this.sendInAppNotification = function (option) {
    // const iconType = IconType[option.type];
    var SendAppNotificationRequest = new DMONotifty.SendAppNotificationRequest(title = option.title,
      recipient = `/systemusers(${option.targetUserId})`,
      body = option.body,
      priority = 200000000,
      iconType = IconType[option.type],
      toastType = 200000000,
      expiry = null,
      overrideContent = null,
      actions = {
        "@odata.type": `Microsoft.Dynamics.CRM.expando`,
        "actions@odata.type": `#Collection(Microsoft.Dynamics.CRM.expando)`,
        "actions": [
          {
            "title": "View records",
            "data": {
              "@odata.type": `#Microsoft.Dynamics.CRM.expando`,
              "type": "url",
              "url": option.url,
              "navigationTarget": "newWindow"
            }
          }
        ]
      }
    );

    Xrm.WebApi.online.execute(SendAppNotificationRequest).then(function (response) {
      if (response.ok) {
        console.log("Status: %s %s", response.status, response.statusText);
        return response.json();
      }
    })
      .then(function (responseBody) {
        console.log("Response Body: %s", responseBody.NotificationId);
      })
      .catch(function (error) {
        console.log(error.message);
      });
  }
}).call(DMOTimelineCommandBarScript);