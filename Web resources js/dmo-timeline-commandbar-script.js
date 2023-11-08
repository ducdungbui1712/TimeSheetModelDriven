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

    var context;

    if (typeof GetGlobalContext === 'function') {
      context = GetGlobalContext();
    } else {
      context = Xrm.Page.context;
    }
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
  const onConfirmSubmit = async (records) => {
    const loggedinUserId = Xrm.Utility.getGlobalContext().userSettings.userId.replace('{', '').replace('}', '');
    try {
      const result = await Xrm.WebApi.online.retrieveRecord(
        'systemuser',
        loggedinUserId,
        '?$select=systemuserid,domainname,parentsystemuserid,_parentsystemuserid_value'
      );
      const parentUserId = result['_parentsystemuserid_value'];
      if (!parentUserId) {
        Xrm.Utility.alertDialog('You do not have a Reporting Manager to submit your timesheet, please contact the DMO admin for assistance.');
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
      setTimeout(() => {
        Xrm.Utility.refreshParentGrid({ entityType: records[0].TypeName, id: records[0].Id });
      }, 100);
    } catch (err) {
      Xrm.Utility.alertDialog(error.message);
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
      // setTimeout(() => {
      //   Xrm.Utility.refreshParentGrid({ entityType: records[0].TypeName, id: records[0].Id });
      // }, 100);
    } catch (err) {
      Xrm.Utility.alertDialog(error.message);
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
        } finally {
          gridContext?.refresh(false);
          // gridContext?.refreshRibbon(false);
          Xrm.Utility.closeProgressIndicator();
        }
      } else console.log('Dialog closed using Cancel button or X.');
    });
  };

  /**
   *
   * @param {SelectedControlAllItemRef[]} records
   */
  this.onClickApproveCommand = async function (records) {
    console.log('data', records);
    const confirmStrings = { text: 'Are you sure to approve this records', title: 'Confirmation approve timesheets' };
    const confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(async function (success) {
      if (success.confirmed) {
        Xrm.Utility.showProgressIndicator('Please waiting a moment...');
        try {
          await onApproveTimeSheet(records);
        } catch (error) {
        } finally {
          Xrm.Utility.closeProgressIndicator();
        }
      } else console.log('Dialog closed using Cancel button or X.');
    });
  };
}).call(DMOTimelineCommandBarScript);
