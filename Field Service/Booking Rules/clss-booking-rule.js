function brErrorCallback(sb) {
    console.log("ErrorCallback");
    return;
}
function brWarningCallback(sb) {
    console.log("WarningCallback");
    return;
}
function brSuccessCallback(sb) {
    console.log("SucessCallback");
    return;
}

function sampleRule(ctx) {
    console.log(ctx);
    var url = Xrm.Page.context.getClientUrl();
    var ruleResult = {
        IsValid: false,
        Message: '',
        Type: 'error'
    };

    //
    // perform some lookups or other validation logic here.
    //

    ruleResult.IsValid = false;
    ruleResult.Message = 'Some Error Message Here.';
    ruleResult.Type = 'error';

    return ruleResult;
}

var MSFSAENG;
(function (MSFSAENG) {
    MSFSAENG.ScheduleBoard = {
        url: Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.1/",
        actionName: "msfsaeng_MSFSAScheduleBoardRuleActionAskReason",
        actionInputParameters: function (ctx) {
            var inputParameters = {};
            if (ctx.isUpdate) {
                inputParameters = {
                    "originalScheduleStart": ctx.oldValues.StartTime,
                    "originalScheduleEnd": ctx.oldValues.EndTime,
                    "originalBookableResource": {
                        "@odata.type": "Microsoft.Dynamics.CRM.bookableresource",
                        "bookableresourceid": ctx.oldValues.ResourceId,
                        "name": ""
                    },
                    "originalScheduleSource": ctx.oldValues.ResourceScheduleSource,
                    "newScheduleStart": ctx.newValues.StartTime,
                    "newScheduleEnd": ctx.newValues.EndTime,
                    "newBookableResource": {
                        "@odata.type": "Microsoft.Dynamics.CRM.bookableresource",
                        "bookableresourceid": ctx.newValues.ResourceId,
                        "name": ""
                    },
                    "newScheduleSource": ctx.newValues.ResourceScheduleSource,
                    "isCreate": ctx.isCreate,
                    "isUpdate": ctx.isUpdate
                };
            }
            else {
                inputParameters = {
                    "newScheduleStart": ctx.newValues.StartTime,
                    "newScheduleEnd": ctx.newValues.EndTime,
                    "newBookableResource": {
                        "@odata.type": "Microsoft.Dynamics.CRM.bookableresource",
                        "bookableresourceid": ctx.newValues.ResourceId,
                        "name": ""
                    },
                    "newScheduleSource": ctx.newValues.ResourceScheduleSource,
                    "isCreate": ctx.isCreate,
                    "isUpdate": ctx.isUpdate
                };
            }
            return JSON.stringify(inputParameters);
        },
        ctx: null,
        ruleResult: {
            IsValid: true,
            Message: "",
            Type: ""
        },
        outputParameters: {
            isError: false,
            isWarning: false,
            errorMessage: "",
            warningMessage: ""
        },
        AskReason: function (context) {
            console.log(context);
            this.ctx = context;
            ScheduleBoardHelper.openCanvasPage(this);
            return this.ruleResult;
        },
        Validate: function (context) {
            this.ctx = context;
            ScheduleBoardHelper.callActionWebApi(this);
            return this.ruleResult;
        },
        errorCallback: brErrorCallback,
        warningCallback: brWarningCallback,
        successCallback: brSuccessCallback
    };
    var ScheduleBoardHelper = (function () {
        function ScheduleBoardHelper() {
        }
        ScheduleBoardHelper.openCanvasPage = function (sb) {
            let ov = sb.ctx.oldValues;
            let recordId = null;

            let fetchXml = `?fetchXml=<fetch>
  <entity name="bookableresourcebooking">
    <attribute name="bookableresourcebookingid" />
    <filter type="and">
      <condition attribute="resource" operator="eq" value="${ov.ResourceId}" />
      <condition attribute="msdyn_resourcerequirement" operator="eq" value="${ov.ResourceRequirementId}" />
      <condition attribute="starttime" operator="eq" value="${ov.StartTime.toISOString()}" />
      <condition attribute="endtime" operator="eq" value="${ov.EndTime.toISOString()}" />
    </filter>
  </entity>
</fetch>`;

            Xrm.WebApi.retrieveMultipleRecords("bookableresourcebooking", fetchXml).then(
                function success(result) {
                    for (var i = 0; i < result.entities.length; i++) {
                        console.log(result.entities[i]);
                        recordId = result.entities[i].bookableresourcebookingid;
                    }

                    // Open a record
                    console.log("Record Ird : ", recordId);

                    let pageInput = {
                        pageType: "custom",
                        name: "axem_statusupdatereasonpage_cc98b",
                        recordId: recordId
                    };

                    let navigationOptions = {
                        target: 2,
                        width: { value: 300, unit: "px" },
                        height: { value: 280, unit: "px" },
                        position: 1,
                        title: "Reason"
                    };

                    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
                        function success() {
                            console.log("Success");
                        },
                        function error() {
                            console.log("Error");
                        }
                    );

                },
                function (error) {
                    console.log(error.message);
                    // handle error conditions
                });

        }
        ScheduleBoardHelper.callAskReason = function (sb) {
            let reasonText = prompt("Eplain why you want to update the status");
            console.log(reasonText, reasonText == "");
            if (!reasonText) {
                sb.ruleResult.IsValid = false;
                sb.ruleResult.Message = 'You must provide a reason behind updating the status';
                sb.ruleResult.Type = 'error';
                if (sb.successCallback) {
                    sb.successCallback(sb);
                }
                return;
            } else {
                sb.ruleResult.IsValid = true;
                sb.ruleResult.Message = '';
                sb.ruleResult.Type = '';
                if (sb.successCallback) {
                    sb.successCallback(sb);
                }
                return;
            }
        }
        ScheduleBoardHelper.callActionWebApi = function (sb) {
            var oDataEndpoint = sb.url + sb.actionName;
            var req = new XMLHttpRequest();
            req.open("POST", oDataEndpoint, false);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.onreadystatechange = function () {
                if (req.readyState == 4) {
                    req.onreadystatechange = null;
                    if (req.status == 200) {
                        sb.outputParameters = JSON.parse(req.response);
                        if (sb.outputParameters.isError) {
                            sb.ruleResult.IsValid = false;
                            sb.ruleResult.Message = sb.outputParameters.errorMessage;
                            sb.ruleResult.Type = 'error';
                            if (sb.errorCallback)
                                sb.errorCallback(sb);
                            return;
                        }
                        else if (sb.outputParameters.isWarning) {
                            sb.ruleResult.IsValid = false;
                            sb.ruleResult.Message = sb.outputParameters.warningMessage;
                            sb.ruleResult.Type = 'warning';
                            if (sb.warningCallback)
                                sb.warningCallback(sb);
                            return;
                        }
                        else {
                            sb.ruleResult.IsValid = true;
                            sb.ruleResult.Message = '';
                            sb.ruleResult.Type = '';
                            if (sb.successCallback)
                                sb.successCallback(sb);
                            return;
                        }
                    }
                    else {
                        alert('Error calling Rule Action. Response = ' + req.response + ', Status = ' + req.statusText);
                    }
                }
            };
            req.send(sb.actionInputParameters(sb.ctx));
        };
        return ScheduleBoardHelper;
    }());
})(MSFSAENG || (MSFSAENG = {}));