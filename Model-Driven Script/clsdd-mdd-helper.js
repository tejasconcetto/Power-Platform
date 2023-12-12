/*
Helper Library for Model-Driven Application

Helper Methods:
  1. openContactAttachmentPage(selectedId)
  2. 
*/

/*
 This method is called via Command Bar action. 
 The parameter "selectedId" reads value from
 the FirstSelectedItemId argument.
*/

function openCustomPage(selectId) {
    let recId = selectId.replaceAll("{", "").replaceAll("}", "");

    let pageInput = {
        pageType: "custom",
        name: "kvp_contactattachment_8f175", // Name of the Custom Page
        recordId: recId
    };

    let navigationOptions = {
        target: 2,
        width: { value: 50, unit: "%" },
        position: 1,
        title: "Attachment"
    };

    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function success() {
            // console.log("Success");
        },
        function error() {
            // console.log("Error");
        }
    );
}

/*
 Open the View of specified entity
*/
function openView(
    entityNameParam,
    selectedControlParam = null,
    selectedId,
    targetParam = 2,
    widthParam = 50,
    positionParam = 1,
    viewIdParam = null) {

    let pageInput = {
        pageType: "entitylist",
        entityName: entityNameParam
    };
    if (viewIdParam) {
        pageInput["viewId"] = viewIdParam;
    }
    let navigationOptions = {
        target: targetParam,
        width: { value: widthParam, unit: "%" },
        position: positionParam
    };
    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function success() {
            //
        },
        function error() {
            //
        }
    );

}

/*
 Open the alert dialog
*/
function alertDialog(
    textParam,
    titleParam = "Alert",
    confirmButtonLabelParam = "Okay",
    size = { height: 120, width: 260 }) {
    let alertStrings = {
        confirmButtonLabel: confirmButtonLabelParam,
        text: textParam,
        title: titleParam
    };
    let alertOptions = size;

    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
        function (success) {
            console.log("Alert dialog closed");
        },
        function (error) {
            console.log(error.message);
        }
    );
}

/*
 Open the confirmation dialog
*/
function confirmDialog(
    textParam,
    titleParam = "Confirmation",
    size = { height: 220, width: 460 }) {
    let confirmStrings = {
        text: textParam,
        title: titleParam
    };
    let confirmOptions = size;

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success) {
                console.log("Okay");
            } else {
                console.log("Cancel");
            }
        },
        function (error) {
            console.log(error.message);
        }
    );
}

/*
 Open the Error Dialog
*/
function errorDialog(errorCodeParam) {
    Xrm.Navigation.openErrorDialog({ errorCode: errorCodeParam }).then(
        function (success) {
            console.log(success);
        },
        function (error) {
            console.log(error);
        });
}

/*
 Open the form to create a record 
*/
function newRecord(entityName) {
    var pageInput = {
        pageType: "entityrecord",
        entityName: entityName,
        formType: 2,
    };
    var navigationOptions = {
        target: 2,
        width: { value: 50, unit: "%" },
        position: 1
    };
    Xrm.Navigation.navigateTo(pageInput, navigationOptions);
}