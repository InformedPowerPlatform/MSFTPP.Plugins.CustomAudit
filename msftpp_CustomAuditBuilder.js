/*
Work to be done
1. Add a delete plugin step function with confirmation dialog

*/

if (typeof MSFTPP === "undefined") {
    MSFTPP = { __namespace: true };
}
if (typeof MSFTPP.CAB === "undefined") {
    MSFTPP.CAB = { __namespace: true };
}

Xrm = window.Xrm || { __namespace: true };
// MSFTPP = window.Xrm || { __namespace: true };
// MSFTPP.CAB = MSFTPP.CAB || { __namespace: true };

$(function () {
    console.log("PreLoad Function!");
});

var uri = "";
var InstanceSubName = "";
var queryStringVals = new Array();
var recordId = "";
var entitySDKName = "";
var orgName = "";
var boolIsDev = true;
var selectedFields = [];
var allTableFields = [];
var allTables = [];
var boolAddNew = false;
var boolUpdateExisting = false;
var strPluginTypeId = "";
var strEntityNameGlobal = "";
var strSDKMessageFilterId = "0";
var strSDKMessageFilterName = "";
var strSDKMessageId = "";
var strSDKMessageName = "";
var strExistingPluginId = "";
var strExistingPluginName = "";

document.onreadystatechange = function () {
    if (document.readyState === "complete") {
        // alert("document.readyState === complete");
        MSFTPP.CAB.showConsoleLog("readyState = complete!");
        MSFTPP.CAB.pageOnLoad();
    }
};

MSFTPP.CAB.pageOnLoad = function () {
    MSFTPP.CAB.showConsoleLog("pageOnLoad... ");
    $("#divLoading").show();
    $("#ddlTables").empty();
    $("#ddlPluginSteps").empty();
    $("#ddlMessages").empty();

    $("#divForm").hide();
    $("#itemColumnTable tr").remove();
    $("#divSelectorRow").hide();
    $("#divColumnTable").hide();
    $("#divAddNewSDKMessage").hide();

    $("#divPnlError").hide();

    $("#divSubmitting").hide();
    $("#divOutput").hide();
    $("#divAddNewTable").hide();
    $("#divEditSteps").hide();

    selectedFields = [];
    allTableFields = [];
    allTables = [];

    // Set the strPluginTypeId
    MSFTPP.CAB.getPluginTypeId();

    $("#divLoading").hide();
    $("#divForm").show();
    $("#divSelectorRow").show();


};

MSFTPP.CAB.getPluginTypeId = function () {
    // https://org1134971a.crm9.dynamics.com//api/data/v9.1/plugintypes?$select=plugintypeid,typename&$filter=startswith(typename,'MSFTPP')
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/plugintypes?$select=plugintypeid,typename&$filter=startswith(typename,'MSFTPP.Plugin')",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=10");
        },
        async: true,
        success: function (data, textStatus, xhr) {
            var results = data;
            for (var i = 0; i < results.value.length; i++) {
                // var plugintypeid = results.value[i]["plugintypeid"];
                // var typename = results.value[i]["typename"];
                strPluginTypeId = results.value[i]["plugintypeid"];
                MSFTPP.CAB.showConsoleLog(" MSFTPP.CAB.getPluginTypeId strPluginTypeId = " + strPluginTypeId);
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            Xrm.Utility.alertDialog(textStatus + " " + errorThrown);
        }
    });
};

MSFTPP.CAB.onClickAddNew = function () {
    $("#divAddNewTable").show();
    $("#tableRow").hide();
    $("#divEditSteps").hide();
    boolAddNew = true;
    boolUpdateExisting = false;
    selectedFields = [];
    allTableFields = [];
    allTables = [];
    MSFTPP.CAB.getTables();
}; // end onClickAddNew

MSFTPP.CAB.onClickEditExisting = function () {
    $("#divAddNewTable").hide();
    $("#tableRow").hide();

    $("#divEditSteps").show();
    boolAddNew = false;
    boolUpdateExisting = true;
    selectedFields = [];
    allTableFields = [];
    allTables = [];
    MSFTPP.CAB.getPluginSteps();
}; // end onClickEditExisting

MSFTPP.CAB.getTables = function () {
    console.log("getEntities... ");
    var localURL = Xrm.Page.context.getClientUrl();
    // localURL += "/api/data/v9.1/EntityDefinitions?$select=LogicalName,SchemaName,MetadataId,HasActivities,DisplayCollectionName&$filter=HasActivities%20eq%20true%20and%20IsLogicalEntity%20eq%20false";
    localURL += "/api/data/v9.1/EntityDefinitions";
    localURL += "?$select=LogicalName,SchemaName,MetadataId,HasActivities,DisplayCollectionName";
    localURL += "&$filter=HasActivities%20eq%20true";

    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: localURL,
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            // XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=10");
            XMLHttpRequest.setRequestHeader(
                "Prefer",
                'odata.include-annotations="*"'
            );
        },
        async: true,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results.value.length === 0) {
                $("#ddlTables").empty();
                let opt = document.createElement("option");
                opt.value = "0";
                opt.innerText = "No entities returned...";
                $("#ddlTables").append(opt);
            } else {
                console.log("Success! ");
                $("#ddlTables").empty();
                let opt = document.createElement("option");
                opt.value = "0";
                opt.innerText = "Select Table...";
                $("#ddlTables").append(opt);

                allTables = [];
                for (var i = 0; i < results.value.length; i++) {
                    var entityId = results.value[i]["MetadataId"];
                    var physicalName = results.value[i]["LogicalName"];
                    var friendlyName = results.value[i]["DisplayCollectionName"]["UserLocalizedLabel"]["Label"];

                    let tableData = {
                        entityId: entityId,
                        physicalName: physicalName,
                        friendlyName: friendlyName,
                    };
                    allTables.push(tableData);

                }


                // Order by doesn't seem to work, so we have to re-sort the list
                allTables.sort(function (x, y) {
                    let a = x.friendlyName.toUpperCase(),
                        b = y.friendlyName.toUpperCase();
                    return a == b ? 0 : a > b ? 1 : -1;
                });

                for (var i = 0; i < allTables.length; i++) {
                    var entityId2 = allTables[i]["entityId"];
                    var physicalName2 = allTables[i]["physicalName"];
                    var friendlyName2 = allTables[i]["friendlyName"];
                    // MSFTPP.CAB.showConsoleLog("LogicalName :: " + LogicalName);
                    let option = document.createElement("option");
                    // option.value = entityid;
                    option.value = physicalName2;
                    option.innerText = friendlyName2 + " [" + physicalName2 + "]";
                    // MSFTPP.CAB.showConsoleLog("friendlyName :: " + friendlyName2 + " [" + physicalName2 + "]");
                    $("#ddlTables").append(option);
                }

            }
        },
        error: function (xhr, textStatus, errorThrown) {
            MSFTPP.CAB.handleError(xhr.responseText);
        },
    });
}; // end getEntities

MSFTPP.CAB.getFields = function (theSelectedOption) {

    var strEntityLogicalName = '0';
    var strEntityFriendlyName = '';
    strEntityLogicalName = theSelectedOption.selectedOptions[0].value.toString();
    strEntityFriendlyName = theSelectedOption.selectedOptions[0].text.toLowerCase();
    strEntityNameGlobal = theSelectedOption.selectedOptions[0].text;

    MSFTPP.CAB.getSDKMessageFiltersForEntity(strEntityLogicalName);
    MSFTPP.CAB.getEntityColumns(strEntityLogicalName, strEntityFriendlyName);
};

MSFTPP.CAB.getSDKMessageFiltersForEntity = function (entityName) {
    $("#ddlMessages").empty();
    // TODO: Need some code to get the available sdkmessagefilters for this entity
    // https://orgff725b8b.crm.dynamics.com/api/data/v9.1/sdkmessagefilters?$select=primaryobjecttypecode,_sdkmessageid_value&$expand=sdkmessageid($select=categoryname,name)&$filter=primaryobjecttypecode%20eq%20%27account%27
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url:
            Xrm.Page.context.getClientUrl() +
            "/api/data/v9.1/sdkmessagefilters?$select=primaryobjecttypecode,_sdkmessageid_value&$expand=sdkmessageid($select=categoryname,name)&$filter=primaryobjecttypecode eq %27" + entityName + "%27",
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            // XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=10");
            XMLHttpRequest.setRequestHeader(
                "Prefer",
                'odata.include-annotations="*"'
            );
        },
        async: true,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results.value.length === 0) {
                $("#ddlMessages").empty();
                let opt = document.createElement("option");
                opt.value = "0";
                opt.innerText = "No messages returned...";
                $("#ddlMessages").append(opt);
            } else {

                $("#ddlMessages").empty();
                let opt = document.createElement("option");
                opt.value = "0";
                opt.innerText = "Select Message...";
                $("#ddlMessages").append(opt);

                for (var i = 0; i < results.value.length; i++) {
                    var sdkmessagefilterid = results.value[i]["sdkmessagefilterid"];

                    var sdkmessageid = results.value[i]["_sdkmessageid_value"];
                    var messagename = results.value[i]["sdkmessageid"]["name"];

                    // MSFTPP.CAB.showConsoleLog("messagename :: " + messagename);
                    if (messagename === 'Create'
                        || messagename === 'Delete'
                        || messagename === 'Update'
                    ) {
                        let option = document.createElement("option");
                        option.value = sdkmessagefilterid + "~" + sdkmessageid;
                        option.innerText = messagename;
                        $("#ddlMessages").append(option);
                    }
                }
                $("#divAddNewSDKMessage").show();
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            MSFTPP.CAB.handleError(textStatus + " " + errorThrown);
        },
    });


}; // getSDKMessageFiltersForEntity

// MSFTPP.CAB.getEntityColumns = function (entityId, entityName) {
MSFTPP.CAB.getEntityColumns = function (entityLogicalName, entityFriendlyName) {

    MSFTPP.CAB.showConsoleLog("entityLogicalName :: " + entityLogicalName);
    MSFTPP.CAB.showConsoleLog("entityFriendlyName :: " + entityFriendlyName);

    // No Table selected
    if (entityLogicalName === "0") {
        $("#itemColumnTable tr").remove();
        $("#divColumnTable").hide();
        return;
    }

    MSFTPP.CAB.showConsoleLog("entityName :: " + entityFriendlyName);
    $("#itemColumnTable tr").remove();
    var headerRow = $("<tr />");
    $("#itemColumnTable").append(headerRow);
    headerRow.append(
        $(
            "<th style='text-align: center;'><input type='checkbox' class='form-check-input' id='checkAllColumns' onclick='MSFTPP.CAB.onClickField(this)' value='0'></th>"
        )
    );
    headerRow.append($("<th>Name</th>"));
    headerRow.append($("<th>Type</th>"));
    var fullURL =
        Xrm.Page.context.getClientUrl() +
        "/api/data/v9.1/EntityDefinitions(LogicalName='" +
        entityLogicalName +
        "')/Attributes?$select=LogicalName,DisplayName,AttributeType&$orderby=LogicalName";
    MSFTPP.CAB.showConsoleLog(fullURL);
    // api/data/v9.1/EntityDefinitions(LogicalName='account')/Attributes?$select=DisplayName,LogicalName&$orderby=LogicalName
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        // url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/EntityDefinitions(LogicalName='" + entityName + "')/Attributes?$select=LogicalName,DisplayName,AttributeType&$orderby=LogicalName",
        url: fullURL,
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            // XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=10");
            XMLHttpRequest.setRequestHeader(
                "Prefer",
                'odata.include-annotations="*"'
            );
        },
        // async: true,
        async: false,
        success: function (data, textStatus, xhr) {
            var results = data;
            if (results.value.length === 0) {
                // TODO:
            } else {
                MSFTPP.CAB.showConsoleLog("getFields - Success! ");
                allTableFields = [];
                for (var i = 0; i < results.value.length; i++) {
                    var LogicalName2 = results.value[i]["LogicalName"];
                    var MetadataId2 = results.value[i]["MetadataId"];
                    var AttributeType2 = results.value[i]["AttributeType"];
                    var LocalLabel2 = "";
                    if(results.value[i]['DisplayName']['UserLocalizedLabel'] !== null) {
                        LocalLabel2 = results.value[i]['DisplayName']['UserLocalizedLabel']['Label'];
                    }
                    if (AttributeType2.toLowerCase() !== "virtual"
                        && LogicalName2.toLowerCase().includes("createdby") !== true
                        && LogicalName2.toLowerCase().includes("modifiedby") !== true
                        && LogicalName2.toLowerCase().includes("yomi") !== true
                        && LogicalName2.toLowerCase().includes("createdon") !== true
                        && LogicalName2.toLowerCase().includes("modifiedon") !== true
                        ) {
                        let tableField = {
                            LogicalName: LogicalName2,
                            MetadataId: MetadataId2,
                            AttributeType: AttributeType2,
                            LocalLabel: LocalLabel2
                        };
                        allTableFields.push(tableField);
                    }
                }

                // Order by doesn't seem to work, so we have to re-sort the list
                
                
                /*
                allTableFields.sort(function (x, y) {
                    let a = x.LogicalName.toUpperCase(),
                        b = y.LogicalName.toUpperCase();
                    return a == b ? 0 : a > b ? 1 : -1;
                });
                */
                // Sort by Local Friendly Label instead
                allTableFields.sort(function (x, y) {
                    let a = x.LocalLabel.toUpperCase(),
                        b = y.LocalLabel.toUpperCase();
                    return a == b ? 0 : a > b ? 1 : -1;
                });

                for (var i = 0; i < allTableFields.length; i++) {
                    var LogicalName = allTableFields[i]["LogicalName"];
                    var MetadataId = allTableFields[i]["MetadataId"];
                    var AttributeType = allTableFields[i]["AttributeType"];
                    var LocalLabel = allTableFields[i]["LocalLabel"];
                    // MSFTPP.CAB.showConsoleLog("LogicalName :: " + LogicalName);
                    $("#divOutputPanelBody").html(LogicalName + ",");
                    var row = $("<tr />");
                    $("#itemColumnTable").append(row);
                    row.append(
                        $(
                            "<td style='text-align: center;'><input type='checkbox' class='form-check-input' onclick='MSFTPP.CAB.onClickField(this)' id='" +
                            MetadataId.toString() +
                            "' name='" +
                            LogicalName +
                            "' value='" +
                            MetadataId.toString() +
                            "' ></td>"
                        )
                    );
                    row.append($("<td>" + LocalLabel + " [" + LogicalName + "]</td>"));
                    row.append($("<td>" + AttributeType + "</td>"));
                }

                // $("#divOutputPanelBody").html("");
                // $("#divColumnTable").show();
                // $("#divOutput").show();
                // $("#tableRow").show();
            }
        },
        error: function (xhr, textStatus, errorThrown) {
            MSFTPP.CAB.handleError(textStatus + " " + errorThrown);
        },
    });
}; // end getFields

MSFTPP.CAB.getqueryStringVals_ssb = function () {
    //Get the any query string parameters and load them
    //into the vals array
    MSFTPP.CAB.showConsoleLog("getqueryStringVals... ");

    if (location.search !== "") {
        if (
            location.hostname.indexOf("dev") > 0 ||
            location.hostname.indexOf("test") > 0
        ) {
            boolIsDev = true;
        }

        if (boolIsDev) {
            // $("#divPnlDebug").show();
            // $("#lblDebug").text(location.search);
        }

        queryStringVals = location.search.substr(1).split("&");
        for (var i in queryStringVals) {
            queryStringVals[i] = queryStringVals[i].replace(/\+/g, " ").split("=");
            var strVal = decodeURIComponent(queryStringVals[i][1].toString());
            queryStringVals[i][1] = strVal;
        }

        //look for the parameter named 'data'
        var found = false;
        for (var i in queryStringVals) {
            switch (queryStringVals[i][0].toLowerCase()) {
                case "id":
                    recordId = queryStringVals[i][1].toLowerCase();
                    MSFTPP.CAB.showConsoleLog("recordId :: " + recordId);
                    break;
                case "orgname":
                    orgName = queryStringVals[i][1].toLowerCase();
                    MSFTPP.CAB.showConsoleLog("orgName :: " + orgName);
                    break;
                case "typename":
                    entitySDKName = queryStringVals[i][1].toLowerCase();
                    MSFTPP.CAB.showConsoleLog("entitySDKName :: " + entitySDKName);
                    break;
            }
        }
    } else {
        MSFTPP.CAB.handleError("No data parameter was passed to this page");
    }
}; // End Function getqueryStringVals

MSFTPP.CAB.pad = function (n, width, z) {
    z = z || "0";
    n = n + "";
    return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
};

MSFTPP.CAB.onClickField = function (theCheckbox) {

    MSFTPP.CAB.showConsoleLog(
        "theCheckbox.id.toString() - " + theCheckbox.id.toString()
    );

    // Check if this is the main CheckALL box
    if (theCheckbox.value.toString() === "0" && theCheckbox.checked) {
        MSFTPP.CAB.checkAllFields(true);
        return;
    }

    if (theCheckbox.value.toString() === "0" && !theCheckbox.checked) {
        $("#divPnlAttributes").hide();
        MSFTPP.CAB.checkAllFields(false);
        return;
    }

    if (theCheckbox.checked) {
        MSFTPP.CAB.showConsoleLog(
            "checked!! = " +
            theCheckbox.value.toString() +
            " :: Name = " +
            theCheckbox.name
        );
        let site = {
            Id: theCheckbox.value.toString(),
            Name: theCheckbox.name,
            MetadataId: theCheckbox.id.toString(),
        };
        selectedFields.push(site);
    } else {
        MSFTPP.CAB.showConsoleLog(
            "NOT checked!! = " +
            theCheckbox.value.toString() +
            " :: Name = " +
            theCheckbox.name
        );
        var theOneToDelete;
        for (var i = 0; i < selectedFields.length; i++) {
            if (selectedFields[i].Id === theCheckbox.value.toString()) {
                MSFTPP.CAB.showConsoleLog(
                    "Removing -> " +
                    i.toString() +
                    " - " +
                    selectedFields[i].Id +
                    " :: " +
                    selectedFields[i].Name
                );
                theOneToDelete = i;
                break;
            } else {
                MSFTPP.CAB.showConsoleLog(
                    "Skipping -> " +
                    i.toString() +
                    " - " +
                    selectedFields[i].Id +
                    " :: " +
                    selectedFields[i].Name
                );
            }
        }
        selectedFields.splice(theOneToDelete, 1);
    }
    MSFTPP.CAB.updateOutputPanel();

}; // end onClickSite

MSFTPP.CAB.updateOutputPanel = function () {
    MSFTPP.CAB.showConsoleLog("selectedFields.length -> " + selectedFields.length);
    $("#divOutputPanelBody").html("");
    for (var i = 0; i < selectedFields.length; i++) {
        var LogicalName = selectedFields[i]["Name"];
        // MSFTPP.CAB.showConsoleLog("LogicalName :: " + LogicalName);
        var xx = $("#divOutputPanelBody").html();
        if (i === 0) {
            xx = LogicalName;
        } else {
            xx = xx + "," + LogicalName;
        }
        $("#divOutputPanelBody").html(xx);
    }
};

MSFTPP.CAB.hideSuccessPanel = function () {
    $("#divPnlSuccess").hide();
};



MSFTPP.CAB.showConsoleLog = function (varString) {
    // This should only run in DEV or TEST.
    if (boolIsDev === true) {
        var t = new Date();
        var timeStamp =
            t.getHours() +
            ":" +
            MSFTPP.CAB.pad(t.getMinutes(), 2) +
            ":" +
            MSFTPP.CAB.pad(t.getSeconds(), 2) +
            "." +
            t.getMilliseconds();

        console.log(timeStamp + " ~ " + varString);
    }
};

MSFTPP.CAB.getPluginSteps = function () {
    MSFTPP.CAB.showConsoleLog("getPluginSteps... ");
    var filteredURL = Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingsteps?$filter=contains(name, 'MSFTPP.')";
    MSFTPP.CAB.showConsoleLog("filteredURL: " + filteredURL);

    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: filteredURL,
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            // XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=10");
            XMLHttpRequest.setRequestHeader(
                "Prefer",
                'odata.include-annotations="*"'
            );
        },
        async: true,
        success: function (data, textStatus, xhr) {
            var results = data;

            if (results.value.length === 0) {
                $("#ddlPluginSteps").empty();
                let opt = document.createElement("option");
                opt.value = "0";
                opt.innerText = "No Steps returned...";
                $("#ddlPluginSteps").append(opt);
            }
            else {
                console.log("Success! ");
                $("#ddlPluginSteps").empty();
                let opt = document.createElement("option");
                opt.value = "0";
                opt.innerText = "Select Plugin Step ...";
                $("#ddlPluginSteps").append(opt);

                for (var i = 0; i < results.value.length; i++) {
                    var description = results.value[i]["name"];
                    var sdkmessageprocessingstepid = results.value[i]["sdkmessageprocessingstepid"];
                    let option = document.createElement("option");
                    option.value = sdkmessageprocessingstepid;
                    option.innerText = description;
                    $("#ddlPluginSteps").append(option);
                }
            }
        },
        error: function (xhr, textStatus, errorThrown) {

            MSFTPP.CAB.handleError(textStatus + " " + errorThrown);
        },
    });
}; // end getEntities

MSFTPP.CAB.getPluginFields = function (theSelectedOption) {
    MSFTPP.CAB.showConsoleLog("getPluginFields... ");

    // &$expand=sdkmessagefilterid($select=name,primaryobjecttypecode,sdkmessagefilterid)

    var strId = theSelectedOption.selectedOptions[0].value.toString();
    var strName = theSelectedOption.selectedOptions[0].text.toLowerCase();
    strExistingPluginId = strId;
    strExistingPluginName = strName;

    var filteredURL = Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingsteps(" + strId + ")?$expand=sdkmessagefilterid($select=name,primaryobjecttypecode)";
    // Get data from this sdkmessageprocessingstep
    $.ajax({
        type: "GET",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: filteredURL,
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
            XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        },
        async: true,
        success: function (data, textStatus, xhr) {
            var result = data;
            var sdkmessageprocessingstepid = result["sdkmessageprocessingstepid"];
            var filteringattributes = result["filteringattributes"];
            if (result.hasOwnProperty("sdkmessagefilterid")) {
                var sdkmessagefilterid_name = result["sdkmessagefilterid"]["name"];
                var sdkmessagefilterid_primaryobjecttypecode = result["sdkmessagefilterid"]["primaryobjecttypecode"];

            }
            filteredURL = Xrm.Page.context.getClientUrl() + "/api/data/v9.1/entities?$select=entityid,logicalcollectionname,logicalname,physicalname&$filter=logicalname eq '" + sdkmessagefilterid_primaryobjecttypecode + "'";

            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                datatype: "json",
                url: filteredURL,
                beforeSend: function (XMLHttpRequest) {
                    XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
                    XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
                    XMLHttpRequest.setRequestHeader("Accept", "application/json");
                    XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                },
                async: true,
                success: function (data, textStatus, xhr) {
                    var results = data;

                    for (var i = 0; i < results.value.length; i++) {
                        var entityId = results.value[i]["entityid"];
                        var logicalcollectionname = results.value[i]["logicalcollectionname"];
                        var logicalname = results.value[i]["logicalname"];
                        var physicalname = results.value[i]["physicalname"];
                        break;
                    }
                    selectedFields = [];
                    MSFTPP.CAB.showConsoleLog("entityId :: " + entityId);
                    MSFTPP.CAB.getEntityColumns(logicalname, physicalname);
                    MSFTPP.CAB.showConsoleLog("filteringattributes :: " + filteringattributes);
                    MSFTPP.CAB.checkSelectedFields(filteringattributes);

                    $("#divColumnTable").show();
                    $("#divOutput").show();
                    $("#tableRow").show();
                },
                error: function (xhr, textStatus, errorThrown) {
                    MSFTPP.CAB.handleError(textStatus + " " + errorThrown);
                }
            });

        },
        error: function (xhr, textStatus, errorThrown) {
            MSFTPP.CAB.handleError(textStatus + " " + errorThrown);

        }
    });

}; // end getPluginFields

MSFTPP.CAB.checkSelectedFields = function (fieldList) {

    // var strArray = fieldList.split(',');
    // for (var i = 0; i < strArray.length; i++) {
    //     MSFTPP.CAB.showConsoleLog(i.toString() + " :: " + strArray[i]);
    // }
    fieldList += ',,,';
    // MSFTPP.CAB.showConsoleLog('checkSelectedFields -> fieldList :: ' + fieldList);

    $("input[type=checkbox]").each(function () {
        var thisItem = $(this);

        var cbValue = this.value.toString();
        var cbName = this.name + ',';
        var cbId = this.id.toString();
        // MSFTPP.CAB.showConsoleLog('checkSelectedFields -- fieldList.indexOf(cbName) = ' + fieldList.indexOf(cbName) + ' :: ' + cbName + ' :: ' + cbId);

        if (fieldList.indexOf(cbName) >= 0 && cbName.length > 1) {
            $(this).prop("checked", true);
            MSFTPP.CAB.onClickField(this);
        }

        // ictr = ictr + 1;
    });

};

MSFTPP.CAB.handleError = function (errorText) {
    $("#divPnlError").show();
    $("#lblError").text(errorText);
    MSFTPP.CAB.showConsoleLog(errorText);
};

MSFTPP.CAB.messageOnChange = function (theSelectedMessage) {

    // strSDKMessageFilterName =  $("#ddlMessages :selected").text().toLowerCase();
    // strSDKMessageName = $("#ddlMessages :selected").text();
    strSDKMessageName = theSelectedMessage.selectedOptions[0].text;
    strSDKMessageFilterName = strSDKMessageName.toLowerCase();

    //strIDs = $("#ddlMessages").val().toString().split("~");
    strIDs = theSelectedMessage.selectedOptions[0].value.toString().split("~");
    strSDKMessageFilterId = strIDs[0];
    strSDKMessageId = strIDs[1];

    MSFTPP.CAB.showConsoleLog("strSDKMessageFilterId :: " + strSDKMessageFilterId);
    MSFTPP.CAB.showConsoleLog("strSDKMessageId :: " + strSDKMessageId);
    MSFTPP.CAB.showConsoleLog("strSDKMessageFilterName :: " + strSDKMessageFilterName);
    $("#divOutputPanelBody").html("");
    $("#divColumnTable").show();
    $("#divOutput").show();
    $("#tableRow").show();


}; // end messageOnChange

MSFTPP.CAB.addNewPluginStep = function () {

    var validationPassed = true;
    var txtValidationError = "";

    if (strSDKMessageFilterId === "0" || strSDKMessageFilterId === null) {
        txtValidationError += "<p>Need to choose a Message</p>";
        $("#divComments").addClass("has-error");
        validationPassed = false;
    }

    if (!validationPassed) {
        $("#divPnlError").show();
        $("#divPnlError").html(txtValidationError);
        return;
    }

    if (validationPassed) {
        $("#divPnlError").hide();
        $("#divSubmitting").show();

        var fieldList = '';
        $.each(selectedFields, function () {
            var tempName = this.Name.toLowerCase();
            var tempId = this.Id;
            // MSFTPP.CAB.showConsoleLog("tempName: " + tempName);
            fieldList += tempName + ',';
        });
        // Trim off the last comma
        fieldList = fieldList.substr(0, fieldList.lastIndexOf(','));
        MSFTPP.CAB.showConsoleLog("fieldList :: " + fieldList);
        MSFTPP.CAB.showConsoleLog("strPluginTypeId = " + strPluginTypeId);
        const d = new Date();
        let strISODateTimeStamp = d.toISOString();

        // Create Step
        var entity = {};
        entity.stage = 40;
        entity.mode = 1;
        entity.asyncautodelete = true;
        entity["sdkmessagefilterid@odata.bind"] = "/sdkmessagefilters(" + strSDKMessageFilterId + ")";  // This is the On Create/Update/Delete option from sdkmessagefilter(s)
        entity.filteringattributes = fieldList; // List of fields to trigger on
        entity["eventhandler_plugintype@odata.bind"] = "/plugintypes(" + strPluginTypeId + ")";  // This is the parent Plugin
        entity["sdkmessageid@odata.bind"] = "/sdkmessages(" + strSDKMessageId + ")";
        entity.rank = 10;
        entity.name = "MSFTPP.Plugin.CustomAudit - " + strSDKMessageName + " of " + strEntityNameGlobal + " :: " + strISODateTimeStamp;
        entity.description = "MSFTPP.Plugin.CustomAudit - " + strSDKMessageName + " of " + strEntityNameGlobal + " :: " + strISODateTimeStamp;

        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            datatype: "json",
            url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingsteps",
            data: JSON.stringify(entity),
            beforeSend: function (XMLHttpRequest) {
                XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
                XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
                XMLHttpRequest.setRequestHeader("Accept", "application/json");
            },
            async: true,
            success: function (data, textStatus, xhr) {
                var uri = xhr.getResponseHeader("OData-EntityId");
                var regExp = /\(([^)]+)\)/;
                var matches = regExp.exec(uri);
                var newEntityId = matches[1];
                MSFTPP.CAB.showConsoleLog("Step Created! New ID = " + newEntityId);
                MSFTPP.CAB.createImage(newEntityId, fieldList, strSDKMessageName);
                $("#divForm").hide();
                $("#divColumnTable").hide();

            },
            error: function (xhr, textStatus, errorThrown) {
                MSFTPP.CAB.handleError(xhr.responseText);
            }
        });
    }
}; // end addNewPluginStep

MSFTPP.CAB.onClickSubmit = function () {


    if (boolAddNew) { MSFTPP.CAB.addNewPluginStep(); }
    if (boolUpdateExisting) { MSFTPP.CAB.updatePluginStep(); }


}; // end onClickSubmit

MSFTPP.CAB.updatePluginStep = function () {
    MSFTPP.CAB.showConsoleLog("updatePluginStep!");

    var fieldList = '';
    $.each(selectedFields, function () {
        var tempName = this.Name.toLowerCase();
        var tempId = this.Id;
        // MSFTPP.CAB.showConsoleLog("tempName: " + tempName);
        fieldList += tempName + ',';
    });
    // Trim off the last comma
    fieldList = fieldList.substr(0, fieldList.lastIndexOf(','));
    MSFTPP.CAB.showConsoleLog("updatePluginStep fieldList :: " + fieldList);
    // #ddlPluginSteps

    MSFTPP.CAB.showConsoleLog("strExistingPluginId :: " + strExistingPluginId);
    MSFTPP.CAB.showConsoleLog("strExistingPluginName :: " + strExistingPluginName);

    // Update the Step

    var entity = {};
    entity.filteringattributes = fieldList;

    $.ajax({
        type: "PATCH",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingsteps(" + strExistingPluginId + ")",
        data: JSON.stringify(entity),
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
        },
        async: true,
        success: function (data, textStatus, xhr) {
            MSFTPP.CAB.showConsoleLog("Success! - sdkmessageprocessingsteps updated with :: " + fieldList);
            // Get the image for this step
            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                datatype: "json",
                url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingstepimages?$filter=_sdkmessageprocessingstepid_value eq " + strExistingPluginId,
                beforeSend: function (XMLHttpRequest) {
                    XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
                    XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
                    XMLHttpRequest.setRequestHeader("Accept", "application/json");
                    XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
                },
                async: true,
                success: function (data, textStatus, xhr) {
                    var results = data;
                    for (var i = 0; i < results.value.length; i++) {
                        var sdkmessageprocessingstepimageid = results.value[i]["sdkmessageprocessingstepimageid"];
                    }
                    

                    // Delete the existing image
                    $.ajax({
                        type: "DELETE",
                        contentType: "application/json; charset=utf-8",
                        datatype: "json",
                        url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingstepimages(" + sdkmessageprocessingstepimageid + ")",
                        beforeSend: function(XMLHttpRequest) {
                            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
                            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
                            XMLHttpRequest.setRequestHeader("Accept", "application/json");
                        },
                        async: true,
                        success: function(data, textStatus, xhr) {
                            $.ajax({
                                type: "GET",
                                contentType: "application/json; charset=utf-8",
                                datatype: "json",
                                url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingsteps(" + strExistingPluginId + ")?$expand=sdkmessageid($select=name)",
                                beforeSend: function(XMLHttpRequest) {
                                    XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
                                    XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
                                    XMLHttpRequest.setRequestHeader("Accept", "application/json");
                                    XMLHttpRequest.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                                },
                                async: true,
                                success: function(data, textStatus, xhr) {
                                    var result = data;
                                    var sdkmessageprocessingstepid = result["sdkmessageprocessingstepid"];
                                    if (result.hasOwnProperty("sdkmessageid")) {
                                        var sdkmessageid_name = result["sdkmessageid"]["name"];
                                    }
                                    MSFTPP.CAB.createImage(strExistingPluginId, fieldList, sdkmessageid_name);

                                },
                                error: function(xhr, textStatus, errorThrown) {
                                    MSFTPP.CAB.handleError(xhr.responseText);
                                }
                            });
                            
                        },
                        error: function(xhr, textStatus, errorThrown) {
                             MSFTPP.CAB.handleError(xhr.responseText);
                        }
                    });


                },
                error: function (xhr, textStatus, errorThrown) {
                    MSFTPP.CAB.handleError(xhr.responseText);
                }
            });
        },
        error: function (xhr, textStatus, errorThrown) {
            MSFTPP.CAB.handleError(xhr.responseText);
        }
    });


    /*
    
    */


}; // end onClickSubmit

MSFTPP.CAB.createImage = function (newStepId, fieldList, messageType) {

    var entity2 = {};

    // imagetype PostImage = 1,Both = 2
    if (messageType.toLowerCase() === "create") {
        entity2.imagetype = 1;
        entity2.messagepropertyname = "Id";
    }
    else {
        entity2.imagetype = 2;
        entity2.messagepropertyname = "Target";
    }
    // Create Step Image

    entity2.name = "Target";
    entity2.entityalias = "Target";
    entity2["sdkmessageprocessingstepid@odata.bind"] = "/sdkmessageprocessingsteps(" + newStepId + ")";  // Parent Step created above
    entity2.attributes = fieldList; // List of fields to log

    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        datatype: "json",
        url: Xrm.Page.context.getClientUrl() + "/api/data/v9.1/sdkmessageprocessingstepimages",
        data: JSON.stringify(entity2),
        beforeSend: function (XMLHttpRequest) {
            XMLHttpRequest.setRequestHeader("OData-MaxVersion", "4.0");
            XMLHttpRequest.setRequestHeader("OData-Version", "4.0");
            XMLHttpRequest.setRequestHeader("Accept", "application/json");
        },
        async: true,
        success: function (data, textStatus, xhr) {
            var uri2 = xhr.getResponseHeader("OData-EntityId");
            var regExp = /\(([^)]+)\)/;
            var matches2 = regExp.exec(uri2);
            var newEntityId2 = matches2[1];
            MSFTPP.CAB.showConsoleLog("Image Created! New ID = " + newEntityId2);
            $("#divPnlSuccess").show();
            setTimeout(MSFTPP.CAB.pageOnLoad, 5000);
            setTimeout(MSFTPP.CAB.hideSuccessPanel, 5000);
        },
        error: function (xhr, textStatus, errorThrown) {
            MSFTPP.CAB.handleError(xhr.responseText);
        }
    });

}; // end createImage

MSFTPP.CAB.checkAllFields = function (check) {
    var ictr = 0;
    selectedFields = [];

    if (check) {
        MSFTPP.CAB.showConsoleLog("Check ALL Fields!");

        $("input[type=checkbox]").each(function () {
            $(this).prop("checked", true);
            if (this.value.toString() !== "0") {
                let tableColumn = {
                    Id: this.value.toString(),
                    Name: this.name,
                    MetadataId: this.id.toString(),
                };
                selectedFields.push(tableColumn);
            }
            ictr = ictr + 1;
        });
    } else {
        MSFTPP.CAB.showConsoleLog("Un-Check ALL!");
        $("input[type=checkbox]").each(function () {
            $(this).prop("checked", false);
            ictr = ictr + 1;
        });
    }

    MSFTPP.CAB.updateOutputPanel();
};