/// <reference path="../scripts/typings/jquery/jquery.d.ts" />
/// <reference path="../scripts/typings/xrm/xrmtstoolkit.ts" />

function OpenXrmTSToolkitTestPage() {
    var URL = Xrm.Page.context.getClientUrl() + "/WebResources/new_/Tests/XrmTSToolkitTests.html";
    //window.showModelessDialog(URL, Xrm);

    var DialogOptions = { width: 700, height: 600 };
    (<any> Xrm).Internal.openDialog(URL, DialogOptions, null, null, function () {

    });
}

class TestResult {
    constructor(public Result: boolean, public ResultMessage: string, public ResultValue?: any) { }
}

function RunTests(): void {
    try
    {
        $("#testresults").empty();

        var Tests = new Array<any>();
        Tests.push(CreateEntityTest1);
        Tests.push(UpdateEntityTest1);
        Tests.push(RetrieveEntityTest);
        Tests.push(AssociateTest);
        Tests.push(RetrieveManyToManyTest);
        Tests.push(DisassociateTest);
        Tests.push(RetrieveMultipleTest_AllColumns);
        Tests.push(RetrieveMultipleTest_QueryExpressionWithFilters);
        Tests.push(FetchTest);
        Tests.push(SetStateTest_SetInactive);
        Tests.push(SetStateTest_SetActive);
        Tests.push(AssignTest);
        Tests.push(RetrieveAccessTest);
        Tests.push(GrantAccessTest);
        Tests.push(ModifyAccessTest);
        Tests.push(RevokeAccessTest);
        Tests.push(DeleteTest1);
        Tests.push(ExecuteTest);
        Tests.push(WhoAmITest);
        Tests.push(FaultTest1);
        Tests.push(FaultTest2);
        Tests.push(CreateEntityTest2);
        Tests.push(UpdateEntityTest2);
        Tests.push(DeleteTest2);
        Tests.push(ExecuteMultipleTest1);
        Tests.push(ExecuteMultipleTest2);

        var CurrentFunctionIndex = 0;
        function TestComplete(results: TestResult) {
            if (results.Result) {
                //Success
                $("#testresults").append($("<div>" + results.ResultMessage + "</div>"));
                if (!(CurrentFunctionIndex >= Tests.length)) {
                    try {
                        var NewDeferred = Tests[CurrentFunctionIndex](results);
                        CurrentFunctionIndex += 1;
                        NewDeferred.always(TestComplete);
                    }
                    catch (e) {
                        $("#testresults").append($("<div style='color:red'>" + results.ResultMessage + "</div>"));
                        $("#testresults").append($("<div style='color:red'>Cancelling remaining tests.</div>"));
                    }
                }
                else {
                    $("#testresults").append($("<div style='color:green'>All Tests completed successfully!!!</div>"));
                }
            }
            else {
                //Failure
                $("#testresults").append($("<div style='color:red'>" + results.ResultMessage + "</div>"));
                $("#testresults").append($("<div style='color:red'>Cancelling remaining tests.</div>"));
            }
        }

        try
        {
            var Deferred = MainTestFunction();
            Deferred.always(TestComplete);
        }
        catch (e) {
            alert(e);
        }
    }
    catch (e) {
        alert(e.description);
    }
}

function MainTestFunction(): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        dfd.resolve(new TestResult(true, "Beginning the tests..."));
    }).promise();
}

function CreateEntityTest1(): JQueryPromise<TestResult> {
    var Entity = new XrmTSToolkit.Soap.Entity("account");
    Entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(true);
    Entity.Attributes["creditlimit"] = new XrmTSToolkit.Soap.MoneyValue(1000);
    Entity.Attributes["lastusedincampaign"] = new XrmTSToolkit.Soap.DateValue(new Date());
    Entity.Attributes["exchangerate"] = new XrmTSToolkit.Soap.DecimalValue(2000);
    Entity.Attributes["address1_latitude"] = new XrmTSToolkit.Soap.FloatValue(90);
    Entity.Attributes["numberofemployees"] = new XrmTSToolkit.Soap.IntegerValue(4000);
    Entity.Attributes["ownerid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
    Entity.Attributes["description"] = new XrmTSToolkit.Soap.StringValue("This is a long string value");
    Entity.Attributes["accountcategorycode"] = new XrmTSToolkit.Soap.OptionSetValue(1);
    Entity.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Test Account");

    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Create(Entity);
        Promise.done(function (data: XrmTSToolkit.Soap.CreateSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Create1 test succeeded", data.CreateResult));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Create1 test failed: " + result.faultstring, result));
        });
    }).promise();
}

function UpdateEntityTest1(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var Entity = new XrmTSToolkit.Soap.Entity("account", PriorTestResult.ResultValue);
    Entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(true);
    Entity.Attributes["creditlimit"] = new XrmTSToolkit.Soap.MoneyValue(10000);
    Entity.Attributes["lastusedincampaign"] = new XrmTSToolkit.Soap.DateValue(new Date());
    Entity.Attributes["exchangerate"] = new XrmTSToolkit.Soap.DecimalValue(20000);
    Entity.Attributes["address1_latitude"] = new XrmTSToolkit.Soap.FloatValue(-90);
    Entity.Attributes["numberofemployees"] = new XrmTSToolkit.Soap.IntegerValue(40000);
    Entity.Attributes["ownerid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
    Entity.Attributes["description"] = new XrmTSToolkit.Soap.StringValue("This is a long string value - updated");
    Entity.Attributes["accountcategorycode"] = new XrmTSToolkit.Soap.OptionSetValue(2);
    Entity.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Test Account - updated");

    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Update(Entity);
        Promise.done(function (data: XrmTSToolkit.Soap.UpdateSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Update1 test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Update1 test failed: " + result.faultstring, result));
        });
    }).promise();
}

function RetrieveEntityTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Retrieve(EntityId, "account", new XrmTSToolkit.Soap.ColumnSet(true));
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveSoapResponse, result, xhr) {
            var Entity = data.RetrieveResult;
            var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue> Entity.Attributes["creditonhold"]).Value;
            var OwnerId = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Id;
            var OwnerName = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Name;
            dfd.resolve(new TestResult(true, "Retrieve test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Retrieve test failed: " + result.faultstring, result));
        });
    }).promise();
}

function AssociateTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Associate(
            new XrmTSToolkit.Soap.EntityReference(EntityId, "account"),
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
            "new_account_systemuser");

        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Associate test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Associate test failed: " + result.faultstring, result));
        });
    }).promise();
}

function RetrieveManyToManyTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.RetrieveRelatedManyToMany("account", PriorTestResult.ResultValue, "systemuser", "new_account_systemuser", new XrmTSToolkit.Soap.ColumnSet(false));
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Retrieve ManyToManyTest test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Retrieve ManyToManyTest test failed: " + result.faultstring, result));
        });
    }).promise();
}

function DisassociateTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Disassociate(
            new XrmTSToolkit.Soap.EntityReference(EntityId, "account"),
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
            "new_account_systemuser");

        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Disassociate test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Disassociate test failed: " + result.faultstring, result));
        });
    }).promise();
}

function RetrieveMultipleTest_AllColumns(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Query = new XrmTSToolkit.Soap.Query.QueryExpression("account");
        Query.Columns = new XrmTSToolkit.Soap.ColumnSet(false);
        var Promise = XrmTSToolkit.Soap.RetrieveMultiple(Query);
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
            try {
                if (!data.RetrieveMultipleResult.Entities || data.RetrieveMultipleResult.Entities.length <= 0) {
                    throw "The results were empty";
                }
                else {
                    TestEntityValues(data.RetrieveMultipleResult.Entities);
                }
                dfd.resolve(new TestResult(true, "Retrieve Multiple test succeeded: count = " + (data.RetrieveMultipleResult.Entities.length), PriorTestResult.ResultValue));
            }
            catch (e) {
                dfd.reject(new TestResult(false, "Retrieve Multiple test succeeded but iterating over the properties failed: " + e, PriorTestResult.ResultValue));
            }
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Retrieve Multiple test failed: " + result.faultstring, result));
        });
    }).promise();
}

function RetrieveMultipleTest_QueryExpressionWithFilters(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Query = new XrmTSToolkit.Soap.Query.QueryExpression("account");
        Query.Columns = new XrmTSToolkit.Soap.ColumnSet(true);
        Query.Criteria = new XrmTSToolkit.Soap.Query.FilterExpression(XrmTSToolkit.Soap.Query.LogicalOperator.And);
        Query.Criteria.AddCondition(new XrmTSToolkit.Soap.Query.ConditionExpression("name", XrmTSToolkit.Soap.Query.ConditionOperator.Equal, new XrmTSToolkit.Soap.StringValue("Test Account - updated")));
        var Promise = XrmTSToolkit.Soap.RetrieveMultiple(Query);
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
            try {
                if (!data.RetrieveMultipleResult.Entities || data.RetrieveMultipleResult.Entities.length <= 0) {
                    throw "The results were empty";
                }
                else {
                    TestEntityValues(data.RetrieveMultipleResult.Entities);
                }
                dfd.resolve(new TestResult(true, "Retrieve Multiple with filters test succeeded", PriorTestResult.ResultValue));
            }
            catch (e) {
                dfd.reject(new TestResult(false, "Retrieve Multiple with filters test succeeded but iterating over the properties failed: " + e, PriorTestResult.ResultValue));
            }
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Retrieve Multiple with filters test failed: " + result.faultstring, result));
        });
    }).promise();
}

function FetchTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        //Remove the maximum limit by setting the 'RetrieveAllEntities' to 'true'
        XrmTSToolkit.Soap.RetrieveAllEntities = true;
        var FetchXML = "" +
            "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
            "<entity name=\"account\">" +
            "<attribute name=\"name\" />" +
            "<attribute name=\"primarycontactid\" />" +
            "<attribute name=\"telephone1\" />" +
            "<attribute name=\"accountid\" />" +
            "<order attribute=\"name\" descending=\"false\" />" +
            "<filter type=\"and\">" +
            "</filter>" +
            "</entity>" +
            "</fetch>";
        var Promise = XrmTSToolkit.Soap.Fetch(FetchXML);
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
            try {
                if (!data.RetrieveMultipleResult.Entities || data.RetrieveMultipleResult.Entities.length <= 0) {
                    throw "The results were empty";
                }
                else if (data.RetrieveMultipleResult.Entities.length == 5000)
                {
                    throw "The maximum result limit was reached. It should retrieve all available records";
                }
                else {
                    TestEntityValues(data.RetrieveMultipleResult.Entities);
                }
                dfd.resolve(new TestResult(true, "Fetch test succeeded: count = " + (data.RetrieveMultipleResult.Entities.length), PriorTestResult.ResultValue));
            }
            catch (e) {
                dfd.reject(new TestResult(false, "Fetch succeeded but iterating over the properties failed: " + e, PriorTestResult.ResultValue));
            }
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Fetch test failed: " + result.faultstring, result));
        });
    }).promise();
}

function TestEntityValues(Entities: Array<XrmTSToolkit.Soap.Entity>) {
    $.each(Entities, function (i, Entity) {
        for (var AttributeName in Entity.Attributes) {
            var Value = Entity.Attributes[AttributeName];
            if (Value && !Value.IsNull()) {
                switch (Value.Type) {
                    case XrmTSToolkit.Soap.AttributeType.Boolean:
                        var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue>Value).Value;
                        if (!(typeof BooleanValue === "boolean")) { throw "'BooleanValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Date:
                        var DateValue = (<XrmTSToolkit.Soap.DateValue>Value).Value;
                        if (!(DateValue instanceof Date)) { throw "'DateValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Decimal:
                        var NumberValue = (<XrmTSToolkit.Soap.DecimalValue>Value).Value;
                        if (!(typeof NumberValue === "number")) { throw "'DecimalValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Double:
                        var NumberValue = (<XrmTSToolkit.Soap.FloatValue>Value).Value;
                        if (!(typeof NumberValue === "number")) { throw "'FloatValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.EntityReference:
                        var LookupValue = (<XrmTSToolkit.Soap.EntityReference>Value);
                        if (!(LookupValue instanceof XrmTSToolkit.Soap.EntityReference)) { throw "'EntityReference' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Float:
                        var NumberValue = (<XrmTSToolkit.Soap.FloatValue>Value).Value;
                        if (!(typeof NumberValue === "number")) { throw "'FloatValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Guid:
                        var StringValue = (<XrmTSToolkit.Soap.GuidValue>Value).Value;
                        if (!(typeof StringValue === "string")) { throw "'GuidValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Integer:
                        var NumberValue = (<XrmTSToolkit.Soap.IntegerValue>Value).Value;
                        if (!(typeof NumberValue === "number")) { throw "'IntegerValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.Money:
                        var NumberValue = (<XrmTSToolkit.Soap.MoneyValue>Value).Value;
                        if (!(typeof NumberValue === "number")) { throw "'MoneyValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.OptionSetValue:
                        var NumberValue = (<XrmTSToolkit.Soap.OptionSetValue>Value).Value;
                        if (!(typeof NumberValue === "number")) { throw "'OptionSetValue' value did not match the expected type." }
                        break;
                    case XrmTSToolkit.Soap.AttributeType.String:
                        var StringValue = (<XrmTSToolkit.Soap.StringValue>Value).Value;
                        if (!(typeof StringValue === "string")) { throw "'StringValue' value did not match the expected type." }
                        break;
                    default:
                        throw "Please update the 'RetrieveMultipleTest' to handle the attribute type: " + XrmTSToolkit.Soap.AttributeType[Value.Type].toString();
                        break;
                }
            }
        }
    });
}

function ExecuteTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    //Execute a 'WhoAmI' request
    var ExecuteXML = "" +
        "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
        "<request i:type=\"b:WhoAmIRequest\" xmlns:a = \"http://schemas.microsoft.com/xrm/2011/Contracts\" xmlns:b = \"http://schemas.microsoft.com/crm/2011/Contracts\">" +
        "<a:Parameters xmlns:c = \"http://schemas.datacontract.org/2004/07/System.Collections.Generic\" />" +
        "<a:RequestId i:nil = \"true\" />" +
        "<a:RequestName>WhoAmI</a:RequestName>" +
        "</request>" +
        "</Execute>";

    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Execute(ExecuteXML);
        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Execute test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Execute test failed: " + result.faultstring, result));
        });
    }).promise();
}

function SetStateTest_SetInactive(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.SetState(new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"), 1, 2);
        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Set State Inactive test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Set State Inactive test failed: " + result.faultstring, result));
        });
    }).promise();
}

function SetStateTest_SetActive(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.SetState(new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"), new XrmTSToolkit.Soap.OptionSetValue(0), new XrmTSToolkit.Soap.OptionSetValue(1));
        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Set State Active test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Set State Active test failed: " + result.faultstring, result));
        });
    }).promise();
}

function AssignTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var AssignRequest = new XrmTSToolkit.Soap.AssignRequest(
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
            new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"));

        var Promise = XrmTSToolkit.Soap.Execute(AssignRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.AssignResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Assign test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Assign test failed: " + result.faultstring, result));
        });
    }).promise();
}

function GrantAccessTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var GrantAccessRequest = new XrmTSToolkit.Soap.GrantAccessRequest(
            new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"),
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
            XrmTSToolkit.Soap.AccessRights.ShareAccess);

        var Promise = XrmTSToolkit.Soap.Execute(GrantAccessRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.GrantAccessResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Grant Access test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Grant Access test failed: " + result.faultstring, result));
        });
    }).promise();
}

function RetrieveAccessTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var RetrieveAccessRequest = new XrmTSToolkit.Soap.RetrievePrincipleAccessRequest(
            new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"),
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"));

        var Promise = XrmTSToolkit.Soap.Execute(RetrieveAccessRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.RetrievePrincipleAccessResponse, result, xhr) {
            var RightsStrings = [];
            $.each(data.AccessRights, function (i, Right) {
                RightsStrings.push(XrmTSToolkit.Soap.AccessRights[Right].toString());
            });
            dfd.resolve(new TestResult(true, "Retrieve Access test succeeded: " + RightsStrings.join(", "), PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Retrieve Access test failed: " + result.faultstring, result));
        });
    }).promise();
}

function ModifyAccessTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var ModifyAccessRequest = new XrmTSToolkit.Soap.ModifyAccessRequest(
            new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"),
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
            XrmTSToolkit.Soap.AccessRights.WriteAccess);

        var Promise = XrmTSToolkit.Soap.Execute(ModifyAccessRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.GrantAccessResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Modify Access test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Modify Access test failed: " + result.faultstring, result));
        });
    }).promise();
}

function RevokeAccessTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var RevokeAccessRequest = new XrmTSToolkit.Soap.RevokeAccessRequest(
            new XrmTSToolkit.Soap.EntityReference(PriorTestResult.ResultValue, "account"),
            new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"));

        var Promise = XrmTSToolkit.Soap.Execute(RevokeAccessRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.RevokeAccessResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Revoke Access test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Revoke Access test failed: " + result.faultstring, result));
        });
    }).promise();
}

function DeleteTest1(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Delete(new XrmTSToolkit.Soap.EntityReference(EntityId, "account"));
        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Delete1 test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Delete1 test failed: " + result.faultstring, result));
        });
    }).promise();
}

function WhoAmITest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Execute(new XrmTSToolkit.Soap.WhoAmIRequest());
        Promise.done(function (data: XrmTSToolkit.Soap.WhoAmIResponse, result, xhr) {
            var UserId = data.UserId;
            if (!UserId) {
                throw "Error retrieving UserId";
            }
            dfd.resolve(new TestResult(true, "WhoAmI test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "WhoAmI test failed: " + result.faultstring, result));
        });
    }).promise();
}


function FaultTest1(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    //Try to update the deleted record
    var Entity = new XrmTSToolkit.Soap.Entity("account", PriorTestResult.ResultValue);
    Entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(false);

    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Update(Entity);
        Promise.done(function (data: XrmTSToolkit.Soap.UpdateSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(false, "FaultTest test failed - did not throw an exception.", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(true, "FaultTest test succeeded: " + result.detail.OrganizationServiceFault.Message));
        });
    }).promise();
}

function FaultTest2(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    //try to assign a record with missing information
    return $.Deferred<TestResult>(function (dfd) {
        var AssignRequest = new XrmTSToolkit.Soap.AssignRequest(
            new XrmTSToolkit.Soap.EntityReference("00000000-0000-0000-0000-000000000000", ""),
            new XrmTSToolkit.Soap.EntityReference("00000000-0000-0000-0000-000000000000", ""));

        var Promise = XrmTSToolkit.Soap.Execute(AssignRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.AssignResponse, result, xhr) {
            dfd.resolve(new TestResult(false, "FaultTest test failed - did not throw an exception.", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(true, "FaultTest test succeeded: " + result.detail.OrganizationServiceFault.Message));
        });
    }).promise();
}

function CreateEntityTest2(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var Entity = new XrmTSToolkit.Soap.Entity("account");
    Entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(true);
    Entity.Attributes["creditlimit"] = new XrmTSToolkit.Soap.MoneyValue(1000);
    Entity.Attributes["lastusedincampaign"] = new XrmTSToolkit.Soap.DateValue(new Date());
    Entity.Attributes["exchangerate"] = new XrmTSToolkit.Soap.DecimalValue(2000);
    Entity.Attributes["address1_latitude"] = new XrmTSToolkit.Soap.FloatValue(90);
    Entity.Attributes["numberofemployees"] = new XrmTSToolkit.Soap.IntegerValue(4000);
    Entity.Attributes["ownerid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
    Entity.Attributes["description"] = new XrmTSToolkit.Soap.StringValue("This is a long string value");
    Entity.Attributes["accountcategorycode"] = new XrmTSToolkit.Soap.OptionSetValue(1);
    Entity.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Test Account 2");

    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Execute(new XrmTSToolkit.Soap.CreateRequest(Entity));
        Promise.done(function (data: XrmTSToolkit.Soap.CreateResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Create2 test succeeded", data.id));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Create2 test failed: " + result.faultstring, result));
        });
    }).promise();
}
function UpdateEntityTest2(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var Entity = new XrmTSToolkit.Soap.Entity("account", PriorTestResult.ResultValue);
    Entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(true);
    Entity.Attributes["creditlimit"] = new XrmTSToolkit.Soap.MoneyValue(10000);
    Entity.Attributes["lastusedincampaign"] = new XrmTSToolkit.Soap.DateValue(new Date());
    Entity.Attributes["exchangerate"] = new XrmTSToolkit.Soap.DecimalValue(20000);
    Entity.Attributes["address1_latitude"] = new XrmTSToolkit.Soap.FloatValue(-90);
    Entity.Attributes["numberofemployees"] = new XrmTSToolkit.Soap.IntegerValue(40000);
    Entity.Attributes["ownerid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
    Entity.Attributes["description"] = new XrmTSToolkit.Soap.StringValue("This is a long string value - updated");
    Entity.Attributes["accountcategorycode"] = new XrmTSToolkit.Soap.OptionSetValue(2);
    Entity.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Test Account 2 - updated");

    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Execute(new XrmTSToolkit.Soap.UpdateRequest(Entity));
        Promise.done(function (data: XrmTSToolkit.Soap.UpdateSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Update2 test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Update2 test failed: " + result.faultstring, result));
        });
    }).promise();
}
function DeleteTest2(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var DeleteRequest = new XrmTSToolkit.Soap.DeleteRequest(new XrmTSToolkit.Soap.EntityReference(EntityId, "account"));
        var Promise = XrmTSToolkit.Soap.Execute(DeleteRequest);
        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Delete2 test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Delete2 test failed: " + result.faultstring, result));
        });
    }).promise();
}

function ExecuteMultipleTest1(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    //Create a new record first and then Update and Delete it in the execute multiple
    return $.Deferred<TestResult>(function (dfd) {
        var CreatePromise = CreateEntityTest1();
        CreatePromise.done(function (CreateTestResult: TestResult) {
            var EntityId = CreateTestResult.ResultValue;
            var EntityToUpdate = new XrmTSToolkit.Soap.Entity("account", EntityId);
            EntityToUpdate.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(false);
            var UpdateRequest = new XrmTSToolkit.Soap.UpdateRequest(EntityToUpdate);

            var DeleteRequest = new XrmTSToolkit.Soap.DeleteRequest(new XrmTSToolkit.Soap.EntityReference(EntityId, "account"));

            var ExecuteMultipleRequest = new XrmTSToolkit.Soap.ExecuteMultipleRequest();
            ExecuteMultipleRequest.Requests.push(UpdateRequest);
            ExecuteMultipleRequest.Requests.push(DeleteRequest);
            ExecuteMultipleRequest.Settings.ContinueOnError = true;
            ExecuteMultipleRequest.Settings.ReturnResponses = true;

            var RequestMultiplePromise = XrmTSToolkit.Soap.Execute<XrmTSToolkit.Soap.ExecuteMultipleResponse>(ExecuteMultipleRequest);
            RequestMultiplePromise.done(function (data, result, xhr) {
                var Thing = data;
                $.each(data.Responses, function (i, ResponseItem) {
                    if (ResponseItem.Fault) {
                        //There was an error with the specific request
                        
                    }
                    else {

                    }
                });
                dfd.resolve(new TestResult(true, "Execute Multiple test succeeded", PriorTestResult.ResultValue));
            });
            RequestMultiplePromise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
                dfd.reject(new TestResult(false, "Execute Multiple test failed: " + result.faultstring, result));
            });

        });
    }).promise();
}

function ExecuteMultipleTest2(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    //Create a bunch of records
    return $.Deferred<TestResult>(function (dfd) {
        var ExecuteMultipleRequest = new XrmTSToolkit.Soap.ExecuteMultipleRequest();
        ExecuteMultipleRequest.Settings.ContinueOnError = true;
        ExecuteMultipleRequest.Settings.ReturnResponses = true;
        var TotalRecordsToCreate = 10;
        for (var i: number = 0; i < TotalRecordsToCreate; i++) {
            var EntityToCreate = new XrmTSToolkit.Soap.Entity("account");
            EntityToCreate.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Account Test " + i);
            ExecuteMultipleRequest.Requests.push(new XrmTSToolkit.Soap.CreateRequest(EntityToCreate));
        }

        var RequestMultiplePromise = XrmTSToolkit.Soap.Execute<XrmTSToolkit.Soap.ExecuteMultipleResponse>(ExecuteMultipleRequest);
        RequestMultiplePromise.done(function (data, result, xhr) {
            $.each(data.Responses, function (i, ResponseItem) {
                var AccountId = (<XrmTSToolkit.Soap.CreateResponse>  ResponseItem.Response).id;
            });
            dfd.resolve(new TestResult(true, "Execute2 Multiple test succeeded: " + TotalRecordsToCreate.toString() + " records created", PriorTestResult.ResultValue));
        });
        RequestMultiplePromise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
            dfd.reject(new TestResult(false, "Execute2 Multiple test failed: " + result.faultstring, result));
        });
    }).promise();
}
