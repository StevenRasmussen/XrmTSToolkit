/// <reference path="../scripts/typings/jquery/jquery.d.ts" />
/// <reference path="../scripts/typings/xrm/xrmtstoolkit.ts" />

function OpenXrmTSToolkitTestPage() {
    var URL = Xrm.Page.context.getClientUrl() + "/WebResources/new_/Tests/XrmTSToolkitTests.html";
    window.showModelessDialog(URL, Xrm);
}

class TestResult {
    constructor(public Result: boolean, public ResultMessage: string, public ResultValue?: any) { }
}

function RunTests(): void {
    try
    {
        $("#testresults").empty();

        var Tests = new Array<any>();
        Tests.push(CreateEntityTest);
        Tests.push(EditEntityTest);
        Tests.push(RetrieveEntityTest);
        Tests.push(AssociateTest);
        Tests.push(RetrieveManyToManyTest);
        Tests.push(DisassociateTest);
        Tests.push(RetrieveMultipleTest_AllColumns);
        Tests.push(RetrieveMultipleTest_QueryExpressionWithFilters);
        Tests.push(DeleteTest);

        var CurrentFunctionIndex = 0;
        function TestComplete(results: TestResult) {
            if (results.Result) {
                //Success
                $("#testresults").append($("<div>" + results.ResultMessage + "</div>"));
                if (!(CurrentFunctionIndex >= Tests.length)) {
                    var NewDeferred = Tests[CurrentFunctionIndex](results);
                    CurrentFunctionIndex += 1;
                    NewDeferred.always(TestComplete);
                }
            }
            else {
                //Failure
                $("#testresults").append($("<div style='color:red'>" + results.ResultMessage + "</div>"));
                $("#testresults").append($("<div style='color:red'>Cancelling remaining tests.</div>"));
            }
        }

        var Deferred = MainTestFunction();
        Deferred.always(TestComplete);
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

function CreateEntityTest(): JQueryPromise<TestResult> {
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
            dfd.resolve(new TestResult(true, "Create test succeeded", data.CreateResult));
        });
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Create test failed: " + result.response, result));
        });
    }).promise();
}

function EditEntityTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
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
            dfd.resolve(new TestResult(true, "Update test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Update test failed: " + result.response, result));
        });
    }).promise();
}

function RetrieveEntityTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Retrieve(EntityId, "account");
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Retrieve test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Retrieve test failed: " + result.response, result));
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
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Associate test failed: " + result.response, result));
        });
    }).promise();
}

function RetrieveManyToManyTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.RetrieveRelatedManyToMany("account", PriorTestResult.ResultValue, "systemuser", "new_account_systemuser", new XrmTSToolkit.Soap.ColumnSet(false));
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Retrieve ManyToManyTest test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Retrieve ManyToManyTest test failed: " + result.response, result));
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
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Disassociate test failed: " + result.response, result));
        });
    }).promise();
}

function RetrieveMultipleTest_AllColumns(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    return $.Deferred<TestResult>(function (dfd) {
        var Query = new XrmTSToolkit.Soap.Query.QueryExpression("account");
        Query.Columns = new XrmTSToolkit.Soap.ColumnSet(true);
        var Promise = XrmTSToolkit.Soap.RetrieveMultiple(Query);
        Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
            try {
                if (!data.RetrieveMultipleResult.Entities || data.RetrieveMultipleResult.Entities.length <= 0) {
                    throw "The results were empty";
                }
                else {
                    TestEntityValues(data.RetrieveMultipleResult.Entities);
                }
                dfd.resolve(new TestResult(true, "Retrieve Multiple test succeeded", PriorTestResult.ResultValue));
            }
            catch (e) {
                dfd.reject(new TestResult(false, "Retrieve Multiple test succeeded but iterating over the properties failed: " + e, PriorTestResult.ResultValue));
            }
        });
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Retrieve Multiple test failed: " + result.response, result));
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
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Retrieve Multiple with filters test failed: " + result.response, result));
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

function DeleteTest(PriorTestResult: TestResult): JQueryPromise<TestResult> {
    var EntityId = PriorTestResult.ResultValue;
    return $.Deferred<TestResult>(function (dfd) {
        var Promise = XrmTSToolkit.Soap.Delete(new XrmTSToolkit.Soap.EntityReference(EntityId, "account"));
        Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
            dfd.resolve(new TestResult(true, "Delete test succeeded", PriorTestResult.ResultValue));
        });
        Promise.fail(function (result) {
            dfd.reject(new TestResult(false, "Delete test failed: " + result.response, result));
        });
    }).promise();
}
