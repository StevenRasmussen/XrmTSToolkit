## Welcome to XrmTSToolkit GitHub page
XrmTSToolkit is a TypeScript library that you can use to perform all the basic Dynamics CRM webservice methods.  This library was designed to be similar in nature to the CRM SDK such that anyone familiar with the .Net SDK should be able to use the library without much help.  Where possible, all the same methods and overloads have been reproduced.

XrmTSToolkit also uses jQuery Promises so that every method is run Asynchronously and will be familiar to anyone who has already worked with jQuery promises before.

The following methods are supported:

### Soap Methods
* Create
* Update
* Delete
* Retrieve
* RetrieveMultiple
* RetrieveRelatedManyToMany
* Associate
* Disassociate
* Fetch
* Execute

### REST methods
XrmTSToolkit does not support the REST endpoint.  This may be added in a future release.

## Sample Code:
### Create
```javascript

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

var Promise = XrmTSToolkit.Soap.Create(Entity);
Promise.done(function (data: XrmTSToolkit.Soap.CreateSoapResponse, result, xhr) {
    var NewAccountId = data.CreateResult;
});
Promise.fail(function (result) {
    //Error occured
});

```

### Update

```typescript
var Entity = new XrmTSToolkit.Soap.Entity("account", "9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA");
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

var Promise = XrmTSToolkit.Soap.Update(Entity);
Promise.done(function (data: XrmTSToolkit.Soap.UpdateSoapResponse, result, xhr) {
    //Successfully updated
});
Promise.fail(function (result) {
    //Error occured
});

```

### Delete

```typescript
var Promise = XrmTSToolkit.Soap.Delete(new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));
Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Successfully deleted
});
Promise.fail(function (result) {
    //Error occured
});
```

### Retrieve

```typescript
var Promise = XrmTSToolkit.Soap.Retrieve("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account", new XrmTSToolkit.Soap.ColumnSet(true));
Promise.done(function (data: XrmTSToolkit.Soap.RetrieveSoapResponse, result, xhr) {
    var Entity = data.RetrieveResponse;
    var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue> Entity.Attributes["creditonhold"]).Value;
    var OwnerId = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Id;
    var OwnerName = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Name;
});
Promise.fail(function (result) {
    //Error occured
});
```

### RetrieveMultiple

```typescript
var Query = new XrmTSToolkit.Soap.Query.QueryExpression("account");
Query.Columns = new XrmTSToolkit.Soap.ColumnSet(true);
Query.Criteria = new XrmTSToolkit.Soap.Query.FilterExpression(XrmTSToolkit.Soap.Query.LogicalOperator.And);
Query.Criteria.AddCondition(new XrmTSToolkit.Soap.Query.ConditionExpression("name", XrmTSToolkit.Soap.Query.ConditionOperator.Equal, new XrmTSToolkit.Soap.StringValue("Test Account")));
var Promise = XrmTSToolkit.Soap.RetrieveMultiple(Query);
Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
    var Entities = data.RetrieveMultipleResult.Entities;
    $.each(Entities, function (i, Entity) {
        var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue> Entity.Attributes["creditonhold"]).Value;
        var OwnerId = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Id;
        var OwnerName = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Name;
    });
});
Promise.fail(function (result) {
    //Error occured
});
```

### RetrieveRelatedManyToMany

```typescript
var Promise = XrmTSToolkit.Soap.RetrieveRelatedManyToMany("account", "9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "systemuser", "new_account_systemuser", new XrmTSToolkit.Soap.ColumnSet(false));
Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
    var Entities = data.RetrieveMultipleResult.Entities;
    $.each(Entities, function (i, Entity) {
        var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue> Entity.Attributes["creditonhold"]).Value;
        var OwnerId = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Id;
        var OwnerName = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Name;
    });
});
Promise.fail(function (result) {
    //Error occured
});
```

### Associate

```typescript
var Promise = XrmTSToolkit.Soap.Associate(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    "new_account_systemuser");

Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Associate completed successfully
});
Promise.fail(function (result) {
    //Error occurred
});
```

### Disassociate

```typescript
var Promise = XrmTSToolkit.Soap.Disassociate(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    "new_account_systemuser");

Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Disassociate completed successfully
});
Promise.fail(function (result) {
    //Error occurred
});
```

### Fetch

```typescript
var FetchXML = "" +
    "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
    "<entity name=\"account\">" +
    "<attribute name=\"name\" />" +
    "<attribute name=\"primarycontactid\" />" +
    "<attribute name=\"telephone1\" />" +
    "<attribute name=\"accountid\" />" +
    "<order attribute=\"name\" descending=\"false\" />" +
    "<filter type=\"and\">" +
    "<condition attribute=\"name\" operator=\"eq\" value=\"Test Account - updated\" />" +
    "</filter>" +
    "</entity>" +
    "</fetch>";
var Promise = XrmTSToolkit.Soap.Fetch(FetchXML);
Promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
    var Entities = data.RetrieveMultipleResult.Entities;
    $.each(Entities, function (i, Entity) {
        var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue> Entity.Attributes["creditonhold"]).Value;
        var OwnerId = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Id;
        var OwnerName = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Name;
    });
});
Promise.fail(function (result) {
    dfd.reject(new TestResult(false, "Fetch test failed: " + result.response, result));
});
```

### Execute

```typescript
//Execute a 'WhoAmI' request
var ExecuteXML = "" +
    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
    "<request i:type=\"b:WhoAmIRequest\" xmlns:a = \"http://schemas.microsoft.com/xrm/2011/Contracts\" xmlns:b = \"http://schemas.microsoft.com/crm/2011/Contracts\">" +
    "<a:Parameters xmlns:c = \"http://schemas.datacontract.org/2004/07/System.Collections.Generic\" />" +
    "<a:RequestId i:nil = \"true\" />" +
    "<a:RequestName>WhoAmI</a:RequestName>" +
    "</request>" +
    "</Execute>";

var Promise = XrmTSToolkit.Soap.Execute(ExecuteXML);
Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //WhoAmIRequest executed successfully
});
Promise.fail(function (result) {
    //Error occurred
});
```
