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
* SetState
* Execute

#### Soap Requests - Used with 'Execute' or 'ExecuteMultiple'
* ExecuteMultipleRequest
* CreateRequest
* UpdateRequest
* DeleteRequest
* AssociateRequest
* DisassociateRequest
* SetStateRequest
* WhoAmIRequest
* AssignRequest
* GrantAccessRequest
* ModifyAccessRequest
* RevokeAccessRequest
* RetrievePrincipleAccessRequest

#### Other Organization Requests
Any organization request not listed above can easily be implemented by inheriting from the 'ExecuteRequest' class. An example is shown below illustrating how the 'CreateRequest' is implemented. Additional requests may be added to the XrmTSToolkit as deemed necessary or perhaps a separate TypeScript file will be created that will contain additional organization requests.  Please submit an issue on GitHub for consideration of another request to be added to the main library.

#### Custom Actions
You can execute your own custom actions very easily by following the example given below.

#### Retrieve Metadata
XrmTSToolkit allows you To retrieve entity metadata from CRM.

### REST methods
XrmTSToolkit does not support the REST endpoint.  This may be added in a future release.

## Sample Code:
### Create
```typescript

var entity = new XrmTSToolkit.Soap.Entity("account");
entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(true);
entity.Attributes["donotemail"] = true; //You can also just pass a bool instead of a 'BooleanValue'
entity.Attributes["creditlimit"] = new XrmTSToolkit.Soap.MoneyValue(1000);
entity.Attributes["lastusedincampaign"] = new XrmTSToolkit.Soap.DateValue(new Date());
entity.Attributes["exchangerate"] = new XrmTSToolkit.Soap.DecimalValue(2000);
entity.Attributes["address1_latitude"] = new XrmTSToolkit.Soap.FloatValue(90);
entity.Attributes["numberofemployees"] = new XrmTSToolkit.Soap.IntegerValue(4000);
entity.Attributes["ownerid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
entity.Attributes["description"] = new XrmTSToolkit.Soap.StringValue("This is a long string value");
entity.Attributes["telephone1"] = "(999) 123-4567"; //You can also just pass a string instead of a 'StringValue'
entity.Attributes["accountcategorycode"] = new XrmTSToolkit.Soap.OptionSetValue(1);
entity.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Test Account");

var promise = XrmTSToolkit.Soap.Create(entity);
promise.done(function (data: XrmTSToolkit.Soap.CreateSoapResponse, result, xhr) {
    var newAccountId = data.CreateResult;
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});

```
### Create Email
```typescript

var email = new XrmTSToolkit.Soap.Entity("email");
email.Attributes["subject"] = "email subject";

var fromActivityParties = new XrmTSToolkit.Soap.EntityCollection();
var fromActivityParty = new XrmTSToolkit.Soap.Entity("activityparty");
fromActivityParty.Attributes["partyid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
fromActivityParties.Items.push(fromActivityParty);
email.Attributes["from"] = new XrmTSToolkit.Soap.EntityCollectionAttribute(fromActivityParties);

var toActivityParties = new XrmTSToolkit.Soap.EntityCollection();
var toActivityParty = new XrmTSToolkit.Soap.Entity("activityparty");
toActivityParty.Attributes["partyid"] = new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "contact");
toActivityParties.Items.push(toActivityParty);
email.Attributes["to"] = new XrmTSToolkit.Soap.EntityCollectionAttribute(toActivityParties);

XrmTSToolkit.Soap.Create(email).done(function (emailResponse: XrmTSToolkit.Soap.CreateSoapResponse) {
	var emailId = emailResponse.CreateResult;
	dfd.resolve(new TestResult(true, "CreateEmail test succeeded: " + emailId, PriorTestResult.ResultValue));
}).fail(function (result) {
	dfd.reject(new TestResult(false, "CreateEmail test failed: " + result.faultstring, result));
});

```

### Update

```typescript
var entity = new XrmTSToolkit.Soap.Entity("account", "9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA");
entity.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(true);
entity.Attributes["creditlimit"] = new XrmTSToolkit.Soap.MoneyValue(10000);
entity.Attributes["lastusedincampaign"] = new XrmTSToolkit.Soap.DateValue(new Date());
entity.Attributes["exchangerate"] = new XrmTSToolkit.Soap.DecimalValue(20000);
entity.Attributes["address1_latitude"] = new XrmTSToolkit.Soap.FloatValue(-90);
entity.Attributes["numberofemployees"] = new XrmTSToolkit.Soap.IntegerValue(40000);
entity.Attributes["ownerid"] = new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser");
entity.Attributes["description"] = new XrmTSToolkit.Soap.StringValue("This is a long string value - updated");
entity.Attributes["accountcategorycode"] = new XrmTSToolkit.Soap.OptionSetValue(2);
entity.Attributes["name"] = new XrmTSToolkit.Soap.StringValue("Test Account - updated");

var promise = XrmTSToolkit.Soap.Update(entity);
promise.done(function (data: XrmTSToolkit.Soap.UpdateSoapResponse, result, xhr) {
    //Successfully updated
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});

```

### Delete

```typescript
var promise = XrmTSToolkit.Soap.Delete(new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));
promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Successfully deleted
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});
```

### Retrieve

```typescript
var promise = XrmTSToolkit.Soap.Retrieve("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account", new XrmTSToolkit.Soap.ColumnSet(true));
promise.done(function (data: XrmTSToolkit.Soap.RetrieveSoapResponse, result, xhr) {
    var entity = data.RetrieveResponse;
	var stringValue = (<XrmTSToolkit.Soap.StringValue> entity.Attributes["name"]).Value;
    var booleanValue = (<XrmTSToolkit.Soap.BooleanValue> entity.Attributes["creditonhold"]).Value;
	var dateValue = (<XrmTSToolkit.Soap.DateValue> entity.Attributes["createdon"]).Value;
	var integerValue = (<XrmTSToolkit.Soap.IntegerValue> entity.Attributes["createdon"]).Value;
    var ownerId = (<XrmTSToolkit.Soap.EntityReference> entity.Attributes["ownerid"]).Id;
    var ownerName = (<XrmTSToolkit.Soap.EntityReference> entity.Attributes["ownerid"]).Name;

	// You can also get the values using the following:
	stringValue = entity.getAttribute<XrmTSToolkit.Soap.StringValue>("name");
	booleanValue = entity.getAttribute<XrmTSToolkit.Soap.BooleanValue>("creditonhold");
	dateValue = entity.getAttribute<XrmTSToolkit.Soap.DateValue>("creditonhold");
	integerValue = entity.getAttribute<XrmTSToolkit.Soap.IntegerValue>("creditonhold");

	// Alternatively you can get the raw value using the corresponding method:
	var stringRaw = entity.getString("name");
	var boolRaw = entity.getBool("creditonhold");
	var dateRaw = entity.getDate("createdon");
	var numberRaw = entity.getNumber("numberofemployees");
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});
```

### RetrieveMultiple

```typescript
var Query = new XrmTSToolkit.Soap.Query.QueryExpression("account");
Query.Columns = new XrmTSToolkit.Soap.ColumnSet(true);
Query.Criteria = new XrmTSToolkit.Soap.Query.FilterExpression(XrmTSToolkit.Soap.Query.LogicalOperator.And);
Query.Criteria.AddCondition(new XrmTSToolkit.Soap.Query.ConditionExpression("name", XrmTSToolkit.Soap.Query.ConditionOperator.Equal, new XrmTSToolkit.Soap.StringValue("Test Account")));
var promise = XrmTSToolkit.Soap.RetrieveMultiple(Query);
promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
    var entities = data.RetrieveMultipleResult.Entities;
    $.each(entities, function (i, entity) {
        var booleanValue = (<XrmTSToolkit.Soap.BooleanValue> entity.Attributes["creditonhold"]).Value;
        var ownerId = (<XrmTSToolkit.Soap.EntityReference> entity.Attributes["ownerid"]).Id;
        var ownerName = (<XrmTSToolkit.Soap.EntityReference> entity.Attributes["ownerid"]).Name;
    });
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});
```

### RetrieveRelatedManyToMany

```typescript
var promise = XrmTSToolkit.Soap.RetrieveRelatedManyToMany("account", "9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "systemuser", "new_account_systemuser", new XrmTSToolkit.Soap.ColumnSet(false));
promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
    var entities = data.RetrieveMultipleResult.Entities;
    $.each(entities, function (i, entity) {
        var booleanValue = (<XrmTSToolkit.Soap.BooleanValue> entity.Attributes["creditonhold"]).Value;
        var ownerId = (<XrmTSToolkit.Soap.EntityReference> entity.Attributes["ownerid"]).Id;
        var ownerName = (<XrmTSToolkit.Soap.EntityReference> entity.Attributes["ownerid"]).Name;
    });
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});
```

### Associate

```typescript
var promise = XrmTSToolkit.Soap.Associate(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    "new_account_systemuser");

promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Associate completed successfully
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occurred
});
```

### Disassociate

```typescript
var promise = XrmTSToolkit.Soap.Disassociate(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    "new_account_systemuser");

promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Disassociate completed successfully
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
var promise = XrmTSToolkit.Soap.Fetch(FetchXML);
promise.done(function (data: XrmTSToolkit.Soap.RetrieveMultipleSoapResponse, result, xhr) {
    var Entities = data.RetrieveMultipleResult.Entities;
    $.each(Entities, function (i, Entity) {
        var BooleanValue = (<XrmTSToolkit.Soap.BooleanValue> Entity.Attributes["creditonhold"]).Value;
        var OwnerId = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Id;
        var OwnerName = (<XrmTSToolkit.Soap.EntityReference> Entity.Attributes["ownerid"]).Name;
    });
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occurred
});
```

### Execute (AssignRequest Example)

```typescript
var assignRequest = new XrmTSToolkit.Soap.AssignRequest(
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));

var promise = XrmTSToolkit.Soap.Execute(assignRequest);
promise.done(function (data: XrmTSToolkit.Soap.AssignResponse, result, xhr) {
    //Assign completed successfully
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Assign Request failed
});
```

### Execute (WhoAmI XML Example)

```typescript
//Execute a 'WhoAmI' request
var executeXML = "" +
    "<Execute xmlns=\"http://schemas.microsoft.com/xrm/2011/Contracts/Services\" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\">" +
    "<request i:type=\"b:WhoAmIRequest\" xmlns:a = \"http://schemas.microsoft.com/xrm/2011/Contracts\" xmlns:b = \"http://schemas.microsoft.com/crm/2011/Contracts\">" +
    "<a:Parameters xmlns:c = \"http://schemas.datacontract.org/2004/07/System.Collections.Generic\" />" +
    "<a:RequestId i:nil = \"true\" />" +
    "<a:RequestName>WhoAmI</a:RequestName>" +
    "</request>" +
    "</Execute>";

var executeRequest = new XrmTSToolkit.Soap.RawExecuteRequest(executeXml);

var promise = XrmTSToolkit.Soap.Execute(executeRequest);
promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //WhoAmIRequest executed successfully
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occurred
});
```

### ExecuteMultiple

```typescript
var accountId = "9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA";
var accountToUpdate = new XrmTSToolkit.Soap.Entity("account", accountId);
accountToUpdate.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(false);

//Create the 'UpdateRequest'
var updateRequest = new XrmTSToolkit.Soap.UpdateRequest(accountToUpdate);

//Create the 'DeleteRequest'
var deleteRequest = new XrmTSToolkit.Soap.DeleteRequest(new XrmTSToolkit.Soap.EntityReference(accountId, "account"));

//Create the 'ExecuteMultipleRequest' and initialize the settings to 'ContinueOnError' and 'ReturnResponses'
var executeMultipleRequest = new XrmTSToolkit.Soap.ExecuteMultipleRequest();
executeMultipleRequest.Settings.ContinueOnError = true;
executeMultipleRequest.Settings.ReturnResponses = true;

//Add the requests to the 'ExecuteMultipleRequest'
executeMultipleRequest.Requests.push(updateRequest);
executeMultipleRequest.Requests.push(deleteRequest);

//Generate the 'RequestMultiple' promise and cast the return type to 'ExecuteMultipleResponse'
var promise = XrmTSToolkit.Soap.Execute<XrmTSToolkit.Soap.ExecuteMultipleResponse>(executeMultipleRequest);
promise.done(function (data, result, xhr) {
    $.each(data.Responses, function (i, responseItem) {
        if (responseItem.Fault) {
            //There was an error with the specific request
            
        }
    });
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //ExecuteMultiple request failed
});
```

### AssignRequest

```typescript
var assignRequest = new XrmTSToolkit.Soap.AssignRequest(
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));

var promise = XrmTSToolkit.Soap.Execute(assignRequest);
promise.done(function (data: XrmTSToolkit.Soap.AssignResponse, result, xhr) {
    //Success
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Assign request failed
});
```

### GrantAccessRequest

```typescript
var grantAccessRequest = new XrmTSToolkit.Soap.GrantAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    XrmTSToolkit.Soap.AccessRights.ShareAccess);

var promise = XrmTSToolkit.Soap.Execute(grantAccessRequest);
promise.done(function (data: XrmTSToolkit.Soap.GrantAccessResponse, result, xhr) {
    //Grant access succeeded
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Grant access failed
});
```

### ModifyAccessRequest

```typescript
var modifyAccessRequest = new XrmTSToolkit.Soap.ModifyAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    XrmTSToolkit.Soap.AccessRights.WriteAccess);

var promise = XrmTSToolkit.Soap.Execute(modifyAccessRequest);
promise.done(function (data: XrmTSToolkit.Soap.GrantAccessResponse, result, xhr) {
    //Modify access completed successfully
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Modify access failed
});
```

### RevokeAccessRequest

```typescript
var revokeAccessRequest = new XrmTSToolkit.Soap.RevokeAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"));

var promise = XrmTSToolkit.Soap.Execute(revokeAccessRequest);
promise.done(function (data: XrmTSToolkit.Soap.RevokeAccessResponse, result, xhr) {
    //Revoke access completed successfully
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Revoke access failed
});
```

### RetrievePrincipleAccessRequest

```typescript
var retrieveAccessRequest = new XrmTSToolkit.Soap.RetrievePrincipleAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"));

var promise = XrmTSToolkit.Soap.Execute(retrieveAccessRequest);
promise.done(function (data: XrmTSToolkit.Soap.RetrievePrincipleAccessResponse, result, xhr) {
    //Retrieve principal access succeeded.  Iterate through all the access rights and get their text
    var rightsStrings = [];
    $.each(data.AccessRights, function (i, right) {
        rightsStrings.push(XrmTSToolkit.Soap.AccessRights[right].toString());
    });
});
promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Retrieve principal access failed
});
```

### Custom Actions
To use custom actions it is required that you create a new TypeScript class that extends the 'ExecuteRequest' class.  The following is an example of how to accomplish it:

```typescript
export class CustomActionRequest extends XrmTSToolkit.Soap.CustomActionRequest {
	constructor(account: XrmTSToolkit.Soap.EntityReference, stringValue: string) {
		super("new_CustomActionName"); //This is the name of the custom action defined in CRM
        this.Parameters["Target"] = account;
        this.Parameters["StringValue"] = stringValue; // Pass in any values needed for the custom action as 'Parameters'
	}

	// Optionally override the following method if you have created a specific response type class as shown below
	CreateResponse(responseXml: string): CustomActionResponse {
        return new CustomActionResponse(responseXml);
    }
}
```

You can also choose To create an appropriate response to make it easier to read any results:
```typescript
export class CustomActionResponse extends XrmTSToolkit.Soap.ExecuteResponse {
	constructor(responseXml: string) {
        super(responseXml);
		// In order for XrmTSToolkit to know how to deserialize the 
        this.PropertyTypes["ReturnValueString"] = "s"; // s == string
        this.PropertyTypes["ReturnValueBool"] = "b"; // b == bool
        this.PropertyTypes["ReturnValueInt"] = "n"; // n == number
        this.PropertyTypes["ReturnValueMoney"] = "n";
    }
    ReturnValueString: string;
    ReturnValueBool: boolean;
    ReturnValueInt: number;
    ReturnValueMoney: number;
}
```

To execute the custom action, you simply instantiate an instance of your custom action request and pass it to the 'execute' method:
```typescript
var customActionRequest = new CustomActionRequest(new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"), "some string value");
XrmTSToolkit.Soap.Execute(customActionRequest).done(function (executeResponse:CustomActionResponse) {
	//Custom action executed successfully
	var result = executeResponse.Result;
}).fail(function (error) {
	//Custom action failed
});
```

### How To Create Your Own Organization Requests

The following example shows how the 'CreateRequest' is implemented so that you can learn to implement other requests as necessary.  For this example we will be basing everything off of the following XML Soap request and response:

##### Soap Request

```xml
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <Execute xmlns="http://schemas.microsoft.com/xrm/2011/Contracts/Services" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
      <request i:type="a:CreateRequest" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
        <a:Parameters xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
          <a:KeyValuePairOfstringanyType>
            <b:key>Target</b:key>
            <b:value i:type="a:Entity">
              <a:Attributes />
              <a:EntityState i:nil="true" />
              <a:FormattedValues />
              <a:Id>00000000-0000-0000-0000-000000000000</a:Id>
              <a:LogicalName>account</a:LogicalName>
              <a:RelatedEntities />
            </b:value>
          </a:KeyValuePairOfstringanyType>
        </a:Parameters>
        <a:RequestId i:nil="true" />
        <a:RequestName>Create</a:RequestName>
      </request>
    </Execute>
  </s:Body>
</s:Envelope>
```

##### Soap Response

```xml
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Body>
    <ExecuteResponse xmlns="http://schemas.microsoft.com/xrm/2011/Contracts/Services" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
      <ExecuteResult i:type="a:CreateResponse" xmlns:a="http://schemas.microsoft.com/xrm/2011/Contracts">
        <a:ResponseName>Create</a:ResponseName>
        <a:Results xmlns:b="http://schemas.datacontract.org/2004/07/System.Collections.Generic">
          <a:KeyValuePairOfstringanyType>
            <b:key>id</b:key>
            <b:value i:type="c:guid" xmlns:c="http://schemas.microsoft.com/2003/10/Serialization/">6cabef12-aece-e411-80be-00155d017a0b</b:value>
          </a:KeyValuePairOfstringanyType>
        </a:Results>
      </ExecuteResult>
    </ExecuteResponse>
  </s:Body>
</s:Envelope>
```

First, create your custom class that inherits from the 'ExecuteRequest':

```typescript
export class CreateRequest extends ExecuteRequest {
```

Create your constructor and pass in any necessary parameters to complete the request, in this case all we need is the 'Entity' to be created. 

```typescript
    constructor(Target: Entity) {
```

Also make a call to the base class constructor passing in the name of the soap request and optionally specify the soap type with the namespace if necessary:

```typescript
        super("Create", "a:CreateRequest");
```

*Note: Generally you will only need to pass in the actual name of the request, ie "Create" in this instance.  However, because the soap type differs from the other requests and namespace we must explicitely specify the type and namespace for the 'CreateRequest'. Here is a list of the namespaces used by XrmTSToolkit:

```typescript
//The default namespace used by XrmTSToolkit to serialize 'Execute' messages is "g" below. If the namespace for your message differs then you will need to specify it by using the list below.
var ns = {
"xmlns" : "http://schemas.microsoft.com/xrm/2011/Contracts/Services",
"s": "http://schemas.xmlsoap.org/soap/envelope/",
"a": "http://schemas.microsoft.com/xrm/2011/Contracts",
"i": "http://www.w3.org/2001/XMLSchema-instance",
"b": "http://schemas.datacontract.org/2004/07/System.Collections.Generic",
"c": "http://www.w3.org/2001/XMLSchema",
"e": "http://schemas.microsoft.com/2003/10/Serialization/",
"f": "http://schemas.microsoft.com/2003/10/Serialization/Arrays",
"g": "http://schemas.microsoft.com/crm/2011/Contracts",
"h": "http://schemas.microsoft.com/xrm/2011/Metadata",
"j": "http://schemas.microsoft.com/xrm/2011/Metadata/Query",
"k": "http://schemas.microsoft.com/xrm/2013/Metadata",
"l": "http://schemas.microsoft.com/xrm/2012/Contracts"};
```

Next, in the contructor method, add the parameters to the 'this.Parameters' named array. The name must match exactly what the actual Soap message requires:

```typescript
        this.Parameters["Target"] = Target;
    }
```

Optionally create a response class.  This is helpful if you are expecting a result and need to query it. For the create request the id of the newly created record is returned:

```typescript
export class CreateResponse extends ExecuteResponse {
    id: string;
}
```

Putting it all together. Here is both the request and response in their entirety:

```typescript
export class CreateRequest extends ExecuteRequest {
    /**
     * Constructor.
     *
     * @param   {Entity}    Target  Entity to be created.
     */
    constructor(Target: Entity) {
        super("Create", "a:CreateRequest");
        this.Parameters["Target"] = Target;
    }
}
export class CreateResponse extends ExecuteResponse {
    id: string;
}
```

XrmTSToolkit is able to serialize the new class and execute the message appropriately.
