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

Other organization requests not listed can easily be implemented by inheriting from the 'ExecuteRequest' class. An example is shown below how the 'CreateRequest' is implemented so that you can create your own requests as needed. Additional requests may be added to the XrmTSToolkit as deemed necessary or perhaps a separate TypeScript file will be created that will contain additional organization requests.  Please submit an issue on GitHub for consideration of another request to be added to the main library.

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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occured
});

```

### Delete

```typescript
var Promise = XrmTSToolkit.Soap.Delete(new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));
Promise.done(function (data: XrmTSToolkit.Soap.SoapResponse, result, xhr) {
    //Successfully deleted
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occurred
});
```

### Execute (AssignRequest Example)

```typescript
var AssignRequest = new XrmTSToolkit.Soap.AssignRequest(
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));

var Promise = XrmTSToolkit.Soap.Execute(AssignRequest);
Promise.done(function (data: XrmTSToolkit.Soap.AssignResponse, result, xhr) {
    //Assign completed successfully
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Assign Request failed
});
```

### Execute (WhoAmI XML Example)

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
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Error occurred
});
```

### ExecuteMultiple

```typescript
var AccountId = "9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA";
var AccountToUpdate = new XrmTSToolkit.Soap.Entity("account", AccountId);
AccountToUpdate.Attributes["creditonhold"] = new XrmTSToolkit.Soap.BooleanValue(false);

//Create the 'UpdateRequest'
var UpdateRequest = new XrmTSToolkit.Soap.UpdateRequest(AccountToUpdate);

//Create the 'DeleteRequest'
var DeleteRequest = new XrmTSToolkit.Soap.DeleteRequest(new XrmTSToolkit.Soap.EntityReference(AccountId, "account"));

//Create the 'ExecuteMultipleRequest' and initialize the settings to 'ContinueOnError' and 'ReturnResponses'
var ExecuteMultipleRequest = new XrmTSToolkit.Soap.ExecuteMultipleRequest();
ExecuteMultipleRequest.Settings.ContinueOnError = true;
ExecuteMultipleRequest.Settings.ReturnResponses = true;

//Add the requests to the 'ExecuteMultipleRequest'
ExecuteMultipleRequest.Requests.push(UpdateRequest);
ExecuteMultipleRequest.Requests.push(DeleteRequest);

//Generate the 'RequestMultiple' promise and cast the return type to 'ExecuteMultipleResponse'
var RequestMultiplePromise = XrmTSToolkit.Soap.Execute<XrmTSToolkit.Soap.ExecuteMultipleResponse>(ExecuteMultipleRequest);
RequestMultiplePromise.done(function (data, result, xhr) {
    $.each(data.Responses, function (i, ResponseItem) {
        if (ResponseItem.Fault) {
            //There was an error with the specific request
            
        }
    });
});
RequestMultiplePromise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //ExecuteMultiple request failed
});
```

### AssignRequest

```typescript
var AssignRequest = new XrmTSToolkit.Soap.AssignRequest(
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"));

var Promise = XrmTSToolkit.Soap.Execute(AssignRequest);
Promise.done(function (data: XrmTSToolkit.Soap.AssignResponse, result, xhr) {
    //Success
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Assign request failed
});
```

### GrantAccessRequest

```typescript
var GrantAccessRequest = new XrmTSToolkit.Soap.GrantAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    XrmTSToolkit.Soap.AccessRights.ShareAccess);

var Promise = XrmTSToolkit.Soap.Execute(GrantAccessRequest);
Promise.done(function (data: XrmTSToolkit.Soap.GrantAccessResponse, result, xhr) {
    //Grant access succeeded
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Grant access failed
});
```

### ModifyAccessRequest

```typescript
var ModifyAccessRequest = new XrmTSToolkit.Soap.ModifyAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"),
    XrmTSToolkit.Soap.AccessRights.WriteAccess);

var Promise = XrmTSToolkit.Soap.Execute(ModifyAccessRequest);
Promise.done(function (data: XrmTSToolkit.Soap.GrantAccessResponse, result, xhr) {
    //Modify access completed successfully
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Modify access failed
});
```

### RevokeAccessRequest

```typescript
var RevokeAccessRequest = new XrmTSToolkit.Soap.RevokeAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"));

var Promise = XrmTSToolkit.Soap.Execute(RevokeAccessRequest);
Promise.done(function (data: XrmTSToolkit.Soap.RevokeAccessResponse, result, xhr) {
    //Revoke access completed successfully
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Revoke access failed
});
```

### RetrievePrincipleAccessRequest

```typescript
var RetrieveAccessRequest = new XrmTSToolkit.Soap.RetrievePrincipleAccessRequest(
    new XrmTSToolkit.Soap.EntityReference("9C8AF527-2D96-4ADB-9C0B-A21BF460CDDA", "account"),
    new XrmTSToolkit.Soap.EntityReference(Xrm.Page.context.getUserId(), "systemuser"));

var Promise = XrmTSToolkit.Soap.Execute(RetrieveAccessRequest);
Promise.done(function (data: XrmTSToolkit.Soap.RetrievePrincipleAccessResponse, result, xhr) {
    //Retrieve principal access succeeded.  Iterate through all the access rights and get their text
    var RightsStrings = [];
    $.each(data.AccessRights, function (i, Right) {
        RightsStrings.push(XrmTSToolkit.Soap.AccessRights[Right].toString());
    });
});
Promise.fail(function (result: XrmTSToolkit.Soap.FaultResponse) {
    //Retrieve principal access failed
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

'''typescript
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
'''

Next, in the contructor method, add the parameters to the 'this.Parameters' named array.  The name must match exactly what the actual Soap message requires

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

