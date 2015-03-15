/// <reference path="../jquery/jquery.d.ts" />
/// <reference path="xrm.d.ts" />


/**
* MSCRM 2011, 2013, 2015 Service Toolkit for TypeScript
* @author Steven Rasmussen
* @current version : 0.6.0
* Credits:
*   The idea of this library was inspired by David Berry and Jaime Ji's XrmServiceToolkit.js
*    
* Date: March 12, 2015
* 
* Required TypeScript Version: 1.4
* 
* Required external libraries:
*   jquery.d.ts - Downloaded from Nuget and included in the project
* 
********************************************
* Version : 0.5.0
* Date: March 12, 2015
*   Initial Beta Release
********************************************
* Version : 0.6.0
* Date: March 13, 2015
*   Added 'Delete', 'Associate', 'Disassociate' methods
*   Added tests for each method
*   Fixed bugs found by tests :)
*   Re-wrote XML response parsing portion
*   Removed requirement for Xrm-2015.d.ts
*   Added Xrm.d.ts to package
*/
module XrmTSToolkit {
    export class Common {
        static GetServerURL(): string {
            var URL: string = document.location.protocol + "//" + document.location.host;
            var Org: string = Xrm.Page.context.getOrgUniqueName();
            if (document.location.pathname.toUpperCase().indexOf(Org.toUpperCase()) > -1) {
                URL += "/" + Org;
            }
            return URL;
        }
        static GetSoapServiceURL(): string {
            return Common.GetServerURL() + "/XRMServices/2011/Organization.svc/web";
        }
        static TestForContext(): void {
            if (typeof Xrm === "undefined") {
                throw new Error("Please make sure to add the ClientGlobalContext.js.aspx file to your web resource.");
            }
        }
        static GetURLParameter(name): string {
            return decodeURI((RegExp(name + '=' + '(.+?)(&|$)').exec(location.search) || [, null])[1]);
        }
        static StripGUID(GUID: string): string {
            return GUID.replace("{", "").replace("}", "");
        }
    }

    export class Soap {
        static Create(Entity: Soap.Entity): JQueryPromise<Soap.CreateSoapResponse> {
            var request = "<entity>" + Entity.Serialize() + "</entity>";
            return $.Deferred<Soap.CreateSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.CreateSoapResponse>(request, "Create");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Update(Entity: Soap.Entity): JQueryPromise<Soap.UpdateSoapResponse> {
            var request = "<entity>" + Entity.Serialize() + "</entity>";
            return $.Deferred<Soap.UpdateSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.UpdateSoapResponse>(request, "Update");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Delete(Entity: Soap.EntityReference);
        static Delete(EntityName: string, EntityId: string);
        static Delete(Entity: string | Soap.EntityReference, EntityId?: string): JQueryPromise<Soap.DeleteSoapResponse> {
            var EntityName = "";
            if (typeof Entity === "string") {
                EntityName = EntityName;
            }
            else if (Entity instanceof Soap.EntityReference) {
                EntityName = Entity.LogicalName;
                EntityId = Entity.Id;
            }
            var XML = "<entityName>" + EntityName + "</entityName>";
            XML += "<id>" + EntityId + "</id>";
            return $.Deferred<Soap.DeleteSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.DeleteSoapResponse>(XML, "Delete");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Retrieve(Id: string, EntityLogicalName: string, ColumnSet?: Soap.ColumnSet): JQueryPromise<Soap.RetrieveSoapResponse> {
            Id = Common.StripGUID(Id);
            if (!ColumnSet || ColumnSet == null) { ColumnSet = new Soap.ColumnSet(false); }
            var msgBody = "<entityName>" + EntityLogicalName + "</entityName><id>" + Id + "</id><columnSet>" + ColumnSet.serialize() + "</columnSet>";
            return $.Deferred<Soap.RetrieveSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.RetrieveSoapResponse>(msgBody, "Retrieve");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static RetrieveMultiple(Query: Soap.Query.QueryExpression): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.RetrieveMultipleSoapResponse>(Query.serialize(), "RetrieveMultiple");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static RetrieveRelatedManyToMany(
            LinkFromEntityName: string,
            LinkFromEntityId: string,
            LinkToEntityName: string,
            IntermediateTableName: string,
            Columns: Soap.ColumnSet = new Soap.ColumnSet(false),
            SortOrders: Array<Soap.Query.OrderExpression> = []): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {

            var Condition = new Soap.Query.ConditionExpression(LinkFromEntityName + "id", Soap.Query.ConditionOperator.Equal, new Soap.GuidValue(LinkFromEntityId));
            var LinkEntity1 = new Soap.Query.LinkEntity(LinkToEntityName, IntermediateTableName, LinkToEntityName + "id", LinkToEntityName + "id", Soap.Query.JoinOperator.Inner);
            LinkEntity1.LinkCriteria = new Soap.Query.FilterExpression();
            LinkEntity1.LinkCriteria.AddCondition(Condition);

            var LinkEntity2 = new Soap.Query.LinkEntity(IntermediateTableName, LinkFromEntityName, LinkFromEntityName + "id", LinkFromEntityName + "id", Soap.Query.JoinOperator.Inner);
            LinkEntity2.LinkCriteria = new Soap.Query.FilterExpression();
            LinkEntity2.LinkCriteria.AddCondition(Condition);
            LinkEntity1.LinkEntities.push(LinkEntity2);

            var Query = new Soap.Query.QueryExpression(LinkToEntityName);
            Query.Columns = Columns;
            Query.LinkEntities.push(LinkEntity1);
            Query.Orders = SortOrders;

            return Soap.RetrieveMultiple(Query);
        }
        static Fetch(fetchXml: string): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            //First decode all invalid characters
            fetchXml = fetchXml.replace(/&amp;/g, "&");
            fetchXml = fetchXml.replace(/&lt;/g, "<");
            fetchXml = fetchXml.replace(/&gt;/g, ">");
            fetchXml = fetchXml.replace(/&apos;/g, "'");
            fetchXml = fetchXml.replace(/&quot;/g, "\"");

            //now re-encode all invalid characters
            fetchXml = fetchXml.replace(/&/g, "&amp;");
            fetchXml = fetchXml.replace(/</g, "&lt;");
            fetchXml = fetchXml.replace(/>/g, "&gt;");
            fetchXml = fetchXml.replace(/'/g, "&apos;");
            fetchXml = fetchXml.replace(/\"/g, "&quot;");

            var msgBody = "<query i:type='a:FetchExpression'><a:Query>" + fetchXml + "</a:Query></query>";
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.RetrieveMultipleSoapResponse>(msgBody, "RetrieveMultiple");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
            The 'Execute' should be either the contents of the '<Execute>' tags,  INCLUDING the '<Execute>' tags or an 'ExecuteRequest' object
        */
        static Execute<T>(Execute: string | Soap.ExecuteRequest): JQueryPromise<T> {
            return $.Deferred<T>(function (dfd) {
                var ExecuteXML = "";
                if (typeof Execute === "string") {
                    ExecuteXML = Execute;
                }
                else if (Execute instanceof Soap.ExecuteRequest) {
                    ExecuteXML = Execute.Serialize();
                }
                var Request = Soap.DoRequest<T>(ExecuteXML, "Execute");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Associate(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string): JQueryPromise<Soap.SoapResponse> {
            var AssociateRequest = new Soap.AssociateRequest(Moniker1, Moniker2, RelationshipName);
            return $.Deferred<Soap.SoapResponse>(function (dfd) {
                var Request = Soap.Execute<Soap.SoapResponse>(AssociateRequest);
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Disassociate(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string): JQueryPromise<Soap.SoapResponse> {
            var DisassociateRequest = new Soap.DisassociateRequest(Moniker1, Moniker2, RelationshipName);
            return $.Deferred<Soap.SoapResponse>(function (dfd) {
                var Request = Soap.Execute<Soap.SoapResponse>(DisassociateRequest);
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        private static DoRequest<T>(SoapBody: string, RequestType: string): JQueryPromise<T> {
            var XML = "";

            if (RequestType == "Execute") {
                XML = "<s:Envelope xmlns:s = \"http://schemas.xmlsoap.org/soap/envelope/\">" +
                "<s:Body>" + SoapBody + "</s:Body></s:Envelope>";
            }
            else {
                SoapBody = "<" + RequestType + ">" + SoapBody + "</" + RequestType + ">";
                //Add in all the different namespaces
                XML = "<s:Envelope" + Soap.GetNameSpacesXML() + "><s:Body>" + SoapBody + "</s:Body></s:Envelope>";
            }

            return $.Deferred<T>(function (dfd) {
                var Request = $.ajax(XrmTSToolkit.Common.GetSoapServiceURL(), {
                    data: XML,
                    type: "POST",
                    beforeSend: function (xhr: JQueryXHR) {
                        xhr.setRequestHeader("Accept", "application/xml, text/xml, */*");
                        xhr.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                        xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/" + RequestType);
                    }
                });
                Request.done(function (data, result, xhr: JQueryXHR) {
                    var sr: any = new Soap.SoapResponse(data);
                    dfd.resolve(<T>sr, result, xhr);
                });
                Request.fail(function (req: XMLHttpRequest, description, exception) {
                    var ErrorText = typeof exception === "undefined" ? description : exception;
                    var fault = new Soap.FaultResponse(req.responseXML);
                    if (fault && fault.Fault && fault.Fault.FaultString) {
                        ErrorText += (": " + fault.Fault.FaultString);
                    }
                    dfd.reject({ success: false, result: ErrorText, fault: fault, response: req.responseText });
                });
            }).promise();
        }

        static GetNameSpacesXML(): string {
            var xmlns = " xmlns = \"http://schemas.microsoft.com/xrm/2011/Contracts/Services\"";
            var ns = {
                "s": "http://schemas.xmlsoap.org/soap/envelope/",
                "a": "http://schemas.microsoft.com/xrm/2011/Contracts",
                "i": "http://www.w3.org/2001/XMLSchema-instance",
                "b": "http://schemas.datacontract.org/2004/07/System.Collections.Generic",
                "c": "http://www.w3.org/2001/XMLSchema",
                //"d": "http://schemas.microsoft.com/xrm/2011/Contracts/Services",
                "e": "http://schemas.microsoft.com/2003/10/Serialization/",
                "f": "http://schemas.microsoft.com/2003/10/Serialization/Arrays",
                "g": "http://schemas.microsoft.com/crm/2011/Contracts",
                "h": "http://schemas.microsoft.com/xrm/2011/Metadata",
                "j": "http://schemas.microsoft.com/xrm/2011/Metadata/Query",
                "k": "http://schemas.microsoft.com/xrm/2013/Metadata",
                "l": "http://schemas.microsoft.com/xrm/2012/Contracts"
            };

            var XML = xmlns;
            for (var i in ns) {
                XML += " xmlns:" + i + "=\"" + ns[i] + "\"";
            }
            return XML;
        }
    }
    export module Soap {
        export class Entity {
            constructor(public LogicalName: string = "", public Id: string = null) {
                this.Attributes = new AttributeCollection();
                if (!Id) {
                    this.Id = "00000000-0000-0000-0000-000000000000";
                }
            }
            Attributes: AttributeCollection;
            FormattedValues: StringDisctionary<string>;
            Serialize(): string {
                var Data: string = "<a:Attributes>";
                for (var AttributeName in this.Attributes) {
                    var Attribute = this.Attributes[AttributeName];
                    Data += "<a:KeyValuePairOfstringanyType>";
                    Data += "<b:key>" + AttributeName + "</b:key>";
                    if (Attribute === null || Attribute.IsNull()) {
                        Data += "<b:value i:nil=\"true\" />";
                    }
                    else {
                        Data += Attribute.Serialize();
                    }
                    Data += "</a:KeyValuePairOfstringanyType>";
                }
                Data += "</a:Attributes><a:EntityState i:nil=\"true\" /><a:FormattedValues />";
                Data += "<a:Id>" + Entity.EncodeValue(this.Id) + "</a:Id>";
                Data += "<a:LogicalName>" + this.LogicalName + "</a:LogicalName>";
                Data += "<a:RelatedEntities />";
                return Data;
            }

            static padNumber(s, len: number = 2): any {
                s = '' + s;
                while (s.length < len) {
                    s = "0" + s;
                }
                return s;
            }

            static EncodeValue(value): string {
                return (typeof value === "object" && value.getTime) ? Entity.EncodeDate(value) : ((typeof (<any>window).CrmEncodeDecode != "undefined") ? (<any>window).CrmEncodeDecode.CrmXmlEncode(value) : Entity.CrmXmlEncode(value));
            }

            static EncodeDate(Value: Date): string {
                return Value.getFullYear() + "-" +
                    Entity.padNumber(Value.getMonth() + 1) + "-" +
                    Entity.padNumber(Value.getDate()) + "T" +
                    Entity.padNumber(Value.getHours()) + ":" +
                    Entity.padNumber(Value.getMinutes()) + ":" +
                    Entity.padNumber(Value.getSeconds());
            }

            static CrmXmlEncode(s: string) {
                if ('undefined' === typeof s || 'unknown' === typeof s || null === s) return s;
                else if (typeof s != "string") s = s.toString();
                return Entity.innerSurrogateAmpersandWorkaround(s);
            }

            static HtmlEncode(s: string) {
                if (s === null || s === "" || s === undefined) return s;
                for (var count = 0, buffer = "", hEncode = "", cnt = 0; cnt < s.length; cnt++) {
                    var c = s.charCodeAt(cnt);
                    if (c > 96 && c < 123 || c > 64 && c < 91 || c === 32 || c > 47 && c < 58 || c === 46 || c === 44 || c === 45 || c === 95)
                        buffer += String.fromCharCode(c);
                    else buffer += "&#" + c + ";";
                    if (++count === 500) {
                        hEncode += buffer; buffer = ""; count = 0;
                    }
                }
                if (buffer.length) hEncode += buffer;
                return hEncode;
            }

            static innerSurrogateAmpersandWorkaround(s: string) {
                var buffer = '';
                var c0;
                for (var cnt = 0; cnt < s.length; cnt++) {
                    c0 = s.charCodeAt(cnt);
                    if (c0 >= 55296 && c0 <= 57343)
                        if (cnt + 1 < s.length) {
                            var c1 = s.charCodeAt(cnt + 1);
                            if (c1 >= 56320 && c1 <= 57343) {
                                buffer += "CRMEntityReferenceOpen" + ((c0 - 55296) * 1024 + (c1 & 1023) + 65536).toString(16) + "CRMEntityReferenceClose"; cnt++;
                            }
                            else
                                buffer += String.fromCharCode(c0);
                        }
                        else buffer += String.fromCharCode(c0);
                    else buffer += String.fromCharCode(c0);
                }
                s = buffer;
                buffer = "";
                for (cnt = 0; cnt < s.length; cnt++) {
                    c0 = s.charCodeAt(cnt);
                    if (c0 >= 55296 && c0 <= 57343)
                        buffer += String.fromCharCode(65533);
                    else buffer += String.fromCharCode(c0);
                }
                s = buffer;
                s = Entity.HtmlEncode(s);
                s = s.replace(/CRMEntityReferenceOpen/g, "&#x");
                s = s.replace(/CRMEntityReferenceClose/g, ";");
                return s;
            }
        }
        export enum AttributeType {
            AliasedValue,
            Boolean,
            Date,
            Decimal,
            Double,
            EntityCollection,
            EntityReference,
            Float,
            Guid,
            Integer,
            OptionSetValue,
            Money,
            String,
        }
        export class AttributeCollection {
            [index: string]: AttributeValue;
        }
        export class StringDisctionary<T> {
            [index: string]: T;
        }
        export class EntityCollection {
            Items: Array<Entity> = [];
            Serialize(): string {
                var XML = "";
                $.each(this.Items, function (index, Entity) {
                    XML += "<a:Entity>" + Entity.Serialize() + "</a:Entity>";
                });
                return XML;
            }
        }
        export class ColumnSet {
            constructor(Columns: Array<string>);
            constructor(AllColumns: boolean);
            constructor(p: any) {
                if (typeof p === "boolean") {
                    this.AllColumns = p;
                    this.Columns = [];
                }
                else {
                    this.Columns = p;
                }
            }

            AllColumns: boolean = false;
            Columns: Array<string>;

            AddColumn(Column: string): void { this.Columns.push(Column); }
            AddColumns(Columns: Array<string>): void {
                for (var Column in Columns) {
                    this.AddColumn(Column);
                }
            }

            serialize(): string {
                var Data: string = "<a:AllColumns>" + this.AllColumns.toString() + "</a:AllColumns>";
                if (this.Columns.length == 0) {
                    Data += "<a:Columns />";
                }
                else {
                    for (var Column in this.Columns) {
                        Data += "<f:string>" + Column + "</f:string>";
                    }
                }
                return Data;
            }
        }
        export class AttributeValue {
            constructor(public Value: any, public Type: AttributeType) { }
            IsNull(): boolean { return this.Value === undefined; }
            FormattedValue: string;
            GetEncodedValue(): string {
                return Entity.EncodeValue(this.Value);
            }
            Serialize(): string { return "<b:value i:type=\"" + this.GetValueType() + "\">" + this.GetEncodedValue() + "</b:value>"; }
            GetValueType(): string { return ""; }
        }
        export class EntityCollectionAttribute extends AttributeValue {
            constructor(public Value: EntityCollection) { super(Value, AttributeType.EntityCollection); }
            GetValueType(): string { return "a:ArrayOfEntity"; }
        }

        export class EntityReference extends AttributeValue {
            constructor(public Id?: string, public LogicalName?: string, public Name?: string) { super(null, AttributeType.EntityReference); }
            Serialize(): string {
                return "<b:value i:type=\"a:EntityReference\"><a:Id>" +
                    Entity.EncodeValue(this.Id) +
                    "</a:Id><a:LogicalName>" +
                    Entity.EncodeValue(this.LogicalName) +
                    "</a:LogicalName><a:Name i:nil=\"true\" /></b:value>";
            }
            GetValueType(): string { return "a:EntityReference"; }
        }
        export class OptionSetValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.OptionSetValue); }
            Serialize(): string {
                return "<b:value i:type=\"a:OptionSetValue\"><a:Value>" + this.GetEncodedValue() + "</a:Value></b:value>";
            }
            GetValueType(): string { return "a:OptionSetValue"; }
        }
        export class MoneyValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Money); }
            Serialize(): string {
                return "<b:value i:type=\"a:Money\"><a:Value>" + this.GetEncodedValue() + "</a:Value></b:value>";
            }
            GetValueType(): string { return "a:Money"; }
        }
        export class AliasedValue extends AttributeValue {
            constructor(public Value?: AttributeValue, public AttributeLogicalName?: string, public EntityLogicalName?: string) { super(Value, AttributeType.AliasedValue); }
            Serialize(): string { throw "Update the 'Serialize' method of the 'AliasedValue'"; }
            GetValueType(): string { throw "Update the 'GetValueType' method of the 'AliasedValue'"; }
        }
        export class BooleanValue extends AttributeValue {
            constructor(public Value?: boolean) { super(Value, AttributeType.Boolean); }
            GetValueType(): string { return "c:boolean"; }
        }
        export class IntegerValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Integer); }
            GetValueType(): string { return "c:int"; }
        }
        export class StringValue extends AttributeValue {
            constructor(public Value?: string) { super(Value, AttributeType.String); }
            GetValueType(): string { return "c:string"; }
        }
        export class DoubleValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Double); }
            GetValueType(): string { return "c:double"; }
        }
        export class DecimalValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Decimal); }
            GetValueType(): string { return "c:decimal"; }
        }
        export class FloatValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Float); }
            GetValueType(): string { return "c:double"; }
        }
        export class DateValue extends AttributeValue {
            constructor(public Value?: Date) { super(Value, AttributeType.Date); }
            GetEncodedValue(): string {
                return Entity.EncodeDate(this.Value);
            }
            GetValueType(): string { return "c:dateTime"; }
        }
        export class GuidValue extends AttributeValue {
            constructor(public Value?: string) { super(Value, AttributeType.Guid); }
            GetValueType(): string { return "e:guid"; }
        }
        export class NodeValue {
            constructor(public Name: string, public Value: any) { }
        }
        export class Collection {
            constructor(public Items?: any) { }
        }

        export class SoapResponse {
            constructor(XHR: JQueryXHR) {
                XML: XHR;
                SoapResponse.ParseResult(XHR, this);
            }
            static ParseResult(XHR: JQueryXHR, ParentObject: SoapResponse): void {
                var XMLDocTemp = new XMLDocument(XHR);
                var XRMMainNode = XMLDocTemp.Root;
                while (XRMMainNode && XRMMainNode.ChildNodes && XRMMainNode.ChildNodes[0] && (XRMMainNode.Name == "Envelope" || XRMMainNode.Name == "Body" || XRMMainNode.Name == "ExecuteResponse")) {
                    XRMMainNode = XRMMainNode.ChildNodes[0];
                }
                XRMMainNode = XRMMainNode.ChildNodes[0];
                if (XRMMainNode) {
                    var XRMBaseObject = new XRMObject(XRMMainNode);
                    SoapResponse.ParseXRMBaseObject(ParentObject, XRMBaseObject);
                }
            }
            static ParseXRMBaseObject(ParentObject: any, XRMBaseObject: XRMObject): void {
                var CurrentObject = SoapResponse.GetObjectFromBaseXRMObject(XRMBaseObject);
                if (CurrentObject) {
                    if (ParentObject instanceof Array) {
                        ParentObject.push(CurrentObject);
                    }
                    else {
                        ParentObject[XRMBaseObject.Name] = CurrentObject;
                    }
                    if (XRMBaseObject.Value instanceof Array) {
                        $.each(XRMBaseObject.Value, function (index, Child) {
                            SoapResponse.ParseXRMBaseObject(CurrentObject, Child);
                        });
                    }
                    else if (XRMBaseObject.Value && !(typeof XRMBaseObject.Value === "string")) {
                        SoapResponse.ParseXRMBaseObject(CurrentObject, XRMBaseObject.Value);
                    }

                    //Populate the 'FormattedValues' of each of the attributes
                    if (CurrentObject instanceof Soap.Entity && (<Soap.Entity>CurrentObject).FormattedValues) {
                        for (var AttributeName in (<Soap.Entity>CurrentObject).FormattedValues) {
                            (<Soap.Entity>CurrentObject).Attributes[AttributeName].FormattedValue = (<Soap.Entity>CurrentObject).FormattedValues[AttributeName];
                        }
                    }
                }
            }
            static GetObjectFromBaseXRMObject(XRMBaseObject: XRMObject): any {
                var CurrentObject = null;
                switch (XRMBaseObject.Type) {
                    case "Object":
                    case "Array":
                        //Could be an object or Array
                        switch (XRMBaseObject.Name) {
                            case "Entities":
                            case "EntityCollection":
                                CurrentObject = new Array<Entity>();
                                break;
                            case "Attributes":
                                CurrentObject = new AttributeCollection();
                                break;
                            case "EntityReferenceCollection":
                                CurrentObject = [];
                                break;
                            case "Entity":
                                CurrentObject = new Soap.Entity();
                                break;
                            default:
                                if (typeof XRMBaseObject.Value === "string") {
                                    CurrentObject = XRMBaseObject.Value;
                                }
                                else {
                                    CurrentObject = {};
                                }
                        }
                        break;
                    case "EntityReference":
                        CurrentObject = new Soap.EntityReference(
                            (<Array<XRMObject>>XRMBaseObject.Value)[0].Value,
                            (<Array<XRMObject>>XRMBaseObject.Value)[1].Value,
                            (<Array<XRMObject>>XRMBaseObject.Value)[2].Value
                            );
                        XRMBaseObject.Value = null;  //Prevents the parsing of the child values
                        break;
                    case "OptionSetValue":
                        CurrentObject = new Soap.OptionSetValue(parseFloat((<XRMObject>XRMBaseObject.Value).Value));
                        XRMBaseObject.Value = null;  //Prevents the parsing of the child values
                        break;
                    case "Money":
                        CurrentObject = new Soap.MoneyValue(parseFloat((<XRMObject>XRMBaseObject.Value).Value));
                        XRMBaseObject.Value = null;  //Prevents the parsing of the child values
                        break;
                    case "AliasedValue":
                        CurrentObject = new Soap.AliasedValue();
                        break;
                    case "int":
                        CurrentObject = new Soap.IntegerValue(parseInt(XRMBaseObject.Value));
                        break;
                    case "double":
                        CurrentObject = new Soap.DoubleValue(parseFloat(XRMBaseObject.Value));
                        break;
                    case "float":
                        CurrentObject = new Soap.FloatValue(parseFloat(XRMBaseObject.Value));
                        break;
                    case "decimal":
                        CurrentObject = new Soap.DecimalValue(parseFloat(XRMBaseObject.Value));
                        break;
                    case "dateTime":
                        //Parse the date value
                        if (XRMBaseObject.Value) {
                            XRMBaseObject.Value = XRMBaseObject.Value.replaceAll("T", " ").replaceAll("-", "/");
                        }
                        CurrentObject = new Soap.DateValue(new Date(XRMBaseObject.Value));
                        break;
                    case "boolean":
                        CurrentObject = new Soap.BooleanValue(XRMBaseObject.Value == "false" ? false : true);
                        break;
                    case "string":
                        CurrentObject = new Soap.StringValue(XRMBaseObject.Value);
                        break;
                    case "guid":
                        CurrentObject = new Soap.GuidValue(XRMBaseObject.Value);
                        break;
                    default:
                        throw "Please Update the 'ParseXRMBaseObject' for type'" + XRMBaseObject.Type;
                }
                return CurrentObject;
            }
            static RemoveNameSpace(Name: string): string {
                var Return = Name;
                if (Return) {
                    var IndexOfNamespace = Return.indexOf(":");
                    if (IndexOfNamespace >= 0) {
                        Return = Return.substring(IndexOfNamespace + 1);
                    }
                }
                return Return;
            }
        }

        class XMLDocument {
            constructor(XHR: JQueryXHR) {
                var EnvelopeElement = XMLNode.GetNextChild($(XHR).find("s\\:Envelope, Envelope"), "element");
                this.Root = new XMLNode(EnvelopeElement);
            }
            Root: XMLNode;
        }
        class XMLNode {
            constructor(XHR: any) {
                this.ChildNodes = new Array<XMLNode>();
                var _This = this;
                this.Name = XHR.baseName;
                if (XHR.childNodes) {
                    $.each(XHR.childNodes, function (index, ChildNode) {
                        var ChildElement = XMLNode.GetNextChild(ChildNode);
                        if (ChildElement) {
                            if (ChildElement.nodeTypeString == "text" && ChildElement.nodeValue) {
                                //This is the value element
                                _This.Value = ChildElement.nodeValue
                            }
                            else if (ChildElement.nodeTypeString == "element") {
                                //This is a chile element
                                _This.ChildNodes.push(new XMLNode(ChildNode));
                            }
                        }
                    });
                }
                this.Type = "Object";
                $.each(XHR.attributes, function (index, Attribute) {
                    if (Attribute.baseName == "type") {
                        _This.Type = SoapResponse.RemoveNameSpace(Attribute.nodeTypedValue);
                        return false;
                    }
                });
            }
            Name: string;
            Value: string;
            Type: string;
            ChildNodes: Array<XMLNode>;

            static GetNextChild(XHR: any, ElementType?: string): any {
                if (ElementType === undefined) {
                    while (XHR && !XHR.nodeTypeString) {
                        XHR = XHR[0];
                    }
                }
                else {
                    while (XHR && !XHR.nodeTypeString && XHR.nodeTypeString != ElementType) {
                        XHR = XHR[0];
                    }
                }
                return XHR;
            }
        }
        class XRMObject {
            constructor(Node: XMLNode) {
                if (Node.Name && Node.Name.indexOf("KeyValuePair") == 0) {
                    //This is a key value pair - so the Name and Value are actually the first and second children
                    this.Name = Node.ChildNodes[0].Value;
                    if (Node.ChildNodes[1].Value) {
                        this.Value = Node.ChildNodes[1].Value;
                        this.Type = Node.ChildNodes[1].Type;
                    }
                    else if (Node.ChildNodes[1].ChildNodes.length == 1) {
                        if (Node.ChildNodes[1].Name === "value" && (Node.ChildNodes[1].ChildNodes[0])) { this.Value = new XRMObject(Node.ChildNodes[1].ChildNodes[0]); }
                        else { this.Value = new XRMObject(Node.ChildNodes[1]); }
                        this.Type = Node.ChildNodes[1].Type;
                    }
                    else {
                        this.Value = [];
                        var _Value = this.Value;
                        this.Type = Node.ChildNodes[1].Type != "Object" ? Node.ChildNodes[1].Type : "Array";
                        $.each(Node.ChildNodes[1].ChildNodes, function (index, ChildNode) {
                            _Value.push(new XRMObject(ChildNode));
                        });
                    }
                }
                else {
                    //Determine the type of object by examining the children
                    this.Name = Node.Name;
                    if (Node.Value || Node.ChildNodes.length == 0) {
                        this.Value = Node.Value;
                        this.Type = Node.Type;
                    }
                    else if (Node.ChildNodes.length == 1) {
                        this.Value = new XRMObject(Node.ChildNodes[0]);
                        this.Type = "Object";
                    }
                    else {
                        this.Value = [];
                        var _Value = this.Value;
                        this.Type = Node.Type != "Object" ? Node.Type : "Array";
                        $.each(Node.ChildNodes, function (index, ChildNode) {
                            _Value.push(new XRMObject(ChildNode));
                        });
                    }
                }
            }
            Name: string;
            Value: any;
            Type: string;
        }

        export class SDKResponse { }
        export class CreateSoapResponse extends SDKResponse {
            CreateResult: string;
        }
        export class UpdateSoapResponse extends SDKResponse { }
        export class DeleteSoapResponse extends SDKResponse { }
        export class RetrieveSoapResponse extends SDKResponse {
            RetrieveResponse: any;
        }
        export class RetrieveMultipleSoapResponse extends SDKResponse {
            RetrieveMultipleResponse: any;
            RetrieveMultipleResult: RetrieveMultipleResult;
        }
        export class RetrieveMultipleResult {
            Entities: Array<Entity>;
            EntityName: string;
            MoreRecords: boolean = false;
            PagingCookie: string;
            TotalRecordCount: number;
            TotalRecordCountLimitExceeded: boolean;
        }
        export class FaultResponse extends SoapResponse {
            Fault: Fault;
        }
        export class Fault {
            FaultCode: string;
            FaultString: string;
        }
        export class ExecuteRequest {
            constructor(public RequestName: string, public Parameters?: AttributeCollection) {
                if (!Parameters || Parameters == null) {
                    this.Parameters = new AttributeCollection();
                }
            }
            RequestId: string;
            Serialize(): string {
                var XML = "<Execute " + Soap.GetNameSpacesXML() + ">";
                XML += "<request i:type='g:" + this.RequestName + "Request'><a:Parameters>";
                for (var ParameterName in this.Parameters) {
                    var Parameter = this.Parameters[ParameterName];
                    XML += "<a:KeyValuePairOfstringanyType>";
                    XML += "<b:key>" + ParameterName + "</b:key>";
                    XML += Parameter.Serialize();
                    XML += "</a:KeyValuePairOfstringanyType>";
                }
                XML += "</a:Parameters>";
                XML += "<a:RequestId i:nil = \"true\" />";
                XML += "<a:RequestName>" + this.RequestName + "</a:RequestName>";
                XML += "</request></Execute>";
                return XML;
            }
        }
        export class AssociateRequest extends ExecuteRequest {
            constructor(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string) {
                super("AssociateEntities");
                this.Parameters["Moniker1"] = Moniker1;
                this.Parameters["Moniker2"] = Moniker2;
                this.Parameters["RelationshipName"] = new Soap.StringValue(RelationshipName);
            }
        }
        export class DisassociateRequest extends ExecuteRequest {
            constructor(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string) {
                super("DisassociateEntities");
                this.Parameters["Moniker1"] = Moniker1;
                this.Parameters["Moniker2"] = Moniker2;
                this.Parameters["RelationshipName"] = new Soap.StringValue(RelationshipName);
            }
        }
    }

    export module Soap.Query {
        export class QueryExpression {
            constructor(public EntityName?: string) {
            }

            Columns: ColumnSet = new ColumnSet(false);
            Criteria: FilterExpression = null;
            PageInfo: PageInfo = new PageInfo();
            Orders: Array<OrderExpression> = [];
            LinkEntities: Array<LinkEntity> = [];
            NoLock: boolean = false;

            serialize(): string {
                //Query
                var Data: string = "<query i:type=\"a:QueryExpression\">";

                //Columnset
                Data += "<a:ColumnSet>" + this.Columns.serialize() + "</a:ColumnSet>";

                //Criteria - Serailize the FilterExpression
                if (this.Criteria == null) {
                    Data += "<a:Criteria i:nil=\"true\"/>";
                }
                else {
                    Data += "<a:Criteria>";
                    Data += this.Criteria.serialize();
                    Data += "</a:Criteria>";
                }

                Data += "<a:Distinct>false</a:Distinct>";
                Data += "<a:EntityName>" + this.EntityName + "</a:EntityName>";

                //Link Entities
                if (this.LinkEntities.length == 0) {
                    Data += "<a:LinkEntities />";
                }
                else {
                    Data += "<a:LinkEntities>";
                    $.each(this.LinkEntities, function (index, LinkEntity) {
                        Data += LinkEntity.serialize();
                    });
                    Data += "</a:LinkEntities>";
                }
                   
                //Sorting
                if (this.Orders.length == 0) {
                    Data += "<a:Orders />";
                }
                else {
                    Data += "<a:Orders>";
                    $.each(this.Orders, function (index, Order) {
                        Data += Order.serialize();
                    });
                    Data += "</a:Orders>";
                }

                //Page Info
                Data += this.PageInfo.serialize();

                //No Lock
                Data += "<a:NoLock>" + this.NoLock.toString() + "</a:NoLock>";

                Data += "</query>";
                return Data;
            }
        }
        export class FilterExpression {
            constructor();
            constructor(FilterOperator: LogicalOperator);
            constructor(FilterOperator?: LogicalOperator) {
                if (FilterOperator === undefined) {
                    this.FilterOperator = LogicalOperator.And;
                }
            }
            FilterOperator: LogicalOperator = LogicalOperator.And;
            Conditions: Array<ConditionExpression> = [];
            Filters: Array<FilterExpression> = [];

            AddFilter(p: LogicalOperator): void;
            AddFilter(p: FilterExpression): void;
            AddFilter(p: any): void {
                if (typeof p === "number") // enum
                    p = new FilterExpression(p);
                this.Filters.push(p);
            }

            AddCondition(p: ConditionExpression): void;
            AddCondition(p: String, Operator: ConditionOperator, Values: any): ConditionExpression;
            AddCondition(p: any, Operator?: ConditionOperator, Values?: any): ConditionExpression {
                if (typeof p === "string") {
                    p = new ConditionExpression(p, Operator, Values);
                }
                this.Conditions.push(p);
                return p;
            }

            IsQuickFindFilter: boolean = false;

            serialize(): string {
                var Data: string = "<a:Conditions>";
                //Conditions
                $.each(this.Conditions, function (i, Condition) {
                    Data += Condition.serialize();
                });
                Data += "</a:Conditions>";

                //Filter Operator
                Data += "<a:FilterOperator>" + LogicalOperator[this.FilterOperator].toString() + "</a:FilterOperator>";

                //Filters
                if (this.Filters.length == 0) {
                    Data += "<a:Filters/>";
                }
                else {
                    Data += "<a:Filters>";
                    $.each(this.Filters, function (i, Filter) {
                        Data += "<a:FilterExpression>"
                        Data += Filter.serialize();
                        Data += "</a:FilterExpression>";
                    });
                    Data += "</a:Filters>";
                }

                //IsQuickFindFilter
                Data += "<a:IsQuickFindFilter>" + this.IsQuickFindFilter.toString() + "</a:IsQuickFindFilter>";
                return Data;
            }
        }

        export class ConditionExpression {
            constructor();
            constructor(AttributeName: string, Operator: ConditionOperator, Values?: AttributeValue);
            constructor(AttributeName: string, Operator: ConditionOperator, Values?: Array<AttributeValue>);
            constructor(public AttributeName?: string, public Operator?: ConditionOperator, Values?: any) {
                if (Values instanceof Array) {
                    this.Values = Values;
                }
                else {
                    this.Values.push(Values);
                }
            }

            Values: Array<AttributeValue> = [];

            serialize(): string {
                var Data: string =
                    "<a:ConditionExpression>" +
                    "<a:AttributeName>" + this.AttributeName + "</a:AttributeName>" +
                    "<a:Operator>" + ConditionOperator[this.Operator].toString() + "</a:Operator>" +
                    "<a:Values>";
                $.each(this.Values, function (i, Value) {
                    Data += "<f:anyType i:type=\"" + Value.GetValueType() + "\">" + Value.GetEncodedValue() + "</f:anyType>";
                });
                Data += "</a:Values>" +
                "</a:ConditionExpression>";
                return Data;
            }
        }

        export class PageInfo {
            constructor() { }
            Count: Number = 0;
            PageNumber: Number = 0;
            PagingCookie: string = null;
            ReturnTotalRecordCount: boolean = false;

            serialize(): string {
                var Data: string = "" +
                    "<a:PageInfo>" +
                    "<a:Count>" + this.Count.toString() + "</a:Count>" +
                    "<a:PageNumber>" + this.PageNumber.toString() + "</a:PageNumber>";
                if (this.PagingCookie == null) {
                    Data += "<a:PagingCookie i:nil = \"true\" />";
                }
                else {
                    Data += "<a:PagingCookie>" + this.PagingCookie + "</a:PagingCookie>";
                }
                Data += "<a:ReturnTotalRecordCount>" + this.ReturnTotalRecordCount.toString() + "</a:ReturnTotalRecordCount></a:PageInfo>";
                return Data;
            }
        }

        export class OrderExpression {
            constructor();
            constructor(AttributeName: string, OrderType: OrderType);
            constructor(public AttributeName?: string, public OrderType?: OrderType) { }
            serialize(): string {
                var Data: string = "<a:OrderExpression>";
                Data += "<a:AttributeName>" + this.AttributeName + "</a:AttributeName>";
                Data += "<a:OrderType>" + OrderType[this.OrderType].toString() + "</a:OrderType>";
                Data += "</a:OrderExpression>";
                return Data;
            }
        }

        export class LinkEntity {
            constructor(
                public LinkFromEntityName: string,
                public LinkToEntityName: string,
                public LinkFromAttributeName: string,
                public LinkToAttributeName: string,
                public JoinOperator: JoinOperator = Soap.Query.JoinOperator.Inner) { }

            Columns: ColumnSet = new ColumnSet(false);
            EntityAlias: string = null;
            LinkCriteria: FilterExpression = new FilterExpression(Soap.Query.LogicalOperator.And);
            LinkEntities: Array<LinkEntity> = [];

            serialize(): string {
                var Data: string = "<a:LinkEntity>";
                Data += this.Columns.serialize();
                if (this.EntityAlias == null) {
                    Data += "<a:EntityAlias i:nil=\"true\"/>";
                }
                else {
                    Data += "<a:EntityAlias>" + this.EntityAlias + "</a:EntityAlias>";
                }
                Data += "<a:JoinOperator>" + JoinOperator[this.JoinOperator].toString() + "</a:JoinOperator>";
                Data += " <a:LinkCriteria>";
                Data += this.LinkCriteria.serialize();
                Data += "</a:LinkCriteria>";
                if (this.LinkEntities.length == 0) {
                    Data += "<a:LinkEntities />";
                }
                else {
                    Data += "<a:LinkEntities>";
                    $.each(this.LinkEntities, function (index, LinkEntity) {
                        Data += LinkEntity.serialize();
                    });
                    Data += "</a:LinkEntities>";
                }
                Data += "<a:LinkFromAttributeName>" + this.LinkFromAttributeName + "</a:LinkFromAttributeName>";
                Data += "<a:LinkFromEntityName>" + this.LinkFromEntityName + "</a:LinkFromEntityName>";
                Data += "<a:LinkToAttributeName>" + this.LinkToAttributeName + "</a:LinkToAttributeName>";
                Data += "<a:LinkToEntityName>" + this.LinkToEntityName + "</a:LinkToEntityName>";
                Data += "</a:LinkEntity>";
                return Data;
            }
        }

        export enum JoinOperator {
            Inner = 0,
            LeftOuter = 1,
            Natural = 2
        }

        export enum OrderType {
            Ascending = 0,
            Descending = 1
        }

        export enum LogicalOperator {
            And = 0,
            Or = 1
        }

        export enum ConditionOperator {
            Equal = 0,
            NotEqual = 1,
            GreaterThan = 2,
            LessThan = 3,
            GreaterEqual = 4,
            LessEqual = 5,
            Like = 6,
            NotLike = 7,
            In = 8,
            NotIn = 9,
            Between = 10,
            NotBetween = 11,
            Null = 12,
            NotNull = 13,
            Yesterday = 14,
            Today = 15,
            Tomorrow = 16,
            Last7Days = 17,
            Next7Days = 18,
            LastWeek = 19,
            ThisWeek = 20,
            NextWeek = 21,
            LastMonth = 22,
            ThisMonth = 23,
            NextMonth = 24,
            On = 25,
            OnOrBefore = 26,
            OnOrAfter = 27,
            LastYear = 28,
            ThisYear = 29,
            NextYear = 30,
            LastXHours = 31,
            NextXHours = 32,
            LastXDays = 33,
            NextXDays = 34,
            LastXWeeks = 35,
            NextXWeeks = 36,
            LastXMonths = 37,
            NextXMonths = 38,
            LastXYears = 39,
            NextXYears = 40,
            EqualUserId = 41,
            NotEqualUserId = 42,
            EqualBusinessId = 43,
            NotEqualBusinessId = 44,
            ChildOf = 45,
            Mask = 46,
            NotMask = 47,
            MasksSelect = 48,
            Contains = 49,
            DoesNotContain = 50,
            EqualUserLanguage = 51,
            NotOn = 52,
            OlderThanXMonths = 53,
            BeginsWith = 54,
            DoesNotBeginWith = 55,
            EndsWith = 56,
            DoesNotEndWith = 57,
            ThisFiscalYear = 58,
            ThisFiscalPeriod = 59,
            NextFiscalYear = 60,
            NextFiscalPeriod = 61,
            LastFiscalYear = 62,
            LastFiscalPeriod = 63,
            LastXFiscalYears = 64,
            LastXFiscalPeriods = 65,
            NextXFiscalYears = 66,
            NextXFiscalPeriods = 67,
            InFiscalYear = 68,
            InFiscalPeriod = 69,
            InFiscalPeriodAndYear = 70,
            InOrBeforeFiscalPeriodAndYear = 71,
            InOrAfterFiscalPeriodAndYear = 72,
            EqualUserTeams = 73
        }
    }
}

interface String {
    replaceAll(search: string, replace: string): string;
}

String.prototype.replaceAll = function (search, replace): string {
    //if replace is null, return original string otherwise it will
    //replace search string with 'undefined'.
    if (!replace)
        return this;
    return this.replace(new RegExp('[' + search + ']', 'g'), replace);
}