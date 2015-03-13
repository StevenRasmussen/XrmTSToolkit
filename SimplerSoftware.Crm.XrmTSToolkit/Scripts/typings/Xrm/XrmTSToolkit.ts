/// <reference path="../jquery/jquery.d.ts" />
/// <reference path="xrm-2015.d.ts" />

/**
* MSCRM 2011, 2013, 2015 Service Toolkit for TypeScript
* @author Steven Rasmusse
* @current version : 0.5.2
* Credits:
*   The idea of this library was inspired by David Berry and Jaime Ji's XrmServiceToolkit.js
*    
* Date: March 12, 2015
* 
* Required TypeScript Version: 1.4
* 
* Required external libraries:
*   jquery.d.ts - Downloaded from Nuget and included in the project
*   Xrm-2015.d.ts - Downloaded from Nuget and included in the project
* 
* Version : 0.5
* Date: March 12, 2015
*   Initial Beta Release
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
                var Request = Soap.DoRequest(request, "Create");
                Request.done(function (data: Soap.SoapResponse, result, xhr) {
                    dfd.resolve(<Soap.CreateSoapResponse> data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Update(Entity: Soap.Entity): JQueryPromise<Soap.UpdateSoapResponse> {
            var request = "<entity>" + Entity.Serialize() + "</entity>";
            return $.Deferred<Soap.UpdateSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest(request, "Update");
                Request.done(function (data: Soap.SoapResponse, result, xhr) {
                    dfd.resolve(<Soap.UpdateSoapResponse> data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static Retrieve(Id: string, EntityLogicalName: string, ColumnSet?: Soap.ColumnSet): JQueryPromise<Soap.RetrieveSoapResponse> {
            Id = Common.StripGUID(Id);
            var msgBody = "<entityName>" + EntityLogicalName + "</entityName><id>" + Id + "</id>" + ColumnSet.serialize();
            return $.Deferred<Soap.RetrieveSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest(msgBody, "Retrieve");
                Request.done(function (data: Soap.SoapResponse, result, xhr) {
                    dfd.resolve(<Soap.RetrieveSoapResponse>data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        static RetrieveMultiple(Query: Soap.Query.QueryExpression): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest(Query.serialize(), "RetrieveMultiple");
                Request.done(function (data: Soap.SoapResponse, result, xhr) {
                    dfd.resolve(<Soap.RetrieveMultipleSoapResponse>data.SDKResponse);
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

            var msgBody = "<query i:type='b:FetchExpression'><b:Query>" + fetchXml + "</b:Query></query>";
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest(msgBody, "RetrieveMultiple");
                Request.done(function (data: Soap.SoapResponse, result, xhr) {
                    dfd.resolve(<Soap.RetrieveMultipleSoapResponse>data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
            The 'ExecuteBody' should be the contents of the '<Execute>' tags,  INCLUDING the '<Execute>' tags
        */
        static Execute<T>(ExecuteBody: string): JQueryPromise<Soap.SoapResponse> {
            return $.Deferred<Soap.SoapResponse>(function (dfd) {
                var Request = Soap.DoRequest(ExecuteBody, "Execute");
                Request.done(function (data: Soap.SoapResponse, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        private static DoRequest(SoapBody: string, RequestType: string): JQueryPromise<Soap.SoapResponse> {
            var soapXml = "";
            if (RequestType == "Execute") {
                soapXml = "<soap:Envelope xmlns:soap = \"http://schemas.xmlsoap.org/soap/envelope/\">" +
                "<soap:Body>" + SoapBody + "</soap:Body></soap:Envelope>";
            }
            else {
                //This is a native request that does not contain any namespaces yet
                var xmlns = " xmlns = \"http://schemas.microsoft.com/xrm/2011/Contracts/Services\"";
                var xmlnsb = " xmlns:b = \"http://schemas.microsoft.com/xrm/2011/Contracts\"";
                var xmlnsc = " xmlns:c = \"http://schemas.microsoft.com/2003/10/Serialization/Arrays\"";
                var xmlnsd = " xmlns:d = \"http://www.w3.org/2001/XMLSchema\"";
                var xmlnse = " xmlns:e = \"http://schemas.microsoft.com/2003/10/Serialization/\"";
                var xmlnsg = " xmlns:g = \"http://schemas.datacontract.org/2004/07/System.Collections.Generic\"";
                var xmlnsi = " xmlns:i = \"http://www.w3.org/2001/XMLSchema-instance\"";
                var xmlnssoap = " xmlns:soap = \"http://schemas.xmlsoap.org/soap/envelope/\"";
                SoapBody = "<" + RequestType + ">" + SoapBody + "</" + RequestType + ">";

                //Add in all the different namespaces
                soapXml = "<soap:Envelope" + xmlns + xmlnsb + xmlnsc + xmlnsd + xmlnse + xmlnsi + xmlnsg + xmlnssoap + ">" +
                "<soap:Body>" + SoapBody + "</soap:Body></soap:Envelope>";
            }

            return $.Deferred<Soap.SoapResponse>(function (dfd) {
                var Request = $.ajax(XrmTSToolkit.Common.GetSoapServiceURL(), {
                    data: soapXml,
                    type: "POST",
                    beforeSend: function (xhr: JQueryXHR) {
                        xhr.setRequestHeader("Accept", "application/xml, text/xml, */*");
                        xhr.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                        xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/" + RequestType);
                    }
                });
                Request.done(function (data, result, xhr: JQueryXHR) {
                    var sr = new Soap.SoapResponse(data);
                    dfd.resolve(sr, result, xhr);
                });
                Request.fail(function (req: XMLHttpRequest, description, exception) {
                    var ErrorText = typeof exception === "undefined" ? description : exception;
                    var fault = new Soap.FaultResponse(req.responseXML);
                    if (fault && fault.Fault && fault.Fault.FaultString) {
                        ErrorText += (": " + fault.Fault.FaultString);
                    }
                    dfd.reject({ success: false, result: ErrorText });
                });
            }).promise();
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
            Serialize(): string {
                var Data: string = "<b:Attributes>";
                for (var AttributeName in this.Attributes) {
                    var Attribute = this.Attributes[AttributeName];
                    Data += "<b:KeyValuePairOfstringanyType>";
                    Data += "<g:key>" + AttributeName + "</g:key>";
                    if (Attribute === null || Attribute.IsNull()) {
                        Data += "<g:value i:nil=\"true\" />";
                    }
                    else {
                        Data += Attribute.Serialize();
                    }
                    Data += "</b:KeyValuePairOfstringanyType>";
                }
                Data += "</b:Attributes><b:EntityState i:nil=\"true\" /><b:FormattedValues />";
                Data += "<b:Id>" + Entity.EncodeValue(this.Id) + "</b:Id>";
                Data += "<b:LogicalName>" + this.LogicalName + "</b:LogicalName>";
                Data += "<b:RelatedEntities />";
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

        export class EntityCollection {
            Items: Array<Entity> = [];
            Serialize(): string {
                var XML = "";
                this.Items.forEach(Entity => {
                    XML += "<b:Entity>" + Entity.Serialize() + "</b:Entity>";
                });
                return XML;
            }
        }

        export class AttributeValue {
            constructor(public Value: any, public Type: AttributeType) { }
            IsNull(): boolean { return this.Value === undefined; }
            FormattedValue: string;
            GetEncodedValue(): string {
                return Entity.EncodeValue(this.Value);
            }
            Serialize(): string { return "<g:value i:type=\"" + this.GetValueType() + "\">" + this.GetEncodedValue() + "</g:value>";; }
            GetValueType(): string { return ""; }
        }
        export class EntityCollectionAttribute extends AttributeValue {
            constructor(public Value: EntityCollection) { super(Value, AttributeType.EntityCollection); }
            GetValueType(): string { return "b:ArrayOfEntity"; }
        }
        export class EntityReference extends AttributeValue {
            constructor(public Id?: string, public LogicalName?: string, public Name?: string) { super({ Id: Id, LogicalName: LogicalName, Name: Name }, AttributeType.EntityReference); }
            Serialize(): string {
                return "<g:value i:type=\"b:EntityReference\"><b:Id>" +
                    Entity.EncodeValue(this.Id) +
                    "</b:Id><b:LogicalName>" +
                    Entity.EncodeValue(this.LogicalName) +
                    "</b:LogicalName><b:Name i:nil=\"true\" /></g:value>";
            }
            GetValueType(): string { return "b:EntityReference"; }
        }
        export class OptionSetValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.OptionSetValue); }
            Serialize(): string {
                return "<g:value i:type=\"b:OptionSetValue\"><b:Value>" + this.GetEncodedValue() + "</b:Value></g:value>";
            }
            GetValueType(): string { return "b:OptionSetValue"; }
        }
        export class MoneyValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Money); }
            Serialize(): string {
                return "<g:value i:type=\"b:Money\"><b:Value>" + this.GetEncodedValue() + "</b:Value></g:value>";
            }
            GetValueType(): string { return "b:Money"; }
        }
        export class AliasedValue extends AttributeValue {
            constructor(public Value?: AttributeValue, public AttributeLogicalName?: string, public EntityLogicalName?: string) { super(Value, AttributeType.AliasedValue); }
            Serialize(): string { throw "Update the 'Serialize' method of the 'AliasedValue'"; }
            GetValueType(): string { throw "Update the 'GetValueType' method of the 'AliasedValue'"; }
        }
        export class BooleanValue extends AttributeValue {
            constructor(public Value?: boolean) { super(Value, AttributeType.Boolean); }
            GetValueType(): string { return "d:boolean"; }
        }
        export class IntegerValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Integer); }
            GetValueType(): string { return "d:int"; }
        }
        export class StringValue extends AttributeValue {
            constructor(public Value?: string) { super(Value, AttributeType.String); }
            GetValueType(): string { return "d:string"; }
        }
        export class DoubleValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Double); }
            GetValueType(): string { return "d:double"; }
        }
        export class DecimalValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Decimal); }
            GetValueType(): string { return "d:decimal"; }
        }
        export class FloatValue extends AttributeValue {
            constructor(public Value?: number) { super(Value, AttributeType.Float); }
            GetValueType(): string { return "d:float"; }
        }
        export class DateValue extends AttributeValue {
            constructor(public Value?: Date) { super(Value, AttributeType.Date); }
            GetEncodedValue(): string {
                return Entity.EncodeDate(this.Value);
            }
            GetValueType(): string { return "d:dateTime"; }
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
            constructor(ResponseXML: JQueryXHR) {
                XML: ResponseXML;
                this.ParseResult(ResponseXML);
            }
            SDKResponse: SDKResponse;
            ParseResult(ResponseXML: JQueryXHR) {
                var BodyNode = $(ResponseXML).find("s\\:Body, Body");
                var val = this.GetNodeValue($(BodyNode[0].firstChild));
                this[val.Name] = val.Value;
                this.SDKResponse = val.Value;
            }
            GetNodeValue(Node): NodeValue {
                var NodeName: string = RemoveNameSpace(Node[0].nodeName);
                if (NodeName === "value") {
                    NodeName = RemoveNameSpace($(Node).attr('i:type'));
                }
                function RemoveNameSpace(Name) {
                    var Return = Name;
                    if (Return) {
                        var IndexOfNamespace = Return.indexOf(":");
                        if (IndexOfNamespace >= 0) {
                            Return = Return.substring(IndexOfNamespace + 1);
                        }
                    }
                    return Return;
                }

                if (Node[0].childNodes && (Node[0].childNodes.length > 1 || (Node[0].childNodes.length > 0 && Node[0].childNodes[0].nodeName !== "#text"))) {
                    var ChildValues = new Array<NodeValue>();
                    $.each(Node[0].childNodes, function (index, ChildNode) {
                        var ChileNodeValue = SoapResponse.prototype.GetNodeValue($(ChildNode));
                        ChildValues.push(ChileNodeValue);
                    });
                    var IsCollection: boolean = false;
                    $.each(Node[0].attributes, function (i, attrib) {
                        if (attrib.value === "http://schemas.datacontract.org/2004/07/System.Collections.Generic") {
                            IsCollection = true;
                            return true;
                        }
                    });

                    if (NodeName === "Fault") {
                        var Fault = new Soap.Fault();
                        Fault.FaultCode = ChildValues[0].Value;
                        Fault.FaultString = ChildValues[1].Value;
                        return new NodeValue(NodeName, Fault);
                    }
                    else if (NodeName === "RetrieveResult") {
                        var Entity = new Soap.Entity();
                        for (var ChildValue in ChildValues) {
                            if (ChildValue.Name === "FormattedValues") {
                                //Add the formatted values to the attributes as well
                                var Attributes = Entity.Attributes;
                                if (Attributes && ChildValue && ChildValue.Value && ChildValue.Value.Items) {
                                    for (var FormattedValue in ChildValue.Value.Items) {
                                        for (var AttributeName in Attributes) {
                                            if (FormattedValue.Name === AttributeName) {
                                                Attributes[AttributeName].FormattedValue = FormattedValue.Value;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        return new NodeValue(NodeName, Entity);
                    }
                    else if (NodeName.indexOf("KeyValuePairOf") === 0) {
                        var KeyName = ChildValues[0].Value;
                        var KeyValue = ChildValues[1].Value;
                        return new NodeValue(KeyName, KeyValue);
                    }
                    else if (NodeName === "EntityReferenceCollection" || IsCollection) {
                        var coll = new Collection();
                        var items = new Array();
                        $.each(ChildValues, function (Index, Value: NodeValue) {
                            coll[Value.Name] = Value.Value;
                            items.push(Value);
                        });
                        coll.Items = items;
                        return new NodeValue(NodeName, coll);
                    }
                    else if (NodeName === "EntityReference") {
                        var ref = new EntityReference();
                        ref.Id = ChildValues[0].Value
                        ref.LogicalName = ChildValues[1].Value
                        ref.Name = ChildValues[2].Value;
                        return new NodeValue(NodeName, ref);
                    }
                    else if (NodeName === "OptionSetValue") {
                        var opt = new OptionSetValue();
                        if (ChildValues[0].Value) { opt.Value = parseInt(ChildValues[0].Value); }
                        return new NodeValue(NodeName, opt);
                    }
                    else if (NodeName === "Money") {
                        var cur = new MoneyValue();
                        if (ChildValues[0].Value) { cur.Value = parseFloat(ChildValues[0].Value); }
                        return new NodeValue(NodeName, cur);
                    }
                    else if (NodeName === "EntityCollection" || NodeName === "Entities") {
                        var items = new Array();
                        $.each(ChildValues, function (Index, Value: NodeValue) {
                            items.push(Value);
                        });
                        return new NodeValue(NodeName, items);
                    }
                    else if (NodeName === "AliasedValue") {
                        var alias = new AliasedValue();
                        alias.AttributeLogicalName = ChildValues[0].Value;
                        alias.EntityLogicalName = ChildValues[1].Value;
                        alias.Value = ChildValues[2].Value;
                        return new NodeValue(NodeName, alias);
                    }
                    else {
                        var obj = new Object();
                        $.each(ChildValues, function (Index, Value: NodeValue) {
                            obj[Value.Name] = Value.Value;
                        });
                        return new NodeValue(NodeName, obj);
                    }
                }
                else {
                    var Value: string = Node.text() === "" ? (Node.attr("i:nil") === "true" ? null : "") : Node.text();
                    if (NodeName === "int") {
                        var int = new IntegerValue();
                        if (Value) { int.Value = parseInt(Value); }
                        return new NodeValue(NodeName, int);
                    }
                    else if (NodeName === "double") {
                        var double = new DoubleValue();
                        if (Value) { double.Value = parseFloat(Value); }
                        return new NodeValue(NodeName, double);
                    }
                    else if (NodeName === "float") {
                        var float = new FloatValue();
                        if (Value) { float.Value = parseFloat(Value); }
                        return new NodeValue(NodeName, float);
                    }
                    else if (NodeName === "decimal") {
                        var dec = new DecimalValue();
                        if (Value) { dec.Value = parseFloat(Value); }
                        return new NodeValue(NodeName, dec);
                    }
                    else if (NodeName === "dateTime") {
                        var date = new DateValue();
                        if (Value) {
                            Value = Value.replaceAll("T", " ").replaceAll("-", "/");
                            date.Value = new Date(Value);
                        }
                        return new NodeValue(NodeName, date);
                    }
                    else if (NodeName === "boolean") {
                        var boolean = new BooleanValue();
                        if (Value) { boolean.Value = Value == 'false' ? false : true; }
                        return new NodeValue(NodeName, boolean);
                    }
                    else if (NodeName === "string") {
                        var str = new StringValue();
                        if (Value) { str.Value = Value; }
                        return new NodeValue(NodeName, str);
                    }
                    else if (NodeName === "guid") {
                        var guid = new GuidValue();
                        if (Value) { guid.Value = Value; }
                        return new NodeValue(NodeName, guid);
                    }
                    else {
                        return new NodeValue(NodeName, Value);
                    }
                }
            }
        }
        export class SDKResponse {

        }
        export class CreateSoapResponse extends SDKResponse {
            CreateResult: string;
        }
        export class UpdateSoapResponse extends SDKResponse {

        }
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
                var Data: string = "<b:ColumnSet>" +
                    "<b:AllColumns>" + this.AllColumns.toString() + "</b:AllColumns>";
                if (this.Columns.length == 0) {
                    Data += "<b:Columns />";
                }
                else {
                    for (var Column in this.Columns) {
                        Data += "<c:string>" + Column + "</c:string>";
                    }
                }
                Data += "</b:ColumnSet>";
                return Data;
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
                var Data: string = "<query i:type=\"b:QueryExpression\">";

                //Columnset
                Data += this.Columns.serialize();

                //Criteria - Serailize the FilterExpression
                if (this.Criteria == null) {
                    Data += "<b:Criteria i:nil=\"true\"/>";
                }
                else {
                    Data += "<b:Criteria>";
                    Data += this.Criteria.serialize();
                    Data += "</b:Criteria>";
                }

                Data += "<b:Distinct>false</b:Distinct>";
                Data += "<b:EntityName>" + this.EntityName + "</b:EntityName>";

                //Link Entities
                if (this.LinkEntities.length == 0) {
                    Data += "<b:LinkEntities />";
                }
                else {
                    Data += "<b:LinkEntities>";
                    this.LinkEntities.forEach(LinkEntity => {
                        Data += LinkEntity.serialize();
                    });
                    Data += "</b:LinkEntities>";
                }
                   
                //Sorting
                if (this.Orders.length == 0) {
                    Data += "<b:Orders />";
                }
                else {
                    Data += "<b:Orders>";
                    this.Orders.forEach(Order => {
                        Data += Order.serialize();
                    });
                    Data += "</b:Orders>";
                }

                //Page Info
                Data += this.PageInfo.serialize();

                //No Lock
                Data += "<b:NoLock>" + this.NoLock.toString() + "</b:NoLock>";

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
                var Data: string = "<b:Conditions>";
                //Conditions
                this.Conditions.forEach(Condition => {
                    Data += Condition.serialize();
                });
                Data += "</b:Conditions>";

                //Filter Operator
                Data += "<b:FilterOperator>" + LogicalOperator[this.FilterOperator].toString() + "</b:FilterOperator>";

                //Filters
                if (this.Filters.length == 0) {
                    Data += "<b:Filters/>";
                }
                else {
                    Data += "<b:Filters>";
                    this.Filters.forEach(Filter => {
                        Data += "<b:FilterExpression>"
                        Data += Filter.serialize();
                        Data += "</b:FilterExpression>";
                    });
                    Data += "</b:Filters>";
                }

                //IsQuickFindFilter
                Data += "<b:IsQuickFindFilter>" + this.IsQuickFindFilter.toString() + "</b:IsQuickFindFilter>";
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
                    "<b:ConditionExpression>" +
                    "<b:AttributeName>" + this.AttributeName + "</b:AttributeName>" +
                    "<b:Operator>" + ConditionOperator[this.Operator].toString() + "</b:Operator>" +
                    "<b:Values>";
                (this.Values).forEach(Value => {
                    Data += "<c:anyType i:type=\"" + Value.GetValueType() + "\">" + Value.GetEncodedValue() + "</c:anyType>";
                });

                Data += "</b:Values>" +
                "</b:ConditionExpression>";
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
                    "<b:PageInfo>" +
                    "<b:Count>" + this.Count.toString() + "</b:Count>" +
                    "<b:PageNumber>" + this.PageNumber.toString() + "</b:PageNumber>";
                if (this.PagingCookie == null) {
                    Data += "<b:PagingCookie i:nil = \"true\" />";
                }
                else {
                    Data += "<b:PagingCookie>" + this.PagingCookie + "</b:PagingCookie>";
                }
                Data += "<b:ReturnTotalRecordCount>" + this.ReturnTotalRecordCount.toString() + "</b:ReturnTotalRecordCount></b:PageInfo>";
                return Data;
            }
        }

        export class OrderExpression {
            constructor();
            constructor(AttributeName: string, OrderType: OrderType);
            constructor(public AttributeName?: string, public OrderType?: OrderType) { }
            serialize(): string {
                var Data: string = "<b:OrderExpression>";
                Data += "<b:AttributeName>" + this.AttributeName + "</b:AttributeName>";
                Data += "<b:OrderType>" + OrderType[this.OrderType].toString() + "</b:OrderType>";
                Data += "</b:OrderExpression>";
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
                var Data: string = "<b:LinkEntity>";
                Data += this.Columns.serialize();
                if (this.EntityAlias == null) {
                    Data += "<b:EntityAlias i:nil=\"true\"/>";
                }
                else {
                    Data += "<b:EntityAlias>" + this.EntityAlias + "</b:EntityAlias>";
                }
                Data += "<b:JoinOperator>" + JoinOperator[this.JoinOperator].toString() + "</b:JoinOperator>";
                Data += " <b:LinkCriteria>";
                Data += this.LinkCriteria.serialize();
                Data += "</b:LinkCriteria>";
                if (this.LinkEntities.length == 0) {
                    Data += "<b:LinkEntities />";
                }
                else {
                    Data += "<b:LinkEntities>";
                    this.LinkEntities.forEach(LinkEntity => {
                        Data += LinkEntity.serialize();
                    });
                    Data += "</b:LinkEntities>";
                }
                Data += "<b:LinkFromAttributeName>" + this.LinkFromAttributeName + "</b:LinkFromAttributeName>";
                Data += "<b:LinkFromEntityName>" + this.LinkFromEntityName + "</b:LinkFromEntityName>";
                Data += "<b:LinkToAttributeName>" + this.LinkToAttributeName + "</b:LinkToAttributeName>";
                Data += "<b:LinkToEntityName>" + this.LinkToEntityName + "</b:LinkToEntityName>";
                Data += "</b:LinkEntity>";
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