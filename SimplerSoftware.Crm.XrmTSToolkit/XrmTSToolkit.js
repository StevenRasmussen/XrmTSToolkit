/// <reference path="scripts/typings/jquery/jquery.d.ts" />
/// <reference path="scripts/typings/xrm/xrm-2015.d.ts" />
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
/**
* MSCRM 2011, 2013, 2015 Service Toolkit for TypeScript
* @author Steven Rasmusse
* @current version : 0.5
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
var XrmTSToolkit;
(function (XrmTSToolkit) {
    var Common = (function () {
        function Common() {
        }
        Common.GetServerURL = function () {
            var URL = document.location.protocol + "//" + document.location.host;
            var Org = Xrm.Page.context.getOrgUniqueName();
            if (document.location.pathname.toUpperCase().indexOf(Org.toUpperCase()) > -1) {
                URL += "/" + Org;
            }
            return URL;
        };
        Common.GetSoapServiceURL = function () {
            return Common.GetServerURL() + "/XRMServices/2011/Organization.svc/web";
        };
        Common.TestForContext = function () {
            if (typeof Xrm === "undefined") {
                throw new Error("Please make sure to add the ClientGlobalContext.js.aspx file to your web resource.");
            }
        };
        Common.GetURLParameter = function (name) {
            return decodeURI((RegExp(name + '=' + '(.+?)(&|$)').exec(location.search) || [, null])[1]);
        };
        Common.StripGUID = function (GUID) {
            return GUID.replace("{", "").replace("}", "");
        };
        return Common;
    })();
    XrmTSToolkit.Common = Common;
    var Soap = (function () {
        function Soap() {
        }
        Soap.Create = function (Entity) {
            var request = "<entity>" + Entity.Serialize() + "</entity>";
            return $.Deferred(function (dfd) {
                var Request = Soap.DoRequest(request, "Create");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        };
        Soap.Update = function (Entity) {
            var request = "<entity>" + Entity.Serialize() + "</entity>";
            return $.Deferred(function (dfd) {
                var Request = Soap.DoRequest(request, "Update");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        };
        Soap.Retrieve = function (Id, EntityLogicalName, ColumnSet) {
            Id = Common.StripGUID(Id);
            var msgBody = "<entityName>" + EntityLogicalName + "</entityName><id>" + Id + "</id>" + ColumnSet.serialize();
            return $.Deferred(function (dfd) {
                var Request = Soap.DoRequest(msgBody, "Retrieve");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        };
        Soap.RetrieveMultiple = function (Query) {
            return $.Deferred(function (dfd) {
                var Request = Soap.DoRequest(Query.serialize(), "RetrieveMultiple");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        };
        Soap.RetrieveRelatedManyToMany = function (LinkFromEntityName, LinkFromEntityId, LinkToEntityName, IntermediateTableName, Columns, SortOrders) {
            if (Columns === void 0) { Columns = new Soap.ColumnSet(false); }
            if (SortOrders === void 0) { SortOrders = []; }
            var Condition = new Soap.Query.ConditionExpression(LinkFromEntityName + "id", 0 /* Equal */, new Soap.GuidValue(LinkFromEntityId));
            var LinkEntity1 = new Soap.Query.LinkEntity(LinkToEntityName, IntermediateTableName, LinkToEntityName + "id", LinkToEntityName + "id", 0 /* Inner */);
            LinkEntity1.LinkCriteria = new Soap.Query.FilterExpression();
            LinkEntity1.LinkCriteria.AddCondition(Condition);
            var LinkEntity2 = new Soap.Query.LinkEntity(IntermediateTableName, LinkFromEntityName, LinkFromEntityName + "id", LinkFromEntityName + "id", 0 /* Inner */);
            LinkEntity2.LinkCriteria = new Soap.Query.FilterExpression();
            LinkEntity2.LinkCriteria.AddCondition(Condition);
            LinkEntity1.LinkEntities.push(LinkEntity2);
            var Query = new Soap.Query.QueryExpression(LinkToEntityName);
            Query.Columns = Columns;
            Query.LinkEntities.push(LinkEntity1);
            Query.Orders = SortOrders;
            return Soap.RetrieveMultiple(Query);
        };
        Soap.Fetch = function (fetchXml) {
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
            return $.Deferred(function (dfd) {
                var Request = Soap.DoRequest(msgBody, "RetrieveMultiple");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data.SDKResponse);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        };
        /**
            The 'ExecuteBody' should be the contents of the '<Execute>' tags,  INCLUDING the '<Execute>' tags
        */
        Soap.Execute = function (ExecuteBody) {
            return $.Deferred(function (dfd) {
                var Request = Soap.DoRequest(ExecuteBody, "Execute");
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        };
        Soap.DoRequest = function (SoapBody, RequestType) {
            var soapXml = "";
            if (RequestType == "Execute") {
                soapXml = "<soap:Envelope xmlns:soap = \"http://schemas.xmlsoap.org/soap/envelope/\">" + "<soap:Body>" + SoapBody + "</soap:Body></soap:Envelope>";
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
                soapXml = "<soap:Envelope" + xmlns + xmlnsb + xmlnsc + xmlnsd + xmlnse + xmlnsi + xmlnsg + xmlnssoap + ">" + "<soap:Body>" + SoapBody + "</soap:Body></soap:Envelope>";
            }
            return $.Deferred(function (dfd) {
                var Request = $.ajax(XrmTSToolkit.Common.GetSoapServiceURL(), {
                    data: soapXml,
                    type: "POST",
                    beforeSend: function (xhr) {
                        xhr.setRequestHeader("Accept", "application/xml, text/xml, */*");
                        xhr.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                        xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/" + RequestType);
                    }
                });
                Request.done(function (data, result, xhr) {
                    var sr = new Soap.SoapResponse(data);
                    dfd.resolve(sr, result, xhr);
                });
                Request.fail(function (req, description, exception) {
                    var ErrorText = typeof exception === "undefined" ? description : exception;
                    var fault = new Soap.FaultResponse(req.responseXML);
                    if (fault && fault.Fault && fault.Fault.FaultString) {
                        ErrorText += (": " + fault.Fault.FaultString);
                    }
                    dfd.reject({ success: false, result: ErrorText });
                });
            }).promise();
        };
        return Soap;
    })();
    XrmTSToolkit.Soap = Soap;
    var Soap;
    (function (Soap) {
        var Entity = (function () {
            function Entity(LogicalName, Id) {
                if (LogicalName === void 0) { LogicalName = ""; }
                if (Id === void 0) { Id = null; }
                this.LogicalName = LogicalName;
                this.Id = Id;
                this.Attributes = new AttributeCollection();
                if (!Id) {
                    this.Id = "00000000-0000-0000-0000-000000000000";
                }
            }
            Entity.prototype.Serialize = function () {
                var Data = "<b:Attributes>";
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
            };
            Entity.padNumber = function (s, len) {
                if (len === void 0) { len = 2; }
                s = '' + s;
                while (s.length < len) {
                    s = "0" + s;
                }
                return s;
            };
            Entity.EncodeValue = function (value) {
                return (typeof value === "object" && value.getTime) ? Entity.EncodeDate(value) : ((typeof window.CrmEncodeDecode != "undefined") ? window.CrmEncodeDecode.CrmXmlEncode(value) : Entity.CrmXmlEncode(value));
            };
            Entity.EncodeDate = function (Value) {
                return Value.getFullYear() + "-" + Entity.padNumber(Value.getMonth() + 1) + "-" + Entity.padNumber(Value.getDate()) + "T" + Entity.padNumber(Value.getHours()) + ":" + Entity.padNumber(Value.getMinutes()) + ":" + Entity.padNumber(Value.getSeconds());
            };
            Entity.CrmXmlEncode = function (s) {
                if ('undefined' === typeof s || 'unknown' === typeof s || null === s)
                    return s;
                else if (typeof s != "string")
                    s = s.toString();
                return Entity.innerSurrogateAmpersandWorkaround(s);
            };
            Entity.HtmlEncode = function (s) {
                if (s === null || s === "" || s === undefined)
                    return s;
                for (var count = 0, buffer = "", hEncode = "", cnt = 0; cnt < s.length; cnt++) {
                    var c = s.charCodeAt(cnt);
                    if (c > 96 && c < 123 || c > 64 && c < 91 || c === 32 || c > 47 && c < 58 || c === 46 || c === 44 || c === 45 || c === 95)
                        buffer += String.fromCharCode(c);
                    else
                        buffer += "&#" + c + ";";
                    if (++count === 500) {
                        hEncode += buffer;
                        buffer = "";
                        count = 0;
                    }
                }
                if (buffer.length)
                    hEncode += buffer;
                return hEncode;
            };
            Entity.innerSurrogateAmpersandWorkaround = function (s) {
                var buffer = '';
                var c0;
                for (var cnt = 0; cnt < s.length; cnt++) {
                    c0 = s.charCodeAt(cnt);
                    if (c0 >= 55296 && c0 <= 57343)
                        if (cnt + 1 < s.length) {
                            var c1 = s.charCodeAt(cnt + 1);
                            if (c1 >= 56320 && c1 <= 57343) {
                                buffer += "CRMEntityReferenceOpen" + ((c0 - 55296) * 1024 + (c1 & 1023) + 65536).toString(16) + "CRMEntityReferenceClose";
                                cnt++;
                            }
                            else
                                buffer += String.fromCharCode(c0);
                        }
                        else
                            buffer += String.fromCharCode(c0);
                    else
                        buffer += String.fromCharCode(c0);
                }
                s = buffer;
                buffer = "";
                for (cnt = 0; cnt < s.length; cnt++) {
                    c0 = s.charCodeAt(cnt);
                    if (c0 >= 55296 && c0 <= 57343)
                        buffer += String.fromCharCode(65533);
                    else
                        buffer += String.fromCharCode(c0);
                }
                s = buffer;
                s = Entity.HtmlEncode(s);
                s = s.replace(/CRMEntityReferenceOpen/g, "&#x");
                s = s.replace(/CRMEntityReferenceClose/g, ";");
                return s;
            };
            return Entity;
        })();
        Soap.Entity = Entity;
        (function (AttributeType) {
            AttributeType[AttributeType["AliasedValue"] = 0] = "AliasedValue";
            AttributeType[AttributeType["Boolean"] = 1] = "Boolean";
            AttributeType[AttributeType["Date"] = 2] = "Date";
            AttributeType[AttributeType["Decimal"] = 3] = "Decimal";
            AttributeType[AttributeType["Double"] = 4] = "Double";
            AttributeType[AttributeType["EntityCollection"] = 5] = "EntityCollection";
            AttributeType[AttributeType["EntityReference"] = 6] = "EntityReference";
            AttributeType[AttributeType["Float"] = 7] = "Float";
            AttributeType[AttributeType["Guid"] = 8] = "Guid";
            AttributeType[AttributeType["Integer"] = 9] = "Integer";
            AttributeType[AttributeType["OptionSetValue"] = 10] = "OptionSetValue";
            AttributeType[AttributeType["Money"] = 11] = "Money";
            AttributeType[AttributeType["String"] = 12] = "String";
        })(Soap.AttributeType || (Soap.AttributeType = {}));
        var AttributeType = Soap.AttributeType;
        var AttributeCollection = (function () {
            function AttributeCollection() {
            }
            return AttributeCollection;
        })();
        Soap.AttributeCollection = AttributeCollection;
        var EntityCollection = (function () {
            function EntityCollection() {
                this.Items = [];
            }
            EntityCollection.prototype.Serialize = function () {
                var XML = "";
                this.Items.forEach(function (Entity) {
                    XML += "<b:Entity>" + Entity.Serialize() + "</b:Entity>";
                });
                return XML;
            };
            return EntityCollection;
        })();
        Soap.EntityCollection = EntityCollection;
        var AttributeValue = (function () {
            function AttributeValue(Value, Type) {
                this.Value = Value;
                this.Type = Type;
            }
            AttributeValue.prototype.IsNull = function () {
                return this.Value === undefined;
            };
            AttributeValue.prototype.GetEncodedValue = function () {
                return Entity.EncodeValue(this.Value);
            };
            AttributeValue.prototype.Serialize = function () {
                return "<g:value i:type=\"" + this.GetValueType() + "\">" + this.GetEncodedValue() + "</g:value>";
                ;
            };
            AttributeValue.prototype.GetValueType = function () {
                return "";
            };
            return AttributeValue;
        })();
        Soap.AttributeValue = AttributeValue;
        var EntityCollectionAttribute = (function (_super) {
            __extends(EntityCollectionAttribute, _super);
            function EntityCollectionAttribute(Value) {
                _super.call(this, Value, 5 /* EntityCollection */);
                this.Value = Value;
            }
            EntityCollectionAttribute.prototype.GetValueType = function () {
                return "b:ArrayOfEntity";
            };
            return EntityCollectionAttribute;
        })(AttributeValue);
        Soap.EntityCollectionAttribute = EntityCollectionAttribute;
        var EntityReference = (function (_super) {
            __extends(EntityReference, _super);
            function EntityReference(Id, LogicalName, Name) {
                _super.call(this, { Id: Id, LogicalName: LogicalName, Name: Name }, 6 /* EntityReference */);
                this.Id = Id;
                this.LogicalName = LogicalName;
                this.Name = Name;
            }
            EntityReference.prototype.Serialize = function () {
                return "<g:value i:type=\"b:EntityReference\"><b:Id>" + Entity.EncodeValue(this.Id) + "</b:Id><b:LogicalName>" + Entity.EncodeValue(this.LogicalName) + "</b:LogicalName><b:Name i:nil=\"true\" /></g:value>";
            };
            EntityReference.prototype.GetValueType = function () {
                return "b:EntityReference";
            };
            return EntityReference;
        })(AttributeValue);
        Soap.EntityReference = EntityReference;
        var OptionSetValue = (function (_super) {
            __extends(OptionSetValue, _super);
            function OptionSetValue(Value) {
                _super.call(this, Value, 10 /* OptionSetValue */);
                this.Value = Value;
            }
            OptionSetValue.prototype.Serialize = function () {
                return "<g:value i:type=\"b:OptionSetValue\"><b:Value>" + this.GetEncodedValue() + "</b:Value></g:value>";
            };
            OptionSetValue.prototype.GetValueType = function () {
                return "b:OptionSetValue";
            };
            return OptionSetValue;
        })(AttributeValue);
        Soap.OptionSetValue = OptionSetValue;
        var MoneyValue = (function (_super) {
            __extends(MoneyValue, _super);
            function MoneyValue(Value) {
                _super.call(this, Value, 11 /* Money */);
                this.Value = Value;
            }
            MoneyValue.prototype.Serialize = function () {
                return "<g:value i:type=\"b:Money\"><b:Value>" + this.GetEncodedValue() + "</b:Value></g:value>";
            };
            MoneyValue.prototype.GetValueType = function () {
                return "b:Money";
            };
            return MoneyValue;
        })(AttributeValue);
        Soap.MoneyValue = MoneyValue;
        var AliasedValue = (function (_super) {
            __extends(AliasedValue, _super);
            function AliasedValue(Value, AttributeLogicalName, EntityLogicalName) {
                _super.call(this, Value, 0 /* AliasedValue */);
                this.Value = Value;
                this.AttributeLogicalName = AttributeLogicalName;
                this.EntityLogicalName = EntityLogicalName;
            }
            AliasedValue.prototype.Serialize = function () {
                throw "Update the 'Serialize' method of the 'AliasedValue'";
            };
            AliasedValue.prototype.GetValueType = function () {
                throw "Update the 'GetValueType' method of the 'AliasedValue'";
            };
            return AliasedValue;
        })(AttributeValue);
        Soap.AliasedValue = AliasedValue;
        var BooleanValue = (function (_super) {
            __extends(BooleanValue, _super);
            function BooleanValue(Value) {
                _super.call(this, Value, 1 /* Boolean */);
                this.Value = Value;
            }
            BooleanValue.prototype.GetValueType = function () {
                return "d:boolean";
            };
            return BooleanValue;
        })(AttributeValue);
        Soap.BooleanValue = BooleanValue;
        var IntegerValue = (function (_super) {
            __extends(IntegerValue, _super);
            function IntegerValue(Value) {
                _super.call(this, Value, 9 /* Integer */);
                this.Value = Value;
            }
            IntegerValue.prototype.GetValueType = function () {
                return "d:int";
            };
            return IntegerValue;
        })(AttributeValue);
        Soap.IntegerValue = IntegerValue;
        var StringValue = (function (_super) {
            __extends(StringValue, _super);
            function StringValue(Value) {
                _super.call(this, Value, 12 /* String */);
                this.Value = Value;
            }
            StringValue.prototype.GetValueType = function () {
                return "d:string";
            };
            return StringValue;
        })(AttributeValue);
        Soap.StringValue = StringValue;
        var DoubleValue = (function (_super) {
            __extends(DoubleValue, _super);
            function DoubleValue(Value) {
                _super.call(this, Value, 4 /* Double */);
                this.Value = Value;
            }
            DoubleValue.prototype.GetValueType = function () {
                return "d:double";
            };
            return DoubleValue;
        })(AttributeValue);
        Soap.DoubleValue = DoubleValue;
        var DecimalValue = (function (_super) {
            __extends(DecimalValue, _super);
            function DecimalValue(Value) {
                _super.call(this, Value, 3 /* Decimal */);
                this.Value = Value;
            }
            DecimalValue.prototype.GetValueType = function () {
                return "d:decimal";
            };
            return DecimalValue;
        })(AttributeValue);
        Soap.DecimalValue = DecimalValue;
        var FloatValue = (function (_super) {
            __extends(FloatValue, _super);
            function FloatValue(Value) {
                _super.call(this, Value, 7 /* Float */);
                this.Value = Value;
            }
            FloatValue.prototype.GetValueType = function () {
                return "d:float";
            };
            return FloatValue;
        })(AttributeValue);
        Soap.FloatValue = FloatValue;
        var DateValue = (function (_super) {
            __extends(DateValue, _super);
            function DateValue(Value) {
                _super.call(this, Value, 2 /* Date */);
                this.Value = Value;
            }
            DateValue.prototype.GetEncodedValue = function () {
                return Entity.EncodeDate(this.Value);
            };
            DateValue.prototype.GetValueType = function () {
                return "d:dateTime";
            };
            return DateValue;
        })(AttributeValue);
        Soap.DateValue = DateValue;
        var GuidValue = (function (_super) {
            __extends(GuidValue, _super);
            function GuidValue(Value) {
                _super.call(this, Value, 8 /* Guid */);
                this.Value = Value;
            }
            GuidValue.prototype.GetValueType = function () {
                return "e:guid";
            };
            return GuidValue;
        })(AttributeValue);
        Soap.GuidValue = GuidValue;
        var NodeValue = (function () {
            function NodeValue(Name, Value) {
                this.Name = Name;
                this.Value = Value;
            }
            return NodeValue;
        })();
        Soap.NodeValue = NodeValue;
        var Collection = (function () {
            function Collection(Items) {
                this.Items = Items;
            }
            return Collection;
        })();
        Soap.Collection = Collection;
        var SoapResponse = (function () {
            function SoapResponse(ResponseXML) {
                XML: ResponseXML;
                this.ParseResult(ResponseXML);
            }
            SoapResponse.prototype.ParseResult = function (ResponseXML) {
                var BodyNode = $(ResponseXML).find("s\\:Body, Body");
                var val = this.GetNodeValue($(BodyNode[0].firstChild));
                this[val.Name] = val.Value;
                this.SDKResponse = val.Value;
            };
            SoapResponse.prototype.GetNodeValue = function (Node) {
                var NodeName = RemoveNameSpace(Node[0].nodeName);
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
                    var ChildValues = new Array();
                    $.each(Node[0].childNodes, function (index, ChildNode) {
                        var ChileNodeValue = SoapResponse.prototype.GetNodeValue($(ChildNode));
                        ChildValues.push(ChileNodeValue);
                    });
                    var IsCollection = false;
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
                        $.each(ChildValues, function (Index, Value) {
                            coll[Value.Name] = Value.Value;
                            items.push(Value);
                        });
                        coll.Items = items;
                        return new NodeValue(NodeName, coll);
                    }
                    else if (NodeName === "EntityReference") {
                        var ref = new EntityReference();
                        ref.Id = ChildValues[0].Value;
                        ref.LogicalName = ChildValues[1].Value;
                        ref.Name = ChildValues[2].Value;
                        return new NodeValue(NodeName, ref);
                    }
                    else if (NodeName === "OptionSetValue") {
                        var opt = new OptionSetValue();
                        if (ChildValues[0].Value) {
                            opt.Value = parseInt(ChildValues[0].Value);
                        }
                        return new NodeValue(NodeName, opt);
                    }
                    else if (NodeName === "Money") {
                        var cur = new MoneyValue();
                        if (ChildValues[0].Value) {
                            cur.Value = parseFloat(ChildValues[0].Value);
                        }
                        return new NodeValue(NodeName, cur);
                    }
                    else if (NodeName === "EntityCollection" || NodeName === "Entities") {
                        var items = new Array();
                        $.each(ChildValues, function (Index, Value) {
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
                        $.each(ChildValues, function (Index, Value) {
                            obj[Value.Name] = Value.Value;
                        });
                        return new NodeValue(NodeName, obj);
                    }
                }
                else {
                    var Value = Node.text() === "" ? (Node.attr("i:nil") === "true" ? null : "") : Node.text();
                    if (NodeName === "int") {
                        var int = new IntegerValue();
                        if (Value) {
                            int.Value = parseInt(Value);
                        }
                        return new NodeValue(NodeName, int);
                    }
                    else if (NodeName === "double") {
                        var double = new DoubleValue();
                        if (Value) {
                            double.Value = parseFloat(Value);
                        }
                        return new NodeValue(NodeName, double);
                    }
                    else if (NodeName === "float") {
                        var float = new FloatValue();
                        if (Value) {
                            float.Value = parseFloat(Value);
                        }
                        return new NodeValue(NodeName, float);
                    }
                    else if (NodeName === "decimal") {
                        var dec = new DecimalValue();
                        if (Value) {
                            dec.Value = parseFloat(Value);
                        }
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
                        if (Value) {
                            boolean.Value = Value == 'false' ? false : true;
                        }
                        return new NodeValue(NodeName, boolean);
                    }
                    else if (NodeName === "string") {
                        var str = new StringValue();
                        if (Value) {
                            str.Value = Value;
                        }
                        return new NodeValue(NodeName, str);
                    }
                    else if (NodeName === "guid") {
                        var guid = new GuidValue();
                        if (Value) {
                            guid.Value = Value;
                        }
                        return new NodeValue(NodeName, guid);
                    }
                    else {
                        return new NodeValue(NodeName, Value);
                    }
                }
            };
            return SoapResponse;
        })();
        Soap.SoapResponse = SoapResponse;
        var SDKResponse = (function () {
            function SDKResponse() {
            }
            return SDKResponse;
        })();
        Soap.SDKResponse = SDKResponse;
        var CreateSoapResponse = (function (_super) {
            __extends(CreateSoapResponse, _super);
            function CreateSoapResponse() {
                _super.apply(this, arguments);
            }
            return CreateSoapResponse;
        })(SDKResponse);
        Soap.CreateSoapResponse = CreateSoapResponse;
        var UpdateSoapResponse = (function (_super) {
            __extends(UpdateSoapResponse, _super);
            function UpdateSoapResponse() {
                _super.apply(this, arguments);
            }
            return UpdateSoapResponse;
        })(SDKResponse);
        Soap.UpdateSoapResponse = UpdateSoapResponse;
        var RetrieveSoapResponse = (function (_super) {
            __extends(RetrieveSoapResponse, _super);
            function RetrieveSoapResponse() {
                _super.apply(this, arguments);
            }
            return RetrieveSoapResponse;
        })(SDKResponse);
        Soap.RetrieveSoapResponse = RetrieveSoapResponse;
        var RetrieveMultipleSoapResponse = (function (_super) {
            __extends(RetrieveMultipleSoapResponse, _super);
            function RetrieveMultipleSoapResponse() {
                _super.apply(this, arguments);
            }
            return RetrieveMultipleSoapResponse;
        })(SDKResponse);
        Soap.RetrieveMultipleSoapResponse = RetrieveMultipleSoapResponse;
        var RetrieveMultipleResult = (function () {
            function RetrieveMultipleResult() {
                this.MoreRecords = false;
            }
            return RetrieveMultipleResult;
        })();
        Soap.RetrieveMultipleResult = RetrieveMultipleResult;
        var FaultResponse = (function (_super) {
            __extends(FaultResponse, _super);
            function FaultResponse() {
                _super.apply(this, arguments);
            }
            return FaultResponse;
        })(SoapResponse);
        Soap.FaultResponse = FaultResponse;
        var Fault = (function () {
            function Fault() {
            }
            return Fault;
        })();
        Soap.Fault = Fault;
        var ColumnSet = (function () {
            function ColumnSet(p) {
                this.AllColumns = false;
                if (typeof p === "boolean") {
                    this.AllColumns = p;
                    this.Columns = [];
                }
                else {
                    this.Columns = p;
                }
            }
            ColumnSet.prototype.AddColumn = function (Column) {
                this.Columns.push(Column);
            };
            ColumnSet.prototype.AddColumns = function (Columns) {
                for (var Column in Columns) {
                    this.AddColumn(Column);
                }
            };
            ColumnSet.prototype.serialize = function () {
                var Data = "<b:ColumnSet>" + "<b:AllColumns>" + this.AllColumns.toString() + "</b:AllColumns>";
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
            };
            return ColumnSet;
        })();
        Soap.ColumnSet = ColumnSet;
    })(Soap = XrmTSToolkit.Soap || (XrmTSToolkit.Soap = {}));
    var Soap;
    (function (Soap) {
        var Query;
        (function (Query) {
            var QueryExpression = (function () {
                function QueryExpression(EntityName) {
                    this.EntityName = EntityName;
                    this.Columns = new Soap.ColumnSet(false);
                    this.Criteria = null;
                    this.PageInfo = new PageInfo();
                    this.Orders = [];
                    this.LinkEntities = [];
                    this.NoLock = false;
                }
                QueryExpression.prototype.serialize = function () {
                    //Query
                    var Data = "<query i:type=\"b:QueryExpression\">";
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
                        this.LinkEntities.forEach(function (LinkEntity) {
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
                        this.Orders.forEach(function (Order) {
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
                };
                return QueryExpression;
            })();
            Query.QueryExpression = QueryExpression;
            var FilterExpression = (function () {
                function FilterExpression(FilterOperator) {
                    this.FilterOperator = 0 /* And */;
                    this.Conditions = [];
                    this.Filters = [];
                    this.IsQuickFindFilter = false;
                    if (FilterOperator === undefined) {
                        this.FilterOperator = 0 /* And */;
                    }
                }
                FilterExpression.prototype.AddFilter = function (p) {
                    if (typeof p === "number")
                        p = new FilterExpression(p);
                    this.Filters.push(p);
                };
                FilterExpression.prototype.AddCondition = function (p, Operator, Values) {
                    if (typeof p === "string") {
                        p = new ConditionExpression(p, Operator, Values);
                    }
                    this.Conditions.push(p);
                    return p;
                };
                FilterExpression.prototype.serialize = function () {
                    var Data = "<b:Conditions>";
                    //Conditions
                    this.Conditions.forEach(function (Condition) {
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
                        this.Filters.forEach(function (Filter) {
                            Data += "<b:FilterExpression>";
                            Data += Filter.serialize();
                            Data += "</b:FilterExpression>";
                        });
                        Data += "</b:Filters>";
                    }
                    //IsQuickFindFilter
                    Data += "<b:IsQuickFindFilter>" + this.IsQuickFindFilter.toString() + "</b:IsQuickFindFilter>";
                    return Data;
                };
                return FilterExpression;
            })();
            Query.FilterExpression = FilterExpression;
            var ConditionExpression = (function () {
                function ConditionExpression(AttributeName, Operator, Values) {
                    this.AttributeName = AttributeName;
                    this.Operator = Operator;
                    this.Values = [];
                    if (Values instanceof Array) {
                        this.Values = Values;
                    }
                    else {
                        this.Values.push(Values);
                    }
                }
                ConditionExpression.prototype.serialize = function () {
                    var Data = "<b:ConditionExpression>" + "<b:AttributeName>" + this.AttributeName + "</b:AttributeName>" + "<b:Operator>" + ConditionOperator[this.Operator].toString() + "</b:Operator>" + "<b:Values>";
                    (this.Values).forEach(function (Value) {
                        Data += "<c:anyType i:type=\"" + Value.GetValueType() + "\">" + Value.GetEncodedValue() + "</c:anyType>";
                    });
                    Data += "</b:Values>" + "</b:ConditionExpression>";
                    return Data;
                };
                return ConditionExpression;
            })();
            Query.ConditionExpression = ConditionExpression;
            var PageInfo = (function () {
                function PageInfo() {
                    this.Count = 0;
                    this.PageNumber = 0;
                    this.PagingCookie = null;
                    this.ReturnTotalRecordCount = false;
                }
                PageInfo.prototype.serialize = function () {
                    var Data = "" + "<b:PageInfo>" + "<b:Count>" + this.Count.toString() + "</b:Count>" + "<b:PageNumber>" + this.PageNumber.toString() + "</b:PageNumber>";
                    if (this.PagingCookie == null) {
                        Data += "<b:PagingCookie i:nil = \"true\" />";
                    }
                    else {
                        Data += "<b:PagingCookie>" + this.PagingCookie + "</b:PagingCookie>";
                    }
                    Data += "<b:ReturnTotalRecordCount>" + this.ReturnTotalRecordCount.toString() + "</b:ReturnTotalRecordCount></b:PageInfo>";
                    return Data;
                };
                return PageInfo;
            })();
            Query.PageInfo = PageInfo;
            var OrderExpression = (function () {
                function OrderExpression(AttributeName, OrderType) {
                    this.AttributeName = AttributeName;
                    this.OrderType = OrderType;
                }
                OrderExpression.prototype.serialize = function () {
                    var Data = "<b:OrderExpression>";
                    Data += "<b:AttributeName>" + this.AttributeName + "</b:AttributeName>";
                    Data += "<b:OrderType>" + OrderType[this.OrderType].toString() + "</b:OrderType>";
                    Data += "</b:OrderExpression>";
                    return Data;
                };
                return OrderExpression;
            })();
            Query.OrderExpression = OrderExpression;
            var LinkEntity = (function () {
                function LinkEntity(LinkFromEntityName, LinkToEntityName, LinkFromAttributeName, LinkToAttributeName, JoinOperator) {
                    if (JoinOperator === void 0) { JoinOperator = 0 /* Inner */; }
                    this.LinkFromEntityName = LinkFromEntityName;
                    this.LinkToEntityName = LinkToEntityName;
                    this.LinkFromAttributeName = LinkFromAttributeName;
                    this.LinkToAttributeName = LinkToAttributeName;
                    this.JoinOperator = JoinOperator;
                    this.Columns = new Soap.ColumnSet(false);
                    this.EntityAlias = null;
                    this.LinkCriteria = new FilterExpression(0 /* And */);
                    this.LinkEntities = [];
                }
                LinkEntity.prototype.serialize = function () {
                    var Data = "<b:LinkEntity>";
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
                        this.LinkEntities.forEach(function (LinkEntity) {
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
                };
                return LinkEntity;
            })();
            Query.LinkEntity = LinkEntity;
            (function (JoinOperator) {
                JoinOperator[JoinOperator["Inner"] = 0] = "Inner";
                JoinOperator[JoinOperator["LeftOuter"] = 1] = "LeftOuter";
                JoinOperator[JoinOperator["Natural"] = 2] = "Natural";
            })(Query.JoinOperator || (Query.JoinOperator = {}));
            var JoinOperator = Query.JoinOperator;
            (function (OrderType) {
                OrderType[OrderType["Ascending"] = 0] = "Ascending";
                OrderType[OrderType["Descending"] = 1] = "Descending";
            })(Query.OrderType || (Query.OrderType = {}));
            var OrderType = Query.OrderType;
            (function (LogicalOperator) {
                LogicalOperator[LogicalOperator["And"] = 0] = "And";
                LogicalOperator[LogicalOperator["Or"] = 1] = "Or";
            })(Query.LogicalOperator || (Query.LogicalOperator = {}));
            var LogicalOperator = Query.LogicalOperator;
            (function (ConditionOperator) {
                ConditionOperator[ConditionOperator["Equal"] = 0] = "Equal";
                ConditionOperator[ConditionOperator["NotEqual"] = 1] = "NotEqual";
                ConditionOperator[ConditionOperator["GreaterThan"] = 2] = "GreaterThan";
                ConditionOperator[ConditionOperator["LessThan"] = 3] = "LessThan";
                ConditionOperator[ConditionOperator["GreaterEqual"] = 4] = "GreaterEqual";
                ConditionOperator[ConditionOperator["LessEqual"] = 5] = "LessEqual";
                ConditionOperator[ConditionOperator["Like"] = 6] = "Like";
                ConditionOperator[ConditionOperator["NotLike"] = 7] = "NotLike";
                ConditionOperator[ConditionOperator["In"] = 8] = "In";
                ConditionOperator[ConditionOperator["NotIn"] = 9] = "NotIn";
                ConditionOperator[ConditionOperator["Between"] = 10] = "Between";
                ConditionOperator[ConditionOperator["NotBetween"] = 11] = "NotBetween";
                ConditionOperator[ConditionOperator["Null"] = 12] = "Null";
                ConditionOperator[ConditionOperator["NotNull"] = 13] = "NotNull";
                ConditionOperator[ConditionOperator["Yesterday"] = 14] = "Yesterday";
                ConditionOperator[ConditionOperator["Today"] = 15] = "Today";
                ConditionOperator[ConditionOperator["Tomorrow"] = 16] = "Tomorrow";
                ConditionOperator[ConditionOperator["Last7Days"] = 17] = "Last7Days";
                ConditionOperator[ConditionOperator["Next7Days"] = 18] = "Next7Days";
                ConditionOperator[ConditionOperator["LastWeek"] = 19] = "LastWeek";
                ConditionOperator[ConditionOperator["ThisWeek"] = 20] = "ThisWeek";
                ConditionOperator[ConditionOperator["NextWeek"] = 21] = "NextWeek";
                ConditionOperator[ConditionOperator["LastMonth"] = 22] = "LastMonth";
                ConditionOperator[ConditionOperator["ThisMonth"] = 23] = "ThisMonth";
                ConditionOperator[ConditionOperator["NextMonth"] = 24] = "NextMonth";
                ConditionOperator[ConditionOperator["On"] = 25] = "On";
                ConditionOperator[ConditionOperator["OnOrBefore"] = 26] = "OnOrBefore";
                ConditionOperator[ConditionOperator["OnOrAfter"] = 27] = "OnOrAfter";
                ConditionOperator[ConditionOperator["LastYear"] = 28] = "LastYear";
                ConditionOperator[ConditionOperator["ThisYear"] = 29] = "ThisYear";
                ConditionOperator[ConditionOperator["NextYear"] = 30] = "NextYear";
                ConditionOperator[ConditionOperator["LastXHours"] = 31] = "LastXHours";
                ConditionOperator[ConditionOperator["NextXHours"] = 32] = "NextXHours";
                ConditionOperator[ConditionOperator["LastXDays"] = 33] = "LastXDays";
                ConditionOperator[ConditionOperator["NextXDays"] = 34] = "NextXDays";
                ConditionOperator[ConditionOperator["LastXWeeks"] = 35] = "LastXWeeks";
                ConditionOperator[ConditionOperator["NextXWeeks"] = 36] = "NextXWeeks";
                ConditionOperator[ConditionOperator["LastXMonths"] = 37] = "LastXMonths";
                ConditionOperator[ConditionOperator["NextXMonths"] = 38] = "NextXMonths";
                ConditionOperator[ConditionOperator["LastXYears"] = 39] = "LastXYears";
                ConditionOperator[ConditionOperator["NextXYears"] = 40] = "NextXYears";
                ConditionOperator[ConditionOperator["EqualUserId"] = 41] = "EqualUserId";
                ConditionOperator[ConditionOperator["NotEqualUserId"] = 42] = "NotEqualUserId";
                ConditionOperator[ConditionOperator["EqualBusinessId"] = 43] = "EqualBusinessId";
                ConditionOperator[ConditionOperator["NotEqualBusinessId"] = 44] = "NotEqualBusinessId";
                ConditionOperator[ConditionOperator["ChildOf"] = 45] = "ChildOf";
                ConditionOperator[ConditionOperator["Mask"] = 46] = "Mask";
                ConditionOperator[ConditionOperator["NotMask"] = 47] = "NotMask";
                ConditionOperator[ConditionOperator["MasksSelect"] = 48] = "MasksSelect";
                ConditionOperator[ConditionOperator["Contains"] = 49] = "Contains";
                ConditionOperator[ConditionOperator["DoesNotContain"] = 50] = "DoesNotContain";
                ConditionOperator[ConditionOperator["EqualUserLanguage"] = 51] = "EqualUserLanguage";
                ConditionOperator[ConditionOperator["NotOn"] = 52] = "NotOn";
                ConditionOperator[ConditionOperator["OlderThanXMonths"] = 53] = "OlderThanXMonths";
                ConditionOperator[ConditionOperator["BeginsWith"] = 54] = "BeginsWith";
                ConditionOperator[ConditionOperator["DoesNotBeginWith"] = 55] = "DoesNotBeginWith";
                ConditionOperator[ConditionOperator["EndsWith"] = 56] = "EndsWith";
                ConditionOperator[ConditionOperator["DoesNotEndWith"] = 57] = "DoesNotEndWith";
                ConditionOperator[ConditionOperator["ThisFiscalYear"] = 58] = "ThisFiscalYear";
                ConditionOperator[ConditionOperator["ThisFiscalPeriod"] = 59] = "ThisFiscalPeriod";
                ConditionOperator[ConditionOperator["NextFiscalYear"] = 60] = "NextFiscalYear";
                ConditionOperator[ConditionOperator["NextFiscalPeriod"] = 61] = "NextFiscalPeriod";
                ConditionOperator[ConditionOperator["LastFiscalYear"] = 62] = "LastFiscalYear";
                ConditionOperator[ConditionOperator["LastFiscalPeriod"] = 63] = "LastFiscalPeriod";
                ConditionOperator[ConditionOperator["LastXFiscalYears"] = 64] = "LastXFiscalYears";
                ConditionOperator[ConditionOperator["LastXFiscalPeriods"] = 65] = "LastXFiscalPeriods";
                ConditionOperator[ConditionOperator["NextXFiscalYears"] = 66] = "NextXFiscalYears";
                ConditionOperator[ConditionOperator["NextXFiscalPeriods"] = 67] = "NextXFiscalPeriods";
                ConditionOperator[ConditionOperator["InFiscalYear"] = 68] = "InFiscalYear";
                ConditionOperator[ConditionOperator["InFiscalPeriod"] = 69] = "InFiscalPeriod";
                ConditionOperator[ConditionOperator["InFiscalPeriodAndYear"] = 70] = "InFiscalPeriodAndYear";
                ConditionOperator[ConditionOperator["InOrBeforeFiscalPeriodAndYear"] = 71] = "InOrBeforeFiscalPeriodAndYear";
                ConditionOperator[ConditionOperator["InOrAfterFiscalPeriodAndYear"] = 72] = "InOrAfterFiscalPeriodAndYear";
                ConditionOperator[ConditionOperator["EqualUserTeams"] = 73] = "EqualUserTeams";
            })(Query.ConditionOperator || (Query.ConditionOperator = {}));
            var ConditionOperator = Query.ConditionOperator;
        })(Query = Soap.Query || (Soap.Query = {}));
    })(Soap = XrmTSToolkit.Soap || (XrmTSToolkit.Soap = {}));
})(XrmTSToolkit || (XrmTSToolkit = {}));
String.prototype.replaceAll = function (search, replace) {
    //if replace is null, return original string otherwise it will
    //replace search string with 'undefined'.
    if (!replace)
        return this;
    return this.replace(new RegExp('[' + search + ']', 'g'), replace);
};
//# sourceMappingURL=XrmTSToolkit.js.map