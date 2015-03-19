/// <reference path="../jquery/jquery.d.ts" />
/// <reference path="xrm.d.ts" />

/**
* MSCRM 2011, 2013, 2015 Service Toolkit for TypeScript
* @author Steven Rasmussen
* @current version : 0.7.0
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
* Version : 0.6.1
* Date: March 13, 2015
*   Added 'Delete', 'Associate', 'Disassociate' methods
*   Added tests for each method
*   Fixed bugs found by tests :)
*   Re-wrote XML response parsing portion
*   Removed requirement for Xrm-2015.d.ts
*   Added Xrm.d.ts to package
********************************************
* Version : 0.7.0
* Date: March 19, 2015
*   Added support for the 'ExecuteMultiple' request.  In order to make the 'ExecuteMultiple' request useful, the following request types were created:
*      'ExecuteMultipleRequest', 'CreateRequest', 'UpdateRequest', 'DeleteRequest', 'AssociateRequest', 'DisassociateRequest', 'SetStateRequest', 'WhoAmIRequest', 'ExecuteMultiple', 
*      'AssignRequest', 'GrantAccessRequest', 'ModifyAccessRequest', 'RevokeAccessRequest', 'RetrievePrincipleAccessRequest'.
*      You can pass any of these objects to the 'Execute' method similar to the SDK.
*   Tests were created for each of the requests above.
*   Improved 'FaultResponse' class definition
*   Added tests for 'FaultResponse'
*   Added comments to methods and requests
*   RetrieveMultiple and Fetch methods are now able to return all results (>5000 records) - set the 'Soap.RetrieveAllEntities = true', default is false
*   Moved XML parsing to XrmTSToolkit.XML class and module
*   Minor bug fixes
********************************************
*/

module XrmTSToolkit {
    export class Common {

        /**
         * Gets server URL.
         *
         * @return  The server URL.
         */
        public static GetServerURL(): string {
            var URL: string = document.location.protocol + "//" + document.location.host;
            var Org: string = Xrm.Page.context.getOrgUniqueName();
            if (document.location.pathname.toUpperCase().indexOf(Org.toUpperCase()) > -1) {
                URL += "/" + Org;
            }
            return URL;
        }

        /**
         * Gets the CRM SOAP service URL.
         *
         * @return  The SOAP service URL.
         */
        public static GetSoapServiceURL(): string {
            return Common.GetServerURL() + "/XRMServices/2011/Organization.svc/web";
        }

        /** Tests for Xrm context. */
        public static TestForContext(): void {
            if (typeof Xrm === "undefined") {
                throw new Error("Please make sure to add the ClientGlobalContext.js.aspx file to your web resource.");
            }
        }

        /**
         * Gets a URL parameter.
         *
         * @param   name    The name of the parameter.
         *
         * @return  The URL parameter.
         */
        static GetURLParameter(name): string {
            return decodeURI((RegExp(name + '=' + '(.+?)(&|$)').exec(location.search) || [, null])[1]);
        }
        static StripGUID(GUID: string): string {
            return GUID.replace("{", "").replace("}", "");
        }
    }
    export class XML {
        public static ParseNode(xNode: Node): XML.XMLElement {
            var Element = new XML.XMLElement();
            var NodeName = $(xNode).prop("nodeName");
            if (NodeName.indexOf(":") > 0) {
                NodeName = NodeName.substring(NodeName.indexOf(":") + 1);
            }
            Element.Name = NodeName;
            Element.TypeName = $(xNode).attr("i:type");
            if (Element.TypeName && Element.TypeName.indexOf(":") > 0) {
                Element.TypeName = Element.TypeName.substring(Element.TypeName.indexOf(":") + 1);
            }
            var ChildNodes = $(xNode).contents();
            if (ChildNodes && ChildNodes.length > 0) {
                $.each(ChildNodes, function (i, ChildNode) {
                    var ChildNodeType = $(ChildNode).prop("nodeType");
                    if (ChildNodeType && ChildNodeType != 3) {
                        var ChildElement = XML.ParseNode(ChildNode);
                        ChildElement.Parent = Element;
                        Element.Children.push(ChildElement);

                    }
                });
            }
            if (Element.Children.length == 0) {
                Element.Text = $(xNode).text();
            }
            return Element;
        }
        public static FilterNode = function (name) {
            return this.find('*').filter(function () {
                return this.nodeName === name;
            });
        };
    }

    export class Soap {

        /** "true" to retrieve all entities for RetrieveMultiple and Fetch operations, otherwise "false" */
        public static RetrieveAllEntities: boolean = false;

        /**
         * Creates a new record in CRM
         *
         * @param   {Soap.Entity}   Entity  The Soap.Entity to create in CRM.
         *
         * @return  A JQueryPromise&lt;Soap.CreateSoapResponse&gt;
         */
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

        /**
         * Updates the given Entity in CRM.
         *
         * @param   {Soap.Entity}   Entity  The Soap.Entity to update in CRM..
         *
         * @return  A JQueryPromise&lt;Soap.UpdateSoapResponse&gt;
         */
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

        /**
         * Deletes the given Entity from CRM.
         *
         * @param   {Soap.EntityReference}  Entity  The entity to delete.
         *
         * @return  A JQueryPromise&lt;Soap.UpdateSoapResponse&gt;
         */
        static Delete(Entity: Soap.EntityReference): JQueryPromise<Soap.DeleteSoapResponse>;

        /**
         * Deletes this object.
         *
         * @param   {string}    EntityName  Logical name of the entity to be deleted.
         * @param   {string}    EntityId    GUID of the entity to be deleted.
         *
         * @return  A JQueryPromise&lt;Soap.DeleteSoapResponse&gt;
         */
        static Delete(EntityName: string, EntityId: string): JQueryPromise<Soap.DeleteSoapResponse>;
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

        /**
         * Retrieves the specified record from CRM.
         *
         * @param   {string}    Id                  The GUID of the entity to retrieve.
         * @param   {string}    EntityLogicalName   Entity logical name.
         * @param   {Soap.ColumnSet}    ColumnSet   Columns to retrieve as part of the entity.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveSoapResponse&gt;
         */
        static Retrieve(Id: string, EntityLogicalName: string, ColumnSet?: Soap.ColumnSet): JQueryPromise<Soap.RetrieveSoapResponse> {
            Id = Common.StripGUID(Id);
            if (!ColumnSet || ColumnSet == null) { ColumnSet = new Soap.ColumnSet(false); }
            var msgBody = "<entityName>" + EntityLogicalName + "</entityName><id>" + Id + "</id><columnSet>" + ColumnSet.Serialize() + "</columnSet>";
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

        /**
         * Retrieves records from CRM that match the provided criteria.
         *
         * @param   {Soap.Query.QueryExpression}    Query   The query specifying the criteria and other parameters.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveMultipleSoapResponse&gt;
         */
        static RetrieveMultiple(Query: Soap.Query.QueryExpression): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.RetrieveMultipleSoapResponse>(Query.serialize(), "RetrieveMultiple");
                Request.done(function (data, result, xhr) {
                    if (data.RetrieveMultipleResult.MoreRecords == true && Soap.RetrieveAllEntities == true) {
                        Query.PageInfo.PagingCookie = Soap.Entity.EncodeValue(data.RetrieveMultipleResult.PagingCookie);
                        Query.PageInfo.PageNumber = (Query.PageInfo.PageNumber + 1);
                        var AdditionalRequest = Soap.RetrieveMultiple(Query);
                        AdditionalRequest.done(function (data1, result, xhr) {
                            $.each(data1.RetrieveMultipleResult.Entities, function (i, Entity) {
                                data.RetrieveMultipleResult.Entities.push(Entity);
                            });
                            dfd.resolve(data);
                        });
                        AdditionalRequest.fail(function (result) {
                            dfd.reject(result);
                        });
                    }
                    else {
                        dfd.resolve(data);
                    }
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Retrieves related many to many.
         *
         * @param   {string}    LinkFromEntityName                          Entity logical name of the known entity.
         * @param   {string}    LinkFromEntityId                            Id of the known entity
         * @param   {string}    LinkToEntityName                            Entity logical name of the records to be retrieved.
         * @param   {string}    IntermediateTableName                       Name of the intermediate table of the N:N relationship.
         * @param   {Soap.ColumnSet}    optional Columns                    The columns to be retrieved on the returned records.
         * @param   {Array{Soap.Query.OrderExpression}} optional SortOrders Sort order of the returned records.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveMultipleSoapResponse&gt;
         */
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

        /**
         * Performs a 'Fetch' operation, similar to a 'RetrieveMultiple' but using FetchXML as the query string.
         *
         * @param   {string}    fetchXml    The FetchXML query.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveMultipleSoapResponse&gt;
         */
        static Fetch(fetchXml: string): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            var XMLDoc = $.parseXML(fetchXml);
            //First decode all invalid characters
            fetchXml = Soap.DecodeFetchXML(fetchXml);

            //now re-encode all invalid characters
            fetchXml = Soap.EncodeFetchXML(fetchXml);

            var msgBody = "<query i:type='a:FetchExpression'><a:Query>" + fetchXml + "</a:Query></query>";
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var Request = Soap.DoRequest<Soap.RetrieveMultipleSoapResponse>(msgBody, "RetrieveMultiple");
                Request.done(function (data, result, xhr) {
                    if (data.RetrieveMultipleResult.MoreRecords == true && Soap.RetrieveAllEntities == true) {
                        var PagingCookie = XMLDoc.firstChild.attributes.getNamedItem("paging-cookie");
                        if (!PagingCookie) { PagingCookie = XMLDoc.createAttribute("paging-cookie"); }
                        else { XMLDoc.firstChild.attributes.removeNamedItem("paging-cookie"); }

                        PagingCookie.value = Soap.Entity.EncodeValue(data.RetrieveMultipleResult.PagingCookie);
                        XMLDoc.firstChild.attributes.setNamedItem(PagingCookie);

                        var PageNumber = XMLDoc.firstChild.attributes.getNamedItem("page");
                        var CurrentPageNumber: number = 1;
                        if (!PageNumber) {
                            PageNumber = XMLDoc.createAttribute("page");
                        }
                        else {
                            CurrentPageNumber = parseInt(PageNumber.value);
                            XMLDoc.firstChild.attributes.removeNamedItem("page");
                        }
                        CurrentPageNumber += 1;
                        PageNumber.value = CurrentPageNumber.toString();
                        XMLDoc.firstChild.attributes.setNamedItem(PageNumber);
                        var AdditionalRequest = Soap.Fetch(Soap.XMLtoString(XMLDoc));
                        AdditionalRequest.done(function (data1, result, xhr) {
                            $.each(data1.RetrieveMultipleResult.Entities, function (i, Entity) {
                                data.RetrieveMultipleResult.Entities.push(Entity);
                            });
                            dfd.resolve(data);
                        });
                        AdditionalRequest.fail(function (result) {
                            dfd.reject(result);
                        });
                    }
                    else {
                        dfd.resolve(data);
                    }
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        private static DecodeFetchXML(FetchXML: string): string {
            FetchXML = FetchXML.replace(/&amp;/g, "&");
            FetchXML = FetchXML.replace(/&lt;/g, "<");
            FetchXML = FetchXML.replace(/&gt;/g, ">");
            FetchXML = FetchXML.replace(/&apos;/g, "'");
            FetchXML = FetchXML.replace(/&quot;/g, "\"");
            return FetchXML;
        }
        private static EncodeFetchXML(FetchXML: string): string {
            FetchXML = FetchXML.replace(/&/g, "&amp;");
            FetchXML = FetchXML.replace(/</g, "&lt;");
            FetchXML = FetchXML.replace(/>/g, "&gt;");
            FetchXML = FetchXML.replace(/'/g, "&apos;");
            FetchXML = FetchXML.replace(/\"/g, "&quot;");
            return FetchXML;
        }
        private static XMLtoString(elem) {
            var serialized;
            try {
                // XMLSerializer exists in current Mozilla browsers
                var serializer = new XMLSerializer();
                serialized = serializer.serializeToString(elem);
            }
            catch (e) {
                // Internet Explorer has a different approach to serializing XML
                serialized = elem.xml;
            }
            return serialized;
        }
        
        /**
         * Performs an 'Execute' request.
         *
         * @tparam  T   Type of the soap response to be returned.
         * @param   {string}    ExecuteXML  contents of the '&lt;Execute&gt;' tags,  INCLUDING the '&lt;
         *                                  Execute&gt;' tags.
         *
         * @return  A JQueryPromise&lt;T&gt;
         */
        static Execute<T>(ExecuteXML: string): JQueryPromise<T>;

        /**
         * Executes the specified request.
         *
         * @tparam  T   Type of the soap response to be returned.
         * @param   {Soap.ExecuteRequest}   ExecuteRequest  The request to execute.
         *
         * @return  A JQueryPromise&lt;T&gt;
         */
        static Execute<T extends Soap.ExecuteResponse>(ExecuteRequest: Soap.ExecuteRequest): JQueryPromise<T>;
        static Execute<T extends Soap.ExecuteResponse>(Execute: string | Soap.ExecuteRequest): JQueryPromise<T> {
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

        /**
         * Associates 2 entities that are linked by a N:N relationship in CRM.
         *
         * @param   {Soap.EntityReference}  Moniker1    The EntityReference of the first entity.
         * @param   {Soap.EntityReference}  Moniker2    The EntityReference of the second entity.
         * @param   {string}    RelationshipName        Name of the N:N relationship (lowercase).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static Associate(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string): JQueryPromise<Soap.SoapResponse> {
            var AssociateRequest = new Soap.AssociateRequest(Moniker1, Moniker2, RelationshipName);
            return $.Deferred<Soap.ExecuteResponse>(function (dfd) {
                var Request = Soap.Execute<Soap.ExecuteResponse>(AssociateRequest);
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Disassociates 2 entities that are linked by a N:N relationship in CRM.
         *
         * @param   {Soap.EntityReference}  Moniker1    The EntityReference of the first entity.
         * @param   {Soap.EntityReference}  Moniker2    The EntityReference of the second entity.
         * @param   {string}    RelationshipName        Name of the N:N relationship (lowercase).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static Disassociate(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string): JQueryPromise<Soap.SoapResponse> {
            var DisassociateRequest = new Soap.DisassociateRequest(Moniker1, Moniker2, RelationshipName);
            return $.Deferred<Soap.ExecuteResponse>(function (dfd) {
                var Request = Soap.Execute<Soap.ExecuteResponse>(DisassociateRequest);
                Request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                Request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Sets a state.
         *
         * @param   {Soap.EntityReference}  EntityMoniker   The EntityReference of the record.
         * @param   {number}    State                       The state (statecode).
         * @param   {number}    Status                      The status (statuscode).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static SetState(EntityMoniker: Soap.EntityReference, State: number, Status: number): JQueryPromise<Soap.SoapResponse>;

        /**
         * Sets a state.
         *
         * @param   {Soap.EntityReference}  EntityMoniker   The EntityReference of the record.
         * @param   {Soap.OptionSetValue}   State           The state (statecode).
         * @param   {Soap.OptionSetValue}   Status          The status (statuscode).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static SetState(EntityMoniker: Soap.EntityReference, State: Soap.OptionSetValue, Status: Soap.OptionSetValue): JQueryPromise<Soap.SoapResponse>;
        static SetState(EntityMoniker: Soap.EntityReference, State: number| Soap.OptionSetValue, Status: number| Soap.OptionSetValue): JQueryPromise<Soap.SoapResponse> {
            var SetStateRequest: Soap.SetStateRequest;
            if (typeof State === "number" && typeof Status === "number") { SetStateRequest = new Soap.SetStateRequest(EntityMoniker, State, Status); }
            else { SetStateRequest = new Soap.SetStateRequest(EntityMoniker, <Soap.OptionSetValue> State, <Soap.OptionSetValue>Status); }
            return $.Deferred<Soap.ExecuteResponse>(function (dfd) {
                var Request = Soap.Execute<Soap.ExecuteResponse>(SetStateRequest);
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
                Request.done(function (data: any, textStatus: string, xhr: JQueryXHR) {
                    var sr: any = new Soap.SoapResponse(xhr.responseText);
                    dfd.resolve(<T>sr, textStatus, xhr);
                });
                Request.fail(function (xhr: JQueryXHR, textStatus: string, exception) {
                    var ErrorText = typeof exception === "undefined" ? textStatus : exception;
                    var fault = new Soap.FaultResponse(xhr.responseText);
                    //if (fault && fault.Fault && fault.Fault.FaultString) {
                    //    ErrorText += (": " + fault.Fault.FaultString);
                    //}
                    dfd.reject(fault);
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
        export interface ISerializable {
            Serialize(): string;
        }
        export class Entity implements ISerializable {

            constructor();
            /**
             * Constructor.
             *
             * @param   {string}    LogicalName Logical Name of the entity.
             */
            constructor(LogicalName: string);

            /**
             * Constructor.
             *
             * @param   {string}    LogicalName Logical Name of the entity.
             * @param   {string}    Id          The GUID of the existing record.
             */
            constructor(LogicalName: string, Id: string);
            constructor(public LogicalName?: string, public Id?: string) {
                this.Attributes = new AttributeCollection();
                if (!Id) {
                    this.Id = "00000000-0000-0000-0000-000000000000";
                }
            }

            /**
             * The attributes of the Entity.  Use the following notation for setting or retrieving an
             * attribute: Entity.Attributes["attributename"] = ...
             */
            Attributes: AttributeCollection;
            FormattedValues: StringDictionary<string>;
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
        export class ParameterCollection {
            [index: string]: ISerializable;
        }
        export class StringDictionary<T> {
            [index: string]: T;
        }
        export class EntityCollection implements ISerializable {
            Items: Array<Entity> = [];
            Serialize(): string {
                var XML = "";
                $.each(this.Items, function (index, Entity) {
                    XML += "<a:Entity>" + Entity.Serialize() + "</a:Entity>";
                });
                return XML;
            }
        }
        export class ColumnSet implements ISerializable {

            /**
             * Constructor.
             *
             * @param   {Array{string}} Columns to be retrieved.
             */
            constructor(Columns: Array<string>);

            /**
             * Constructor.
             *
             * @param   {boolean}   AllColumns  'true' to return all available columns, otherwise 'false'. Default is 'false'
             */
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

            Serialize(): string {
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
        export class AttributeValue implements ISerializable {
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
            constructor();

            /**
             * Constructor.
             *
             * @param   {string}    Id          The GUID of the record.
             * @param   {string}    LogicalName Logical Name of the entity.
             */
            constructor(Id: string, LogicalName: string);

            /**
             * Constructor.
             *
             * @param   {string}    Id          The GUID of the record.
             * @param   {string}    LogicalName Logical Name of the entity.
             * @param   {string}    Name        The 'Display Name' of the record.
             */
            constructor(Id: string, LogicalName: string, Name: string);
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
            constructor(public Value?: string) {
                super(Value, AttributeType.Guid);
                if (!Value) {
                    this.Value = "00000000-0000-0000-0000-000000000000";
                }
            }
            GetValueType(): string { return "e:guid"; }
        }
        export class NodeValue {
            constructor(public Name: string, public Value: any) { }
        }
        export class Collection {
            constructor(public Items?: any) { }
        }

        export class SoapResponse {
            constructor(ResponseXML: string) {
                SoapResponse.ParseResult(ResponseXML, this);
            }
            static ParseResult(ResponseXML: string, ParentObject: SoapResponse): void {
                var XMLDocTemp = $.parseXML(ResponseXML);
                var MainElement = XML.ParseNode(XMLDocTemp.firstChild);

                while (MainElement && MainElement.Children && MainElement.Children[0] && (MainElement.Name == "Envelope" || MainElement.Name == "Body" || MainElement.Name == "ExecuteResponse")) {
                    MainElement = MainElement.Children[0];
                }
                var MainXRMObject = new XRMObject(MainElement, null);
                $.each(MainElement.Children, function (i, ChildNode) {
                    var XRMBaseObject = new XRMObject(ChildNode, MainXRMObject);
                    SoapResponse.ParseXRMBaseObject(ParentObject, XRMBaseObject);
                });
            }

            static ParseXRMBaseObject(ParentObject: any, XRMBaseObject: XRMObject): XRMObject {
                var CurrentObject = SoapResponse.GetObjectFromBaseXRMObject(XRMBaseObject);
                if (CurrentObject) {
                    if (ParentObject instanceof Array) {
                        ParentObject.push(CurrentObject);
                    }
                    else {
                        if (ParentObject.hasOwnProperty(XRMBaseObject.Name)) {
                            try {
                                //This object already has the property - try to parse it into the correct type
                                if (typeof ParentObject[XRMBaseObject.Name] === "number") {
                                    ParentObject[XRMBaseObject.Name] = parseFloat(CurrentObject);
                                }
                                else if (typeof ParentObject[XRMBaseObject.Name] === "boolean") {
                                    ParentObject[XRMBaseObject.Name] = CurrentObject == "true" ? true : false;
                                }
                                else if (typeof ParentObject[XRMBaseObject.Name] === "dateTime") {
                                    ParentObject[XRMBaseObject.Name] = new Date(CurrentObject);
                                }
                                else {
                                    ParentObject[XRMBaseObject.Name] = CurrentObject;
                                }
                            }
                            catch (e) {
                                ParentObject[XRMBaseObject.Name] = CurrentObject;
                            }
                        }
                        else {
                            ParentObject[XRMBaseObject.Name] = CurrentObject;
                        }
                    }
                    var ChildObjectsToProcess = new Array<XRMObject>();
                    if (XRMBaseObject.Value instanceof Array) {
                        $.each(XRMBaseObject.Value, function (index, Child: XRMObject) {
                            ChildObjectsToProcess.push(Child);
                        });
                    }
                    else if (XRMBaseObject.Value && !(typeof XRMBaseObject.Value === "string")) {
                        var Child: XRMObject = XRMBaseObject.Value;
                        ChildObjectsToProcess.push(Child);
                    }

                    if (ChildObjectsToProcess.length > 0) {
                        $.each(ChildObjectsToProcess, function (i, Child) {
                            var ChildObject = SoapResponse.ParseXRMBaseObject(CurrentObject, Child);
                            if (XRMBaseObject.Name == "Results") {
                                var ResultValue = ChildObject.Value;
                                if (!ResultValue) { ResultValue = ChildObject; }
                                ParentObject[Child.Name] = ResultValue;
                            }
                        });
                    }

                    //Populate the 'FormattedValues' of each of the attributes
                    if (CurrentObject instanceof Soap.Entity && (<Soap.Entity>CurrentObject).FormattedValues) {
                        for (var AttributeName in (<Soap.Entity>CurrentObject).FormattedValues) {
                            (<Soap.Entity>CurrentObject).Attributes[AttributeName].FormattedValue = (<Soap.Entity>CurrentObject).FormattedValues[AttributeName];
                        }
                    }

                }
                return CurrentObject;
            }
            static GetObjectFromBaseXRMObject(XRMBaseObject: XRMObject): any {
                var CurrentObject = null;
                switch (XRMBaseObject.TypeName) {
                    case "Object":
                    case "Array":
                    case "KeyValueArray":
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
                            case "RetrieveMultipleResult":
                                CurrentObject = new Soap.RetrieveMultipleResult();
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
                    case "OrganizationResponseCollection":
                        CurrentObject = new Array<ExecuteMultipleResponseItem>();
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
                    case "AccessRights":
                        CurrentObject = new Array<AccessRights>();
                        var Rights: Array<string> = XRMBaseObject.Value.split(" ");
                        $.each(Rights, function (i, Right) {
                            (<Array<AccessRights>>CurrentObject).push(AccessRights[Right]);
                        });
                        break;
                    default:
                        CurrentObject = {};
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

        class XRMObject {
            constructor(Element: XML.XMLElement, Parent: XRMObject) {
                this.Parent = Parent;
                if (Element.Name && Element.Name.indexOf("KeyValuePair") == 0) {
                    //This is a key value pair - so the Name and Value are actually the first and second children
                    Parent.TypeName = "KeyValueArray";
                    this.Name = Element.Children[0].Text;
                    this.IsKeyValuePair = true;
                    if (Element.Children[1].Text) {
                        this.Value = Element.Children[1].Text;
                        this.TypeName = Element.Children[1].TypeName;
                    }
                    else if (Element.Children[1].Children.length == 1) {
                        if (Element.Children[1].Name === "value" && (Element.Children[1].Children[0])) { this.Value = new XRMObject(Element.Children[1].Children[0], this); }
                        else { this.Value = new XRMObject(Element.Children[1], this); }
                        this.TypeName = Element.Children[1].TypeName;
                    }
                    else {
                        this.Value = [];
                        var _Value = this.Value;
                        this.TypeName = Element.Children[1].TypeName != "Object" ? Element.Children[1].TypeName : "Array";
                        $.each(Element.Children[1].Children, function (index, ChildNode) {
                            _Value.push(new XRMObject(ChildNode, this));
                        });
                    }
                }
                else {
                    //Determine the type of object by examining the children
                    this.Name = Element.Name;
                    if (Element.Text || Element.Children.length == 0) {
                        this.Value = Element.Text;
                        this.TypeName = Element.TypeName ? Element.TypeName : "Object";
                    }
                    else if (Element.Children.length == 1) {
                        this.Value = new XRMObject(Element.Children[0], this);
                        this.TypeName = "Object";
                    }
                    else {
                        this.Value = [];
                        var _Value = this.Value;
                        this.TypeName = "Array"; //Element.TypeName != "Object" ? Element.TypeName : "Array";
                        $.each(Element.Children, function (index, ChildNode) {
                            _Value.push(new XRMObject(ChildNode, this));
                        });
                    }
                }
            }
            Name: string;
            Value: any;
            TypeName: string;
            IsKeyValuePair: boolean = false;
            Parent: XRMObject;
        }



        export class SDKResponse { }
        export class CreateSoapResponse extends SDKResponse {
            /** The GUID of the newly created record. */
            CreateResult: string;
        }
        export class UpdateSoapResponse extends SDKResponse { }
        export class DeleteSoapResponse extends SDKResponse { }
        export class RetrieveSoapResponse extends SDKResponse {
            /** The retrieved Soap.Entity */
            RetrieveResult: Entity;
        }
        export class RetrieveMultipleSoapResponse extends SDKResponse {
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
            constructor(ResponseXML: string) { super(ResponseXML); }
            detail: FaultDetail;
            faultcode: string;
            faultstring: string;
        }
        export class FaultDetail {
            OrganizationServiceFault: Fault;
        }
        export class Fault {
            ErrorCode: string;
            ErrorDetails: any;
            InnerFault: Fault;
            Message: string;
            Timestamp: string;
        }

    }

    export module Soap.Query {
        export class QueryExpression {

            /**
             * Constructor.
             *
             * @param   {string}    EntityName  Logical Name of the entity.
             */
            constructor(public EntityName?: string) { }

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
                Data += "<a:ColumnSet>" + this.Columns.Serialize() + "</a:ColumnSet>";

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

            /**
             * Constructor.
             *
             * @param   {LogicalOperator}   FilterOperator  The filter operator.
             */
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

            /**
             * Constructor.
             *
             * @param   {string}    AttributeName       Name of the attribute.
             * @param   {ConditionOperator} Operator    The comparison operator.
             * @param   {AttributeValue}    Value       The value to be compared to.
             */
            constructor(AttributeName: string, Operator: ConditionOperator, Value?: AttributeValue);

            /**
             * Constructor.
             *
             * @param   {string}    AttributeName       Name of the attribute.
             * @param   {ConditionOperator} Operator    The comparison operator.
             * @param   {Array{AttributeValue}} Values  The values to be compared to.
             */
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
            Count: number = 0;
            PageNumber: number = 1;
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

            /**
             * Constructor.
             *
             * @param   {string}    AttributeName   Name of the attribute.
             * @param   {OrderType} OrderType       Type of the order.
             */
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

            /**
             * Constructor.
             *
             * @param   {string}    LinkFromEntityName                  LogicalName of the link from entity.
             * @param   {string}    LinkToEntityName                    LogicalName of the link to entity.
             * @param   {string}    LinkFromAttributeName               LogicalName of the link from
             *                                                          attribute.
             * @param   {string}    LinkToAttributeName                 LogicalName of the link to attribute.
             * @param   {JoinOperator}  optional public JoinOperator    JoinOperator    The optional join
             *                                                          operator.
             */

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
                Data += this.Columns.Serialize();
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
    export module XML {
        export class XMLElement {
            Name: string;
            Text: string;
            TypeName: string;
            Children: Array<XMLElement> = new Array<XMLElement>();
            Parent: XMLElement;
        }
    }
    export module Soap {
        //These are all the Organization Requests and Responses
        export class ExecuteResponse extends SoapResponse {
            Results: AttributeCollection;
        }

        export class ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {string}    RequestName             Name of the request.
             * @param   {string}    RequestType             Type of the request. Required if Request type
             *                                              differs from pattern: "[RequestName]Type" or if
             *                                              the RequestType namesapce differs from "g" in the
             *                                              global list of namespaces.
             * @param   {ParameterCollection}   Parameters  Parameters for the operation.
             */

            constructor(public RequestName: string, public RequestType?: string, public Parameters?: ParameterCollection) {
                if (!Parameters || Parameters == null) {
                    this.Parameters = new AttributeCollection();
                }
                if (!RequestType || RequestType == null) {
                    this.RequestType = "g:" + this.RequestName + "Request";
                }
            }
            RequestId: string;
            ExecuteRequestType: string = "request";
            IncludeExecuteHeader: boolean = true;
            Serialize(): string {
                var XML = "";
                if (this.IncludeExecuteHeader) { XML += "<Execute" + Soap.GetNameSpacesXML() + ">"; }
                XML += "<" + this.ExecuteRequestType + " i:type='" + this.RequestType + "'><a:Parameters>";
                for (var ParameterName in this.Parameters) {
                    var Parameter = this.Parameters[ParameterName];

                    XML += "<a:KeyValuePairOfstringanyType>";
                    XML += "<b:key>" + ParameterName + "</b:key>";
                    if (Parameter instanceof Soap.Entity) {
                        XML += "<b:value i:type=\"a:Entity\">" + Parameter.Serialize() + "</b:value>";
                    }
                    else {
                        XML += Parameter.Serialize();
                    }
                    XML += "</a:KeyValuePairOfstringanyType>";

                }
                XML += "</a:Parameters>";
                XML += "<a:RequestId i:nil = \"true\" />";
                XML += "<a:RequestName>" + this.RequestName + "</a:RequestName>";
                XML += "</" + this.ExecuteRequestType + ">";
                if (this.IncludeExecuteHeader) { XML += "</Execute>"; }
                return XML;
            }
        }
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
        export class UpdateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Entity}    Target  Entity to be updated.
             */

            constructor(Target: Entity) {
                super("Update", "a:UpdateRequest");
                this.Parameters["Target"] = Target;
            }
        }
        export class UpdateResponse extends ExecuteResponse { }
        export class DeleteRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   Target  Entity to be deleted.
             */

            constructor(Target: EntityReference) {
                super("Delete", "a:DeleteRequest");
                this.Parameters["Target"] = Target;
            }
        }
        export class DeleteResponse extends ExecuteResponse { }
        export class AssociateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  Moniker1    The first entity.
             * @param   {Soap.EntityReference}  Moniker2    The second entity.
             * @param   {string}    RelationshipName        LogicalName of the N:N relationship.
             */

            constructor(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string) {
                super("AssociateEntities");
                this.Parameters["Moniker1"] = Moniker1;
                this.Parameters["Moniker2"] = Moniker2;
                this.Parameters["RelationshipName"] = new Soap.StringValue(RelationshipName);
            }
        }
        export class DisassociateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  Moniker1    The first entity.
             * @param   {Soap.EntityReference}  Moniker2    The second entity.
             * @param   {string}    RelationshipName        LogicalName of the N:N relationship.
             */

            constructor(Moniker1: Soap.EntityReference, Moniker2: Soap.EntityReference, RelationshipName: string) {
                super("DisassociateEntities");
                this.Parameters["Moniker1"] = Moniker1;
                this.Parameters["Moniker2"] = Moniker2;
                this.Parameters["RelationshipName"] = new Soap.StringValue(RelationshipName);
            }
        }
        export class SetStateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  EntityMoniker   The entity.
             * @param   {number}    State                       The state (statecode).
             * @param   {number}    Status                      The status (statuscode).
             */

            constructor(EntityMoniker: Soap.EntityReference, State: number, Status: number);

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  EntityMoniker   The entity.
             * @param   {Soap.OptionSetValue}   State           The state (statecode).
             * @param   {Soap.OptionSetValue}   Status          The status (statuscode).
             */

            constructor(EntityMoniker: Soap.EntityReference, State: Soap.OptionSetValue, Status: Soap.OptionSetValue);
            constructor(EntityMoniker: Soap.EntityReference, State: number| Soap.OptionSetValue, Status: number| Soap.OptionSetValue) {
                super("SetState");
                var StateOptionSet: Soap.OptionSetValue;
                if (typeof State === "number") { StateOptionSet = new Soap.OptionSetValue(State); }
                else { StateOptionSet = State; }
                var StatusOptionSet: Soap.OptionSetValue;
                if (typeof Status === "number") { StatusOptionSet = new Soap.OptionSetValue(Status); }
                else { StatusOptionSet = Status; }
                this.Parameters = new Soap.AttributeCollection();
                this.Parameters["EntityMoniker"] = EntityMoniker;
                this.Parameters["State"] = StateOptionSet;
                this.Parameters["Status"] = StatusOptionSet;
            }
        }
        export class WhoAmIRequest extends ExecuteRequest {
            constructor() {
                super("WhoAmI");
            }
        }
        export class WhoAmIResponse extends ExecuteResponse {
            constructor(ResponseXML: string) { super(ResponseXML); }
            public UserId: string;
            public BusinessUnitId: string;
            public OrganizationId: string;
        }
        export class AssignRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   Assignee    The entity (user or team) that the record will be assigned to.
             * @param   {EntityReference}   Target      The record to be assigned.
             */

            constructor(Assignee: EntityReference, Target: EntityReference) {
                super("Assign");
                this.Parameters["Assignee"] = Assignee;
                this.Parameters["Target"] = Target;
            }
        }
        export class AssignResponse extends ExecuteResponse {
            constructor(ResponseXML: string) { super(ResponseXML); }
        }
        type OrganizationRequestCollection = Array<Soap.ExecuteRequest>;
        export class ExecuteMultipleRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {OrganizationRequestCollection} Requests    The requests to be executed.
             * @param   {ExecuteMultipleSettings}   Settings        Options for controlling the operation.
             */

            constructor(public Requests?: OrganizationRequestCollection, public Settings?: ExecuteMultipleSettings) {
                super("ExecuteMultiple", "a:ExecuteMultipleRequest");
                if (!Requests || Requests == null) { this.Requests = []; }
                if (!Settings || Settings == null) { this.Settings = new Soap.ExecuteMultipleSettings(); }
            }
            Serialize(): string {
                var XML = "<Execute" + Soap.GetNameSpacesXML() + ">";
                XML += "<request i:type='" + this.RequestType + "'><a:Parameters>";
                XML += "<a:KeyValuePairOfstringanyType><b:key>Requests</b:key><b:value i:type=\"l:OrganizationRequestCollection\">";
                $.each(this.Requests, function (i, Request) {
                    Request.ExecuteRequestType = "l:OrganizationRequest";
                    Request.IncludeExecuteHeader = false;
                    XML += Request.Serialize();
                });
                XML += "</b:value></a:KeyValuePairOfstringanyType>";
                XML += "<a:KeyValuePairOfstringanyType><b:key>Settings</b:key><b:value i:type=\"l:ExecuteMultipleSettings\">";
                XML += this.Settings.Serialize();
                XML += "</b:value></a:KeyValuePairOfstringanyType>";
                XML += "</a:Parameters>";
                XML += "<a:RequestId i:nil = \"true\" />";
                XML += "<a:RequestName>ExecuteMultiple</a:RequestName>";
                XML += "</request></Execute>";
                return XML;
            }
        }
        export class ExecuteMultipleSettings {

            /**
             * Constructor.
             *
             * @param   {boolean}   optional public ContinueOnError true to continue on error otherwise false.
             * @param   {boolean}   optional public ReturnResponses true to return the individual responses otherwise false.
             */

            constructor(public ContinueOnError: boolean = false, public ReturnResponses: boolean = false) { }
            Serialize(): string {
                var XML = "<l:ContinueOnError>" + this.ContinueOnError.toString() + "</l:ContinueOnError>";
                XML += "<l:ReturnResponses>" + this.ReturnResponses.toString() + "</l:ReturnResponses>";
                return XML;
            }
        }
        export class ExecuteMultipleResponse extends ExecuteResponse {
            Responses: ExecuteMultipleResponseItemCollection;
            ResponseName: string;
        }
        type ExecuteMultipleResponseItemCollection = Array<ExecuteMultipleResponseItem>;

        export class ExecuteMultipleResponseItem {
            Fault: Soap.Fault;
            RequestIndex: number;
            Response: Soap.SoapResponse;
        }
        export class GrantAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   Target      Record to receive access to.
             * @param   {EntityReference}   Principal   The user or team to be granted access.
             * @param   {AccessRights}  AccessMask      The type of access to be granted.
             */

            constructor(Target: EntityReference, Principal: EntityReference, AccessMask: AccessRights) {
                super("GrantAccess");
                this.Parameters["Target"] = Target;
                this.Parameters["PrincipalAccess"] = new PrincipalAccess(AccessMask, Principal);
            }
        }
        export class GrantAccessResponse extends ExecuteResponse { }
        export class ModifyAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   Target      Record for which access will be modified.
             * @param   {EntityReference}   Principal   The user or team for which the access will be modified.
             * @param   {AccessRights}  AccessMask      The type of access to be granted.
             */

            constructor(Target: EntityReference, Principal: EntityReference, AccessMask: AccessRights) {
                super("ModifyAccess");
                this.Parameters["Target"] = Target;
                this.Parameters["PrincipalAccess"] = new PrincipalAccess(AccessMask, Principal);
            }
        }
        export class ModifyAccessResponse extends ExecuteResponse { }
        export class RevokeAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   Target  Record for which access will be removed.
             * @param   {EntityReference}   The user or team to remove access from.
             */

            constructor(Target: EntityReference, Revokee: EntityReference) {
                super("RevokeAccess");
                this.Parameters["Target"] = Target;
                this.Parameters["Revokee"] = Revokee;
            }
        }
        export class RetrievePrincipleAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   Target      Record for which the access will be retrieved.
             * @param   {EntityReference}   Principal   The user or team for which the access will be checked.
             */

            constructor(Target: EntityReference, Principal: EntityReference) {
                super("RetrievePrincipalAccess");
                this.Parameters["Target"] = Target;
                this.Parameters["Principal"] = Principal;
            }
        }
        export class RetrievePrincipleAccessResponse extends ExecuteResponse {
            AccessRights: Array<AccessRights>;
        }
        export class RevokeAccessResponse extends ExecuteResponse { }
        export class PrincipalAccess implements Soap.ISerializable {
            constructor(public AccessMask: AccessRights, public Principal: EntityReference) { }
            Serialize(): string {
                var XML = "<b:value i:type=\"g:PrincipalAccess\"><g:AccessMask>" + AccessRights[this.AccessMask].toString() + "</g:AccessMask>";
                XML += "<g:Principal><a:Id>" + Soap.Entity.EncodeValue(this.Principal.Id) + "</a:Id>";
                XML += "<a:LogicalName>" + Soap.Entity.EncodeValue(this.Principal.LogicalName) + "</a:LogicalName><a:Name i:nil=\"true\" />";
                XML += "</g:Principal></b:value>";
                return XML;
            }
        }

        export enum AccessRights {
            None = 0,
            ReadAccess = 1,
            WriteAccess = 2,
            AppendAccess = 4,
            AppendToAccess = 16,
            CreateAccess = 32,
            DeleteAccess = 65536,
            ShareAccess = 262144,
            AssignAccess = 524288
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