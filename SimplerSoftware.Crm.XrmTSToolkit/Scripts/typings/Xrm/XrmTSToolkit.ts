/// <reference path="../jquery/jquery.d.ts" />
/// <reference path="xrm.d.ts" />

/**
* MSCRM 2011, 2013, 2015 Service Toolkit for TypeScript
* @author Steven Rasmussen
* @current version : 0.8.0
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
*   RetrieveMultiple and Fetch methods are now able to return all results (>5000 records) - set the 'Soap.RetrieveAllEntities static property to true', default is false
*   Moved XML parsing to XrmTSToolkit.XML class and module
*   Minor bug fixes
********************************************
* Version : 0.8.0
* Date: February 2, 2016
*   Updated to TypeScript v1.7.6
*   Fixed issue with serialization of columns specified in a columnset
*   Added RetrieveEntityMetadataRequest and tests
*   Rewrite of soap response parsing to handle the entity metadata objects and to be more generic
*   Removed unused method 'FilterNode'
*   Added the ability to just use a 'bool' or 'string' as a value instead of needing to create a 'BooleanValue' or 'StringValue' when creating a record
*   Added the private members 'id', 'entityType', and 'name' to the Entity Reference class along with getters and setters for the public properties to allow an EntityReference object to be used
*       directly with the Xrm.Page.setValue method when working with form scripts - removes the need to create a new object to pass as the value to a lookup attribute on a form.
*   Added 'EntityReferenceCollection' as a possible type
*/

module XrmTSToolkit {
    export class Common {

        /**
         * Gets server URL.
         *
         * @return  The server URL.
         */
        public static GetServerURL(): string {
            var url: string = document.location.protocol + "//" + document.location.host;
            var org: string = Xrm.Page.context.getOrgUniqueName();
            if (document.location.pathname.toUpperCase().indexOf(org.toUpperCase()) > -1) {
                url += "/" + org;
            }
            return url;
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
        static StripGUID(guid: string): string {
            return guid.replace("{", "").replace("}", "");
        }
    }
    export class XML {
        public static ParseNode(node: Node): XML.XMLElement {
            var element = new XML.XMLElement();
            var nodeName = $(node).prop("nodeName");
            var timerName = "  Get Contents of node: " + nodeName;
            console.time(timerName);
            if (nodeName.indexOf(":") > 0) {
                nodeName = nodeName.substring(nodeName.indexOf(":") + 1);
            }
            element.Name = nodeName;
            if (element.Name == "AttributeMetadata") {
                var Thing = "";
            }
            element.TypeName = $(node).attr("i:type");
            if (element.TypeName && element.TypeName.indexOf(":") > 0) {
                element.TypeName = element.TypeName.substring(element.TypeName.indexOf(":") + 1);
            }
            var childNodes = $(node).contents();

            if (childNodes && childNodes.length > 0) {
                for (var i = 0; i < childNodes.length; i++) {
                    var childNode = childNodes[i];
                    var childNodeType = $(childNode).prop("nodeType");
                    if (childNodeType && childNodeType != 3) {
                        var childElement = XML.ParseNode(childNode);
                        childElement.Parent = element;
                        element.Children.push(childElement);

                    }
                }
            }
            if (element.Children.length == 0) {
                element.Text = $(node).text();
            }
            console.timeEnd(timerName);
            return element;
        }
        //public static FilterNode = function (name) {
        //    return this.find('*').filter(function () {
        //        return this.nodeName === name;
        //    });
        //};
    }
    export class Soap {

        /** "true" to retrieve all entities for RetrieveMultiple and Fetch operations, otherwise "false" */
        public static RetrieveAllEntities: boolean = false;

        /**
         * Creates a new record in CRM
         *
         * @param   {Soap.Entity}   entity  The Soap.Entity to create in CRM.
         *
         * @return  A JQueryPromise&lt;Soap.CreateSoapResponse&gt;
         */
        static Create(entity: Soap.Entity): JQueryPromise<Soap.CreateSoapResponse> {
            var requestBody = "<entity>" + entity.Serialize() + "</entity>";
            return $.Deferred<Soap.CreateSoapResponse>(function (dfd) {
                var request = Soap.DoRequest<Soap.CreateSoapResponse>(requestBody, "Create");
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Updates the given Entity in CRM.
         *
         * @param   {Soap.Entity}   entity  The Soap.Entity to update in CRM..
         *
         * @return  A JQueryPromise&lt;Soap.UpdateSoapResponse&gt;
         */
        static Update(entity: Soap.Entity): JQueryPromise<Soap.UpdateSoapResponse> {
            var requestBody = "<entity>" + entity.Serialize() + "</entity>";
            return $.Deferred<Soap.UpdateSoapResponse>(function (dfd) {
                var request = Soap.DoRequest<Soap.UpdateSoapResponse>(requestBody, "Update");
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Deletes the given Entity from CRM.
         *
         * @param   {Soap.EntityReference}  entity  The entity to delete.
         *
         * @return  A JQueryPromise&lt;Soap.UpdateSoapResponse&gt;
         */
        static Delete(entity: Soap.EntityReference): JQueryPromise<Soap.DeleteSoapResponse>;

        /**
         * Deletes this object.
         *
         * @param   {string}    entityName  Logical name of the entity to be deleted.
         * @param   {string}    entityId    GUID of the entity to be deleted.
         *
         * @return  A JQueryPromise&lt;Soap.DeleteSoapResponse&gt;
         */
        static Delete(entityName: string, entityId: string): JQueryPromise<Soap.DeleteSoapResponse>;
        static Delete(entity: string | Soap.EntityReference, entityId?: string): JQueryPromise<Soap.DeleteSoapResponse> {
            var entityName = "";
            if (typeof entity === "string") {
                entityName = entityName;
            }
            else if (entity instanceof Soap.EntityReference) {
                entityName = entity.LogicalName;
                entityId = entity.Id;
            }
            var xml = "<entityName>" + entityName + "</entityName>";
            xml += "<id>" + entityId + "</id>";
            return $.Deferred<Soap.DeleteSoapResponse>(function (dfd) {
                var request = Soap.DoRequest<Soap.DeleteSoapResponse>(xml, "Delete");
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Retrieves the specified record from CRM.
         *
         * @param   {string}    id                  The GUID of the entity to retrieve.
         * @param   {string}    entityLogicalName   Entity logical name.
         * @param   {Soap.ColumnSet}    columnSet   Columns to retrieve as part of the entity.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveSoapResponse&gt;
         */
        static Retrieve(id: string, entityLogicalName: string, columnSet?: Soap.ColumnSet): JQueryPromise<Soap.RetrieveSoapResponse> {
            id = Common.StripGUID(id);
            if (!columnSet || columnSet == null) { columnSet = new Soap.ColumnSet(false); }
            var msgBody = "<entityName>" + entityLogicalName + "</entityName><id>" + id + "</id><columnSet>" + columnSet.Serialize() + "</columnSet>";
            return $.Deferred<Soap.RetrieveSoapResponse>(function (dfd) {
                var request = Soap.DoRequest<Soap.RetrieveSoapResponse>(msgBody, "Retrieve");
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Retrieves records from CRM that match the provided criteria.
         *
         * @param   {Soap.Query.QueryExpression}    query   The query specifying the criteria and other parameters.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveMultipleSoapResponse&gt;
         */
        static RetrieveMultiple(query: Soap.Query.QueryExpression): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var request = Soap.DoRequest<Soap.RetrieveMultipleSoapResponse>(query.serialize(), "RetrieveMultiple");
                request.done(function (data, result, xhr) {
                    if (data.RetrieveMultipleResult.MoreRecords == true && Soap.RetrieveAllEntities == true) {
                        query.PageInfo.PagingCookie = Soap.Entity.EncodeValue(data.RetrieveMultipleResult.PagingCookie);
                        query.PageInfo.PageNumber = (query.PageInfo.PageNumber + 1);
                        var additionalRequest = Soap.RetrieveMultiple(query);
                        additionalRequest.done(function (data1, result, xhr) {
                            $.each(data1.RetrieveMultipleResult.Entities, function (i, Entity) {
                                data.RetrieveMultipleResult.Entities.push(Entity);
                            });
                            dfd.resolve(data);
                        });
                        additionalRequest.fail(function (result) {
                            dfd.reject(result);
                        });
                    }
                    else {
                        dfd.resolve(data);
                    }
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Retrieves related many to many.
         *
         * @param   {string}    linkFromEntityName                          Entity logical name of the known entity.
         * @param   {string}    linkFromEntityId                            Id of the known entity
         * @param   {string}    linkToEntityName                            Entity logical name of the records to be retrieved.
         * @param   {string}    intermediateTableName                       Name of the intermediate table of the N:N relationship.
         * @param   {Soap.ColumnSet}    optional columns                    The columns to be retrieved on the returned records.
         * @param   {Array{Soap.Query.OrderExpression}} optional sortOrders Sort order of the returned records.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveMultipleSoapResponse&gt;
         */
        static RetrieveRelatedManyToMany(
            linkFromEntityName: string,
            linkFromEntityId: string,
            linkToEntityName: string,
            intermediateTableName: string,
            columns: Soap.ColumnSet = new Soap.ColumnSet(false),
            sortOrders: Array<Soap.Query.OrderExpression> = []): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {

            var condition = new Soap.Query.ConditionExpression(linkFromEntityName + "id", Soap.Query.ConditionOperator.Equal, new Soap.GuidValue(linkFromEntityId));
            var linkEntity1 = new Soap.Query.LinkEntity(linkToEntityName, intermediateTableName, linkToEntityName + "id", linkToEntityName + "id", Soap.Query.JoinOperator.Inner);
            linkEntity1.LinkCriteria = new Soap.Query.FilterExpression();
            linkEntity1.LinkCriteria.AddCondition(condition);

            var linkEntity2 = new Soap.Query.LinkEntity(intermediateTableName, linkFromEntityName, linkFromEntityName + "id", linkFromEntityName + "id", Soap.Query.JoinOperator.Inner);
            linkEntity2.LinkCriteria = new Soap.Query.FilterExpression();
            linkEntity2.LinkCriteria.AddCondition(condition);
            linkEntity1.LinkEntities.push(linkEntity2);

            var query = new Soap.Query.QueryExpression(linkToEntityName);
            query.Columns = columns;
            query.LinkEntities.push(linkEntity1);
            query.Orders = sortOrders;

            return Soap.RetrieveMultiple(query);
        }

        /**
         * Performs a 'Fetch' operation, similar to a 'RetrieveMultiple' but using FetchXML as the query string.
         *
         * @param   {string}    fetchXml    The FetchXML query.
         *
         * @return  A JQueryPromise&lt;Soap.RetrieveMultipleSoapResponse&gt;
         */
        static Fetch(fetchXml: string): JQueryPromise<Soap.RetrieveMultipleSoapResponse> {
            var xmlDoc = $.parseXML(fetchXml);
            //First decode all invalid characters
            fetchXml = Soap.DecodeFetchXML(fetchXml);

            //now re-encode all invalid characters
            fetchXml = Soap.EncodeFetchXML(fetchXml);

            var msgBody = "<query i:type='a:FetchExpression'><a:Query>" + fetchXml + "</a:Query></query>";
            return $.Deferred<Soap.RetrieveMultipleSoapResponse>(function (dfd) {
                var request = Soap.DoRequest<Soap.RetrieveMultipleSoapResponse>(msgBody, "RetrieveMultiple");
                request.done(function (data, result, xhr) {
                    if (data.RetrieveMultipleResult.MoreRecords == true && Soap.RetrieveAllEntities == true) {
                        var pagingCookie = xmlDoc.firstChild.attributes.getNamedItem("paging-cookie");
                        if (!pagingCookie) { pagingCookie = xmlDoc.createAttribute("paging-cookie"); }
                        else { xmlDoc.firstChild.attributes.removeNamedItem("paging-cookie"); }

                        pagingCookie.value = Soap.Entity.EncodeValue(data.RetrieveMultipleResult.PagingCookie);
                        xmlDoc.firstChild.attributes.setNamedItem(pagingCookie);

                        var pageNumber = xmlDoc.firstChild.attributes.getNamedItem("page");
                        var currentPageNumber: number = 1;
                        if (!pageNumber) {
                            pageNumber = xmlDoc.createAttribute("page");
                        }
                        else {
                            currentPageNumber = parseInt(pageNumber.value);
                            xmlDoc.firstChild.attributes.removeNamedItem("page");
                        }
                        currentPageNumber += 1;
                        pageNumber.value = currentPageNumber.toString();
                        xmlDoc.firstChild.attributes.setNamedItem(pageNumber);
                        var additionalRequest = Soap.Fetch(Soap.XMLtoString(xmlDoc));
                        additionalRequest.done(function (data1, result, xhr) {
                            $.each(data1.RetrieveMultipleResult.Entities, function (i, Entity) {
                                data.RetrieveMultipleResult.Entities.push(Entity);
                            });
                            dfd.resolve(data);
                        });
                        additionalRequest.fail(function (result) {
                            dfd.reject(result);
                        });
                    }
                    else {
                        dfd.resolve(data);
                    }
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        private static DecodeFetchXML(fetchXML: string): string {
            fetchXML = fetchXML.replace(/&amp;/g, "&");
            fetchXML = fetchXML.replace(/&lt;/g, "<");
            fetchXML = fetchXML.replace(/&gt;/g, ">");
            fetchXML = fetchXML.replace(/&apos;/g, "'");
            fetchXML = fetchXML.replace(/&quot;/g, "\"");
            return fetchXML;
        }
        private static EncodeFetchXML(fetchXML: string): string {
            fetchXML = fetchXML.replace(/&/g, "&amp;");
            fetchXML = fetchXML.replace(/</g, "&lt;");
            fetchXML = fetchXML.replace(/>/g, "&gt;");
            fetchXML = fetchXML.replace(/'/g, "&apos;");
            fetchXML = fetchXML.replace(/\"/g, "&quot;");
            return fetchXML;
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
         * @param   {string}    executeXML  contents of the '&lt;Execute&gt;' tags,  INCLUDING the '&lt;
         *                                  Execute&gt;' tags.
         *
         * @return  A JQueryPromise&lt;T&gt;
         */
        static Execute<T>(executeXML: string): JQueryPromise<T>;

        /**
         * Executes the specified request.
         *
         * @tparam  T   Type of the soap response to be returned.
         * @param   {Soap.ExecuteRequest}   executeRequest  The request to execute.
         *
         * @return  A JQueryPromise&lt;T&gt;
         */
        static Execute<T extends Soap.ExecuteResponse>(executeRequest: Soap.ExecuteRequest): JQueryPromise<T>;
        static Execute<T extends Soap.ExecuteResponse>(execute: string | Soap.ExecuteRequest): JQueryPromise<T> {
            return $.Deferred<T>(function (dfd) {
                var executeXML = "";
                if (typeof execute === "string") {
                    executeXML = execute;
                }
                else if (execute instanceof Soap.ExecuteRequest) {
                    executeXML = execute.Serialize();
                }
                var request = Soap.DoRequest<T>(executeXML, "Execute");
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Associates 2 entities that are linked by a N:N relationship in CRM.
         *
         * @param   {Soap.EntityReference}  moniker1    The EntityReference of the first entity.
         * @param   {Soap.EntityReference}  moniker2    The EntityReference of the second entity.
         * @param   {string}    relationshipName        Name of the N:N relationship (lowercase).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static Associate(moniker1: Soap.EntityReference, moniker2: Soap.EntityReference, relationshipName: string): JQueryPromise<Soap.SoapResponse> {
            var AssociateRequest = new Soap.AssociateRequest(moniker1, moniker2, relationshipName);
            return $.Deferred<Soap.ExecuteResponse>(function (dfd) {
                var request = Soap.Execute<Soap.ExecuteResponse>(AssociateRequest);
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Disassociates 2 entities that are linked by a N:N relationship in CRM.
         *
         * @param   {Soap.EntityReference}  moniker1    The EntityReference of the first entity.
         * @param   {Soap.EntityReference}  moniker2    The EntityReference of the second entity.
         * @param   {string}    relationshipName        Name of the N:N relationship (lowercase).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static Disassociate(moniker1: Soap.EntityReference, moniker2: Soap.EntityReference, relationshipName: string): JQueryPromise<Soap.SoapResponse> {
            var DisassociateRequest = new Soap.DisassociateRequest(moniker1, moniker2, relationshipName);
            return $.Deferred<Soap.ExecuteResponse>(function (dfd) {
                var request = Soap.Execute<Soap.ExecuteResponse>(DisassociateRequest);
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }

        /**
         * Sets a state.
         *
         * @param   {Soap.EntityReference}  entityMoniker   The EntityReference of the record.
         * @param   {number}    state                       The state (statecode).
         * @param   {number}    status                      The status (statuscode).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static SetState(entityMoniker: Soap.EntityReference, state: number, status: number): JQueryPromise<Soap.SoapResponse>;

        /**
         * Sets a state.
         *
         * @param   {Soap.EntityReference}  entityMoniker   The EntityReference of the record.
         * @param   {Soap.OptionSetValue}   state           The state (statecode).
         * @param   {Soap.OptionSetValue}   status          The status (statuscode).
         *
         * @return  A JQueryPromise&lt;Soap.SoapResponse&gt;
         */
        static SetState(entityMoniker: Soap.EntityReference, state: Soap.OptionSetValue, status: Soap.OptionSetValue): JQueryPromise<Soap.SoapResponse>;
        static SetState(entityMoniker: Soap.EntityReference, state: number | Soap.OptionSetValue, status: number | Soap.OptionSetValue): JQueryPromise<Soap.SoapResponse> {
            var setStateRequest: Soap.SetStateRequest;
            if (typeof state === "number" && typeof status === "number") { setStateRequest = new Soap.SetStateRequest(entityMoniker, state, status); }
            else { setStateRequest = new Soap.SetStateRequest(entityMoniker, <Soap.OptionSetValue>state, <Soap.OptionSetValue>status); }
            return $.Deferred<Soap.ExecuteResponse>(function (dfd) {
                var request = Soap.Execute<Soap.ExecuteResponse>(setStateRequest);
                request.done(function (data, result, xhr) {
                    dfd.resolve(data);
                });
                request.fail(function (result) {
                    dfd.reject(result);
                });
            }).promise();
        }
        private static DoRequest<T>(soapBody: string, requestType: string): JQueryPromise<T> {
            var xml = "";
            if (requestType == "Execute") {
                xml = "<s:Envelope xmlns:s = \"http://schemas.xmlsoap.org/soap/envelope/\">" +
                    "<s:Body>" + soapBody + "</s:Body></s:Envelope>";
            }
            else {
                soapBody = "<" + requestType + ">" + soapBody + "</" + requestType + ">";
                //Add in all the different namespaces
                xml = "<s:Envelope" + Soap.GetNameSpacesXML() + "><s:Body>" + soapBody + "</s:Body></s:Envelope>";
            }

            return $.Deferred<T>(function (dfd) {
                var Request = $.ajax(XrmTSToolkit.Common.GetSoapServiceURL(), {
                    data: xml,
                    type: "POST",
                    beforeSend: function (xhr: JQueryXHR) {
                        xhr.setRequestHeader("Accept", "application/xml, text/xml, */*");
                        xhr.setRequestHeader("Content-Type", "text/xml; charset=utf-8");
                        xhr.setRequestHeader("SOAPAction", "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/" + requestType);
                    }
                });
                Request.done(function (data: any, textStatus: string, xhr: JQueryXHR) {
                    var sr: Soap.SoapResponse = null;
                    var rt = xhr.responseText;
                    switch (requestType) {
                        case "Create":
                            sr = new Soap.CreateSoapResponse(rt);
                            break;
                        case "Update":
                            sr = new Soap.UpdateSoapResponse(rt);
                            break;
                        case "Delete":
                            sr = new Soap.DeleteSoapResponse(rt);
                            break;
                        case "Retrieve":
                            sr = new Soap.RetrieveSoapResponse(rt);
                            break;
                        case "RetrieveMultiple":
                            sr = new Soap.RetrieveMultipleSoapResponse(rt);
                            break;
                        case "Execute":
                            sr = new Soap.ExecuteResponse(rt);
                            break;
                        default:
                            sr = new Soap.SoapResponse(rt);
                            break;
                    }
                    sr.ParseResult();
                    dfd.resolve(<T><any>sr, textStatus, xhr);
                });
                Request.fail(function (xhr: JQueryXHR, textStatus: string, exception) {
                    var errorText = typeof exception === "undefined" ? textStatus : exception;
                    var fault = new Soap.FaultResponse(xhr.responseText);
                    fault.ParseResult();
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

            var xml = xmlns;
            for (var i in ns) {
                xml += " xmlns:" + i + "=\"" + ns[i] + "\"";
            }
            return xml;
        }
    }
    export module Soap {
        export interface ISerializable {
            Serialize(): string;
        }
        export class SoapResponse {
            constructor(public responseXML: string) { }
            ParseResult(): void {
                this.ParseResultInernal(this.responseXML);
            }
            PropertyTypes = new PropertyTypeCollection();

            protected ParseResultInernal(responseXML: string): void {
                console.time("  ParseResultInernal");
                var xmlDoc = $.parseXML(responseXML);
                console.time("  Parse Main Element");
                var mainElement = XML.ParseNode(xmlDoc.firstChild);
                console.timeEnd("  Parse Main Element");
                while (mainElement && mainElement.Children && mainElement.Children[0] && (mainElement.Name == "Envelope" || mainElement.Name == "Body" || mainElement.Name == "ExecuteResponse")) {
                    mainElement = mainElement.Children[0];
                }

                console.time("  Parse Main XRM Object");
                var mainXRMObject = new XRMObject(mainElement, null);
                console.timeEnd("  Parse Main XRM Object");

                var parentObjects = new Array<any>();
                parentObjects.push(this);
                $.each(mainElement.Children, function (i, ChildNode) {
                    console.time("  Parse Child XRM Object " + (i + 1));
                    var xrmBaseObject = new XRMObject(ChildNode, mainXRMObject);
                    SoapResponse.ParseXRMBaseObject(parentObjects, xrmBaseObject);
                    console.timeEnd("  Parse Child XRM Object " + (i + 1));
                });
                console.timeEnd("  ParseResultInernal");
            }

            static ParseXRMBaseObject(parentObjects: Array<any>, xrmBaseObject: XRMObject): XRMObject {
                var parentObject = parentObjects[parentObjects.length - 1];
                var currentObject = SoapResponse.GetObjectFromBaseXRMObject(parentObjects, xrmBaseObject);
                if (currentObject) {
                    parentObjects.push(currentObject); //Push the current object as the next parent object
                    if (parentObject instanceof Array) {
                        parentObject.push(currentObject);
                    }
                    else {
                        parentObject[xrmBaseObject.Name] = currentObject;
                    }
                    var childObjectsToProcess = new Array<XRMObject>();
                    if (xrmBaseObject.Value instanceof Array) {
                        $.each(xrmBaseObject.Value, function (index, Child: XRMObject) {
                            childObjectsToProcess.push(Child);
                        });
                    }
                    else if (xrmBaseObject.Value && !(typeof xrmBaseObject.Value === "string")) {
                        var Child: XRMObject = xrmBaseObject.Value;
                        childObjectsToProcess.push(Child);
                    }

                    if (childObjectsToProcess.length > 0) {
                        var currentPropertyType = SoapResponse.GetPropertyTypeFromParent(parentObject, xrmBaseObject.Name);
                        var childType: string = null;
                        if (currentPropertyType && currentPropertyType.indexOf("[") == 0) {
                            childType = currentPropertyType.substr(1, currentPropertyType.length - 2);
                        }
                        $.each(childObjectsToProcess, function (i, Child) {
                            if (childType != null) { Child.TypeName = childType; }
                            var childObject = SoapResponse.ParseXRMBaseObject(parentObjects, Child);
                            if (childObject && xrmBaseObject.Name == "Results") {
                                var result = childObject.Value;
                                if (result === undefined) { result = childObject; }
                                parentObject[Child.Name] = result;
                            }
                        });
                    }

                    //Populate the 'FormattedValues' of each of the attributes
                    if (currentObject instanceof Soap.Entity && currentObject.FormattedValues) {
                        for (var attributeName in currentObject.FormattedValues) {
                            var attribute = currentObject.Attributes[attributeName];
                            if (attribute instanceof AttributeValue)
                                attribute.FormattedValue = (<Soap.Entity>currentObject).FormattedValues[attributeName];
                        }
                    }
                    parentObjects.pop(); //Remove the current object as the last parent object
                }
                return currentObject;
            }
            static GetObjectFromBaseXRMObject(parentObjects: Array<any>, xrmBaseObject: XRMObject): any {
                var parentObject = parentObjects[parentObjects.length - 1];
                var currentObject = null;
                var currentPropertyType = xrmBaseObject.TypeName;
                if (currentPropertyType === undefined) {
                    currentPropertyType = SoapResponse.GetPropertyTypeFromParent(parentObject, xrmBaseObject.Name);
                }
                if (currentPropertyType) {
                    if (currentPropertyType == "s") {
                        currentObject = xrmBaseObject.Value;
                    } else if (currentPropertyType == "b") {
                        currentObject = xrmBaseObject.Value == "true" ? true : false;
                    } else if (currentPropertyType == "n") {
                        currentObject = parseFloat(xrmBaseObject.Value);
                    } else if (currentPropertyType == "f") {
                        currentObject = parseFloat(xrmBaseObject.Value);
                    } else if (currentPropertyType == "i") {
                        currentObject = parseInt(xrmBaseObject.Value);
                    } else if (currentPropertyType == "d") {
                        currentObject = new Date(xrmBaseObject.Value);
                    } else if (currentPropertyType == "Array" || currentPropertyType.indexOf("[") == 0) {
                        currentObject = new Array();
                    }
                    else if (currentPropertyType == "AccessRights") {
                        currentObject = new Array<AccessRights>();
                        var rights: Array<string> = xrmBaseObject.Value.split(" ");
                        $.each(rights, function (i, Right) {
                            (<Array<AccessRights>>currentObject).push(AccessRights[Right]);
                        });
                    }
                    else if (currentPropertyType == "AliasedValue") {
                        currentObject = new Soap.AliasedValue();
                    }
                    else if (currentPropertyType == "boolean") {
                        currentObject = new Soap.BooleanValue(xrmBaseObject.Value == "false" ? false : true);
                    }
                    else if (currentPropertyType == "double") {
                        currentObject = new Soap.DoubleValue(parseFloat(xrmBaseObject.Value));
                    }
                    else if (currentPropertyType == "decimal") {
                        currentObject = new Soap.DecimalValue(parseFloat(xrmBaseObject.Value));
                    }
                    else if (currentPropertyType == "dateTime") {
                        if (xrmBaseObject.Value) {
                            xrmBaseObject.Value = xrmBaseObject.Value.replaceAll("T", " ").replaceAll("-", "/");
                        }
                        currentObject = new Soap.DateValue(new Date(xrmBaseObject.Value));
                    }
                    else if (currentPropertyType == "EntityReference") {
                        currentObject = new Soap.EntityReference(
                            SoapResponse.FindArrayElementByName((<Array<XRMObject>>xrmBaseObject.Value), "Name", "Id").Value,
                            SoapResponse.FindArrayElementByName((<Array<XRMObject>>xrmBaseObject.Value), "Name", "LogicalName").Value,
                            SoapResponse.FindArrayElementByName((<Array<XRMObject>>xrmBaseObject.Value), "Name", "Name").Value
                        );
                        xrmBaseObject.Value = null;  //Prevents the parsing of the child values
                    }
                    else if (currentPropertyType == "EntityReferenceCollection") {
                        currentObject = new EntityReferenceCollection();
                    }
                    else if (currentPropertyType == "float") {
                        currentObject = new Soap.FloatValue(parseFloat(xrmBaseObject.Value));
                    }
                    else if (currentPropertyType == "guid") {
                        currentObject = new Soap.GuidValue(xrmBaseObject.Value);
                    }
                    else if (currentPropertyType == "int") {
                        currentObject = new Soap.IntegerValue(parseInt(xrmBaseObject.Value));
                    }
                    else if (currentPropertyType == "Money") {
                        currentObject = new Soap.MoneyValue(parseFloat((<XRMObject>xrmBaseObject.Value).Value));
                        xrmBaseObject.Value = null;  //Prevents the parsing of the child values
                    }
                    else if (currentPropertyType == "OptionSetValue") {
                        currentObject = new Soap.OptionSetValue(parseFloat((<XRMObject>xrmBaseObject.Value).Value));
                        xrmBaseObject.Value = null;  //Prevents the parsing of the child values
                    }
                    else if (currentPropertyType == "OrganizationResponseCollection") {
                        currentObject = new Array<ExecuteMultipleResponseItem>();
                    }
                    else if (currentPropertyType == "string") {
                        currentObject = new Soap.StringValue(xrmBaseObject.Value);
                    }
                    else if (currentPropertyType == "ManagedProperty") {
                        currentObject = new Soap.ManagedProperty(
                            (<Array<XRMObject>>xrmBaseObject.Value)[0].Value == "true",
                            (<Array<XRMObject>>xrmBaseObject.Value)[1].Value == "true",
                            (<Array<XRMObject>>xrmBaseObject.Value)[2].Value == "true"
                        );
                        xrmBaseObject.Value = null;  //Prevents the parsing of the child values
                    }
                    else {
                        currentObject = SoapResponse.CreateObject(currentPropertyType);
                    }
                }
                else if (xrmBaseObject.IsKeyValuePair) {
                    currentObject = xrmBaseObject.Value;
                }
                else {
                    currentObject = SoapResponse.CreateObject(xrmBaseObject.Name);
                }
                return currentObject;
            }
            private static CreateObject(ObjectType: string): any {
                var currentObject = null;
                try {
                    currentObject = new Soap[ObjectType]();
                }
                catch (e1) {
                    try {
                        currentObject = new window[ObjectType]();
                    }
                    catch (e2) {
                        if (!currentObject) {
                            currentObject = {};
                        }
                    }
                }
                return currentObject;
            }
            private static GetPropertyTypeFromParent(parentObject: any, propertyName): string {
                var type: string = null;
                if (parentObject.hasOwnProperty("PropertyTypes")) {
                    var PropertyTypes = <PropertyTypeCollection>parentObject.PropertyTypes;
                    type = PropertyTypes[propertyName];
                }
                return type;
            }
            private static RemoveNameSpace(name: string): string {
                var result = name;
                if (result) {
                    var IndexOfNamespace = result.indexOf(":");
                    if (IndexOfNamespace >= 0) {
                        result = result.substring(IndexOfNamespace + 1);
                    }
                }
                return result;
            }
            private static FindArrayElementByName(array: Array<any>, elementName: string, name: string): any {
                for (var index in array) {
                    var item = array[index];
                    if (item[elementName] == name)
                        return item;
                }
            }
        }
        export class Entity implements ISerializable {
            PropertyTypes = {
                Attributes: "AttributeCollection",
                FormattedValues: "Object",
                LogicalName: "s",
                Id: "s"
            }
            constructor();
            /**
             * Constructor.
             *
             * @param   {string}    logicalName Logical Name of the entity.
             */
            constructor(logicalName: string);

            /**
             * Constructor.
             *
             * @param   {string}    logicalName Logical Name of the entity.
             * @param   {string}    id          The GUID of the existing record.
             */
            constructor(logicalName: string, id: string);
            constructor(public logicalName?: string, public id?: string) {
                this.Attributes = new AttributeCollection();
                if (!id) {
                    this.id = "00000000-0000-0000-0000-000000000000";
                }
            }

            /**
             * The attributes of the Entity.  Use the following notation for setting or retrieving an
             * attribute: Entity.Attributes["attributename"] = ...
             */
            Attributes: AttributeCollection;
            FormattedValues: StringDictionary<string>;
            Serialize(): string {
                var data: string = "<a:Attributes>";
                for (var attributeName in this.Attributes) {
                    var attribute = this.Attributes[attributeName];
                    data += "<a:KeyValuePairOfstringanyType>";
                    data += "<b:key>" + attributeName + "</b:key>";
                    if (attribute === null || (attribute instanceof AttributeValue && attribute.IsNull()))
                        data += "<b:value i:nil=\"true\" />";
                    else if (typeof attribute === "boolean")
                        data += new BooleanValue(attribute).Serialize();
                    else if (typeof attribute === "string")
                        data += new StringValue(attribute).Serialize();
                    else
                        data += attribute.Serialize();

                    data += "</a:KeyValuePairOfstringanyType>";
                }
                data += "</a:Attributes><a:EntityState i:nil=\"true\" /><a:FormattedValues />";
                data += "<a:Id>" + Entity.EncodeValue(this.id) + "</a:Id>";
                data += "<a:LogicalName>" + this.logicalName + "</a:LogicalName>";
                data += "<a:RelatedEntities />";
                return data;
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

            static EncodeDate(date: Date): string {
                return date.getFullYear() + "-" +
                    Entity.padNumber(date.getMonth() + 1) + "-" +
                    Entity.padNumber(date.getDate()) + "T" +
                    Entity.padNumber(date.getHours()) + ":" +
                    Entity.padNumber(date.getMinutes()) + ":" +
                    Entity.padNumber(date.getSeconds());
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
        export class EntityReferenceCollection extends Array<Soap.EntityReference>{ }
        export class AttributeCollection {
            [index: string]: AttributeValue | boolean | string;
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
                var xml = "";
                $.each(this.Items, function (index, entity) {
                    xml += "<a:Entity>" + entity.Serialize() + "</a:Entity>";
                });
                return xml;
            }
        }
        export class ColumnSet implements ISerializable {
            /**
             * Constructor.
             *
             * @param   {Array{string}} columns to be retrieved.
             */
            constructor(columns: Array<string>);

            /**
             * Constructor.
             *
             * @param   {boolean}   allColumns  'true' to return all available columns, otherwise 'false'. Default is 'false'
             */
            constructor(allColumns: boolean);
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

            AddColumn(column: string): void { this.Columns.push(column); }
            AddColumns(columns: Array<string>): void {
                for (var column in columns) {
                    this.AddColumn(column);
                }
            }

            Serialize(): string {
                var data: string = "<a:AllColumns>" + this.AllColumns.toString() + "</a:AllColumns>";
                if (this.Columns.length == 0) {
                    data += "<a:Columns />";
                }
                else {
                    data += "<a:Columns>";
                    $.each(this.Columns, function (i, Column) {
                        data += "<f:string>" + Column + "</f:string>";
                    });
                    data += "</a:Columns>";
                }
                return data;
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
            PropertyTypes = {
                Id: "s",
                LogicalName: "s",
                Name: "s"
            }
            constructor();

            /**
             * Constructor.
             *
             * @param   {string}    id          The GUID of the record.
             * @param   {string}    logicalName Logical Name of the entity.
             */
            constructor(id: string, logicalName: string);

            /**
             * Constructor.
             *
             * @param   {string}    id          The GUID of the record.
             * @param   {string}    logicalName Logical Name of the entity.
             * @param   {string}    name        The 'Display Name' of the record.
             */
            constructor(id: string, logicalName: string, name: string);
            constructor(public id?: string, public entityType?: string, public name?: string) { super(null, AttributeType.EntityReference); }
            Serialize(): string {
                return "<b:value i:type=\"a:EntityReference\"><a:Id>" +
                    Entity.EncodeValue(this.Id) +
                    "</a:Id><a:LogicalName>" +
                    Entity.EncodeValue(this.LogicalName) +
                    "</a:LogicalName><a:Name i:nil=\"true\" /></b:value>";
            }
            GetValueType(): string { return "a:EntityReference"; }
            public get Id(): string {
                return this.id;
            }
            public set Id(value: string) {
                this.id = value;
            }
            public get LogicalName(): string {
                return this.entityType;
            }
            public set LogicalName(value: string) {
                this.entityType = value;
            }
            public get Name(): string {
                return this.name;
            }
            public set Name(value: string) {
                this.name = value;
            }
        }
        export class OptionSetValue extends AttributeValue {
            constructor(public value?: number) { super(value, AttributeType.OptionSetValue); }
            Serialize(): string {
                return "<b:value i:type=\"a:OptionSetValue\"><a:Value>" + this.GetEncodedValue() + "</a:Value></b:value>";
            }
            GetValueType(): string { return "a:OptionSetValue"; }
        }
        export class MoneyValue extends AttributeValue {
            constructor(public value?: number) { super(value, AttributeType.Money); }
            Serialize(): string {
                return "<b:value i:type=\"a:Money\"><a:Value>" + this.GetEncodedValue() + "</a:Value></b:value>";
            }
            GetValueType(): string { return "a:Money"; }
        }
        export class AliasedValue extends AttributeValue {
            constructor(public value?: AttributeValue, public attributeLogicalName?: string, public entityLogicalName?: string) { super(value, AttributeType.AliasedValue); }
            Serialize(): string { throw "Update the 'Serialize' method of the 'AliasedValue'"; }
            GetValueType(): string { throw "Update the 'GetValueType' method of the 'AliasedValue'"; }
        }
        export class BooleanValue extends AttributeValue {
            constructor(public value?: boolean) { super(value, AttributeType.Boolean); }
            GetValueType(): string { return "c:boolean"; }
        }
        export class IntegerValue extends AttributeValue {
            constructor(public value?: number) { super(value, AttributeType.Integer); }
            GetValueType(): string { return "c:int"; }
        }
        export class StringValue extends AttributeValue {
            constructor(public value?: string) { super(value, AttributeType.String); }
            GetValueType(): string { return "c:string"; }
        }
        export class DoubleValue extends AttributeValue {
            constructor(public value?: number) { super(value, AttributeType.Double); }
            GetValueType(): string { return "c:double"; }
        }
        export class DecimalValue extends AttributeValue {
            constructor(public value?: number) { super(value, AttributeType.Decimal); }
            GetValueType(): string { return "c:decimal"; }
        }
        export class FloatValue extends AttributeValue {
            constructor(public value?: number) { super(value, AttributeType.Float); }
            GetValueType(): string { return "c:double"; }
        }
        export class DateValue extends AttributeValue {
            constructor(public value?: Date) { super(value, AttributeType.Date); }
            GetEncodedValue(): string {
                return Entity.EncodeDate(this.value);
            }
            GetValueType(): string { return "c:dateTime"; }
        }
        export class GuidValue extends AttributeValue {
            constructor(public value?: string) {
                super(value, AttributeType.Guid);
                if (!value) {
                    this.value = "00000000-0000-0000-0000-000000000000";
                }
            }
            GetValueType(): string { return "e:guid"; }
        }
        class XRMObject {
            constructor(element: XML.XMLElement, parent: XRMObject) {
                this.Parent = parent;
                if (element.Name && element.Name.indexOf("KeyValuePair") == 0) {
                    //This is a key value pair - so the Name and Value are actually the first and second children
                    this.Name = element.Children[0].Text;
                    this.IsKeyValuePair = true;
                    if (element.Children[1].Text) {
                        this.Value = element.Children[1].Text;
                        this.TypeName = element.Children[1].TypeName;
                    }
                    else if (element.Children[1].Children.length == 1) {
                        if (element.Children[1].Name === "value" && (element.Children[1].Children[0])) { this.Value = new XRMObject(element.Children[1].Children[0], this); }
                        else { this.Value = new XRMObject(element.Children[1], this); }
                        this.TypeName = element.Children[1].TypeName;
                    }
                    else {
                        this.Value = [];
                        var value = this.Value;
                        this.TypeName = element.Children[1].TypeName;
                        $.each(element.Children[1].Children, function (index, ChildNode) {
                            value.push(new XRMObject(ChildNode, this));
                        });
                    }
                }
                else {
                    this.Name = element.Name;
                    this.TypeName = element.TypeName;
                    if (element.Text || element.Children.length == 0) {
                        this.Value = element.Text;
                        this.TypeName = element.TypeName;
                    }
                    else if (element.Children.length == 1) {
                        this.Value = new XRMObject(element.Children[0], this);
                    }
                    else {
                        this.Value = [];
                        var value = this.Value;
                        $.each(element.Children, function (index, ChildNode) {
                            value.push(new XRMObject(ChildNode, this));
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
        export class CreateSoapResponse extends SoapResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["CreateResult"] = "s";
            }
            /** The GUID of the newly created record. */
            CreateResult: string;
        }
        export class UpdateSoapResponse extends SoapResponse { }
        export class DeleteSoapResponse extends SoapResponse { }
        export class RetrieveSoapResponse extends SoapResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["RetrieveResult"] = "Entity";
            }
            /** The retrieved Soap.Entity */
            RetrieveResult: Entity;
        }
        export class RetrieveMultipleSoapResponse extends SoapResponse {
            RetrieveMultipleResult: RetrieveMultipleResult;
        }
        export class RetrieveMultipleResult {
            PropertyTypes = {
                Entities: "[Entity]",
                EntityName: "s",
                MoreRecords: "b",
                PagingCookie: "s",
                TotalRecordCount: "n",
                TotalRecordCountLimitExceeded: "b",
                MinActiveRowVersion: "n"
            }
            Entities: Array<Entity>;
            EntityName: string;
            MoreRecords: boolean = false;
            PagingCookie: string;
            TotalRecordCount: number;
            TotalRecordCountLimitExceeded: boolean;
            MinActiveRowVersion: number;
        }
        export class FaultResponse extends SoapResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["detail"] = "FaultDetail";
                this.PropertyTypes["faultcode"] = "s";
                this.PropertyTypes["faultstring"] = "s";
            }
            detail: FaultDetail;
            faultcode: string;
            faultstring: string;
        }
        export class FaultDetail {
            PropertyTypes = {
                OrganizationServiceFault: "Fault"
            }
            OrganizationServiceFault: Fault;
        }
        export class Fault {
            PropertyTypes = {
                ErrorCode: "s",
                InnerFault: "Fault",
                Message: "s",
                Timestamp: "s"
            }
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
            constructor(public entityName?: string) { }

            Columns: ColumnSet = new ColumnSet(false);
            Criteria: FilterExpression = null;
            PageInfo: PageInfo = new PageInfo();
            Orders: Array<OrderExpression> = [];
            LinkEntities: Array<LinkEntity> = [];
            NoLock: boolean = false;

            serialize(): string {
                //Query
                var data: string = "<query i:type=\"a:QueryExpression\">";

                //Columnset
                data += "<a:ColumnSet>" + this.Columns.Serialize() + "</a:ColumnSet>";

                //Criteria - Serailize the FilterExpression
                if (this.Criteria == null) {
                    data += "<a:Criteria i:nil=\"true\"/>";
                }
                else {
                    data += "<a:Criteria>";
                    data += this.Criteria.serialize();
                    data += "</a:Criteria>";
                }

                data += "<a:Distinct>false</a:Distinct>";
                data += "<a:EntityName>" + this.entityName + "</a:EntityName>";

                //Link Entities
                if (this.LinkEntities.length == 0) {
                    data += "<a:LinkEntities />";
                }
                else {
                    data += "<a:LinkEntities>";
                    $.each(this.LinkEntities, function (index, LinkEntity) {
                        data += LinkEntity.serialize();
                    });
                    data += "</a:LinkEntities>";
                }
                   
                //Sorting
                if (this.Orders.length == 0) {
                    data += "<a:Orders />";
                }
                else {
                    data += "<a:Orders>";
                    $.each(this.Orders, function (index, Order) {
                        data += Order.serialize();
                    });
                    data += "</a:Orders>";
                }

                //Page Info
                data += this.PageInfo.serialize();

                //No Lock
                data += "<a:NoLock>" + this.NoLock.toString() + "</a:NoLock>";

                data += "</query>";
                return data;
            }
        }
        export class FilterExpression {
            constructor();

            /**
             * Constructor.
             *
             * @param   {LogicalOperator}   filterOperator  The filter operator.
             */
            constructor(filterOperator: LogicalOperator);
            constructor(filterOperator?: LogicalOperator) {
                if (filterOperator === undefined) {
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
                var data: string = "<a:Conditions>";
                //Conditions
                $.each(this.Conditions, function (i, Condition) {
                    data += Condition.serialize();
                });
                data += "</a:Conditions>";

                //Filter Operator
                data += "<a:FilterOperator>" + LogicalOperator[this.FilterOperator].toString() + "</a:FilterOperator>";

                //Filters
                if (this.Filters.length == 0) {
                    data += "<a:Filters/>";
                }
                else {
                    data += "<a:Filters>";
                    $.each(this.Filters, function (i, Filter) {
                        data += "<a:FilterExpression>"
                        data += Filter.serialize();
                        data += "</a:FilterExpression>";
                    });
                    data += "</a:Filters>";
                }

                //IsQuickFindFilter
                data += "<a:IsQuickFindFilter>" + this.IsQuickFindFilter.toString() + "</a:IsQuickFindFilter>";
                return data;
            }
        }
        export class ConditionExpression {
            constructor();

            /**
             * Constructor.
             *
             * @param   {string}    attributeName       Name of the attribute.
             * @param   {ConditionOperator} operator    The comparison operator.
             * @param   {AttributeValue}    value       The value to be compared to.
             */
            constructor(attributeName: string, operator: ConditionOperator, value?: AttributeValue);

            /**
             * Constructor.
             *
             * @param   {string}    attributeName       Name of the attribute.
             * @param   {ConditionOperator} operator    The comparison operator.
             * @param   {Array{AttributeValue}} values  The values to be compared to.
             */
            constructor(attributeName: string, operator: ConditionOperator, values?: Array<AttributeValue>);
            constructor(public attributeName?: string, public operator?: ConditionOperator, values?: any) {
                if (values instanceof Array) {
                    this.Values = values;
                }
                else {
                    this.Values.push(values);
                }
            }

            Values: Array<AttributeValue> = [];

            serialize(): string {
                var data: string =
                    "<a:ConditionExpression>" +
                    "<a:AttributeName>" + this.attributeName + "</a:AttributeName>" +
                    "<a:Operator>" + ConditionOperator[this.operator].toString() + "</a:Operator>" +
                    "<a:Values>";
                $.each(this.Values, function (i, value) {
                    data += "<f:anyType i:type=\"" + value.GetValueType() + "\">" + value.GetEncodedValue() + "</f:anyType>";
                });
                data += "</a:Values>" +
                    "</a:ConditionExpression>";
                return data;
            }
        }
        export class PageInfo {
            constructor() { }
            Count: number = 0;
            PageNumber: number = 1;
            PagingCookie: string = null;
            ReturnTotalRecordCount: boolean = false;

            serialize(): string {
                var data: string = "" +
                    "<a:PageInfo>" +
                    "<a:Count>" + this.Count.toString() + "</a:Count>" +
                    "<a:PageNumber>" + this.PageNumber.toString() + "</a:PageNumber>";
                if (this.PagingCookie == null) {
                    data += "<a:PagingCookie i:nil = \"true\" />";
                }
                else {
                    data += "<a:PagingCookie>" + this.PagingCookie + "</a:PagingCookie>";
                }
                data += "<a:ReturnTotalRecordCount>" + this.ReturnTotalRecordCount.toString() + "</a:ReturnTotalRecordCount></a:PageInfo>";
                return data;
            }
        }
        export class OrderExpression {
            constructor();

            /**
             * Constructor.
             *
             * @param   {string}    attributeName   Name of the attribute.
             * @param   {OrderType} orderType       Type of the order.
             */
            constructor(attributeName: string, orderType: OrderType);
            constructor(public attributeName?: string, public orderType?: OrderType) { }
            serialize(): string {
                var data: string = "<a:OrderExpression>";
                data += "<a:AttributeName>" + this.attributeName + "</a:AttributeName>";
                data += "<a:OrderType>" + OrderType[this.orderType].toString() + "</a:OrderType>";
                data += "</a:OrderExpression>";
                return data;
            }
        }
        export class LinkEntity {

            /**
             * Constructor.
             *
             * @param   {string}    linkFromEntityName                  LogicalName of the link from entity.
             * @param   {string}    linkToEntityName                    LogicalName of the link to entity.
             * @param   {string}    linkFromAttributeName               LogicalName of the link from
             *                                                          attribute.
             * @param   {string}    linkToAttributeName                 LogicalName of the link to attribute.
             * @param   {JoinOperator}  optional public joinOperator    JoinOperator    The optional join
             *                                                          operator.
             */
            constructor(
                public linkFromEntityName: string,
                public linkToEntityName: string,
                public linkFromAttributeName: string,
                public linkToAttributeName: string,
                public joinOperator: JoinOperator = Soap.Query.JoinOperator.Inner) { }

            Columns: ColumnSet = new ColumnSet(false);
            EntityAlias: string = null;
            LinkCriteria: FilterExpression = new FilterExpression(Soap.Query.LogicalOperator.And);
            LinkEntities: Array<LinkEntity> = [];

            serialize(): string {
                var data: string = "<a:LinkEntity>";
                data += this.Columns.Serialize();
                if (this.EntityAlias == null) {
                    data += "<a:EntityAlias i:nil=\"true\"/>";
                }
                else {
                    data += "<a:EntityAlias>" + this.EntityAlias + "</a:EntityAlias>";
                }
                data += "<a:JoinOperator>" + JoinOperator[this.joinOperator].toString() + "</a:JoinOperator>";
                data += " <a:LinkCriteria>";
                data += this.LinkCriteria.serialize();
                data += "</a:LinkCriteria>";
                if (this.LinkEntities.length == 0) {
                    data += "<a:LinkEntities />";
                }
                else {
                    data += "<a:LinkEntities>";
                    $.each(this.LinkEntities, function (index, linkEntity) {
                        data += linkEntity.serialize();
                    });
                    data += "</a:LinkEntities>";
                }
                data += "<a:LinkFromAttributeName>" + this.linkFromAttributeName + "</a:LinkFromAttributeName>";
                data += "<a:LinkFromEntityName>" + this.linkFromEntityName + "</a:LinkFromEntityName>";
                data += "<a:LinkToAttributeName>" + this.linkToAttributeName + "</a:LinkToAttributeName>";
                data += "<a:LinkToEntityName>" + this.linkToEntityName + "</a:LinkToEntityName>";
                data += "</a:LinkEntity>";
                return data;
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
        export class PropertyTypeCollection {
            [index: string]: string;
        }
        //These are all the Organization Requests and Responses
        export class ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {string}    requestName             Name of the request.
             * @param   {string}    requestType             Type of the request. Required if Request type
             *                                              differs from pattern: "[RequestName]Type" or if
             *                                              the RequestType namesapce differs from "g" in the
             *                                              global list of namespaces.
             * @param   {ParameterCollection}   parameters  Parameters for the operation.
             */
            constructor(public requestName: string, public requestType?: string, public parameters?: ParameterCollection) {
                if (!parameters || parameters == null) {
                    this.parameters = new ParameterCollection();
                }
                if (!requestType || requestType == null) {
                    this.requestType = "g:" + this.requestName + "Request";
                }
            }
            RequestId: string;
            ExecuteRequestType: string = "request";
            IncludeExecuteHeader: boolean = true;
            Serialize(): string {
                var xml = "";
                if (this.IncludeExecuteHeader) { xml += "<Execute" + Soap.GetNameSpacesXML() + ">"; }
                xml += "<" + this.ExecuteRequestType + " i:type='" + this.requestType + "'><a:Parameters>";
                for (var parameterName in this.parameters) {
                    var parameter = this.parameters[parameterName];

                    xml += "<a:KeyValuePairOfstringanyType>";
                    xml += "<b:key>" + parameterName + "</b:key>";
                    if (parameter instanceof Soap.Entity) {
                        xml += "<b:value i:type=\"a:Entity\">" + parameter.Serialize() + "</b:value>";
                    }
                    else {
                        xml += parameter.Serialize();
                    }
                    xml += "</a:KeyValuePairOfstringanyType>";
                }
                xml += "</a:Parameters>";
                xml += "<a:RequestId i:nil = \"true\" />";
                xml += "<a:RequestName>" + this.requestName + "</a:RequestName>";
                xml += "</" + this.ExecuteRequestType + ">";
                if (this.IncludeExecuteHeader) { xml += "</Execute>"; }
                return xml;
            }
        }
        export class ExecuteResponse extends SoapResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["Results"] = "AttributeCollection";
            }
            Results: AttributeCollection;
            ResponseName: string;
        }
        export class CreateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Entity}    target  Entity to be created.
             */
            constructor(target: Entity) {
                super("Create", "a:CreateRequest");
                this.parameters["Target"] = target;
            }
        }
        export class CreateResponse extends ExecuteResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["id"] = "s";
                this.PropertyTypes["CreateResult"] = "s";
            }
            ParseResult(): void {
                super.ParseResultInernal(this.responseXML);
                if ((<any>this).CreateResult) {
                    this.id = (<any>this).CreateResult;
                }
            }
            id: string;
        }
        export class UpdateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Entity}    target  Entity to be updated.
             */
            constructor(target: Entity) {
                super("Update", "a:UpdateRequest");
                this.parameters["Target"] = target;
            }
        }
        export class UpdateResponse extends ExecuteResponse { }
        export class DeleteRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   target  Entity to be deleted.
             */
            constructor(target: EntityReference) {
                super("Delete", "a:DeleteRequest");
                this.parameters["Target"] = target;
            }
        }
        export class DeleteResponse extends ExecuteResponse { }
        export class RetrieveRequest extends ExecuteRequest {

        }
        export class RetrieveResponse extends ExecuteResponse {

        }

        export class AssociateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  moniker1    The first entity.
             * @param   {Soap.EntityReference}  moniker2    The second entity.
             * @param   {string}    relationshipName        LogicalName of the N:N relationship.
             */
            constructor(moniker1: Soap.EntityReference, moniker2: Soap.EntityReference, relationshipName: string) {
                super("AssociateEntities");
                this.parameters["Moniker1"] = moniker1;
                this.parameters["Moniker2"] = moniker2;
                this.parameters["RelationshipName"] = new Soap.StringValue(relationshipName);
            }
        }
        export class AssociateResponse extends ExecuteResponse { }
        export class DisassociateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  moniker1    The first entity.
             * @param   {Soap.EntityReference}  moniker2    The second entity.
             * @param   {string}    relationshipName        LogicalName of the N:N relationship.
             */
            constructor(moniker1: Soap.EntityReference, moniker2: Soap.EntityReference, relationshipName: string) {
                super("DisassociateEntities");
                this.parameters["Moniker1"] = moniker1;
                this.parameters["Moniker2"] = moniker2;
                this.parameters["RelationshipName"] = new Soap.StringValue(relationshipName);
            }
        }
        export class DisassociateResponse extends ExecuteResponse { }
        export class SetStateRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  entityMoniker   The entity.
             * @param   {number}    state                       The state (statecode).
             * @param   {number}    status                      The status (statuscode).
             */
            constructor(entityMoniker: Soap.EntityReference, state: number, status: number);

            /**
             * Constructor.
             *
             * @param   {Soap.EntityReference}  entityMoniker   The entity.
             * @param   {Soap.OptionSetValue}   state           The state (statecode).
             * @param   {Soap.OptionSetValue}   status          The status (statuscode).
             */
            constructor(entityMoniker: Soap.EntityReference, state: Soap.OptionSetValue, Status: Soap.OptionSetValue);
            constructor(entityMoniker: Soap.EntityReference, state: number | Soap.OptionSetValue, status: number | Soap.OptionSetValue) {
                super("SetState");
                var stateOptionSet: Soap.OptionSetValue;
                if (typeof state === "number") { stateOptionSet = new Soap.OptionSetValue(state); }
                else { stateOptionSet = state; }
                var statusOptionSet: Soap.OptionSetValue;
                if (typeof status === "number") { statusOptionSet = new Soap.OptionSetValue(status); }
                else { statusOptionSet = status; }
                this.parameters = new Soap.ParameterCollection();
                this.parameters["EntityMoniker"] = entityMoniker;
                this.parameters["State"] = stateOptionSet;
                this.parameters["Status"] = statusOptionSet;
            }
        }
        export class SetStateResponse extends ExecuteResponse { }
        export class WhoAmIRequest extends ExecuteRequest {
            constructor() {
                super("WhoAmI");
            }
        }
        export class WhoAmIResponse extends ExecuteResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["UserId"] = "s";
                this.PropertyTypes["BusinessUnitId"] = "s";
                this.PropertyTypes["OrganizationId"] = "s";
            }
            public UserId: string;
            public BusinessUnitId: string;
            public OrganizationId: string;
        }
        export class AssignRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   assignee    The entity (user or team) that the record will be assigned to.
             * @param   {EntityReference}   target      The record to be assigned.
             */
            constructor(assignee: EntityReference, target: EntityReference) {
                super("Assign");
                this.parameters["Assignee"] = assignee;
                this.parameters["Target"] = target;
            }
        }
        export class AssignResponse extends ExecuteResponse { }
        type OrganizationRequestCollection = Array<Soap.ExecuteRequest>;
        export class ExecuteMultipleRequest extends ExecuteRequest {
            /**
             * Constructor.
             *
             * @param   {OrganizationRequestCollection} requests    The requests to be executed.
             * @param   {ExecuteMultipleSettings}   settings        Options for controlling the operation.
             */
            constructor(public Requests?: OrganizationRequestCollection, public Settings?: ExecuteMultipleSettings) {
                super("ExecuteMultiple", "a:ExecuteMultipleRequest");
                if (!Requests || Requests == null) { this.Requests = []; }
                if (!Settings || Settings == null) { this.Settings = new Soap.ExecuteMultipleSettings(); }
            }
            Serialize(): string {
                var xml = "<Execute" + Soap.GetNameSpacesXML() + ">";
                xml += "<request i:type='" + this.requestType + "'><a:Parameters>";
                xml += "<a:KeyValuePairOfstringanyType><b:key>Requests</b:key><b:value i:type=\"l:OrganizationRequestCollection\">";
                $.each(this.Requests, function (i, request) {
                    request.ExecuteRequestType = "l:OrganizationRequest";
                    request.IncludeExecuteHeader = false;
                    xml += request.Serialize();
                });
                xml += "</b:value></a:KeyValuePairOfstringanyType>";
                xml += "<a:KeyValuePairOfstringanyType><b:key>Settings</b:key><b:value i:type=\"l:ExecuteMultipleSettings\">";
                xml += this.Settings.Serialize();
                xml += "</b:value></a:KeyValuePairOfstringanyType>";
                xml += "</a:Parameters>";
                xml += "<a:RequestId i:nil = \"true\" />";
                xml += "<a:RequestName>ExecuteMultiple</a:RequestName>";
                xml += "</request></Execute>";
                return xml;
            }
        }
        export class ExecuteMultipleSettings {

            /**
             * Constructor.
             *
             * @param   {boolean}   optional public continueOnError true to continue on error otherwise false.
             * @param   {boolean}   optional public returnResponses true to return the individual responses otherwise false.
             */
            constructor(public ContinueOnError: boolean = false, public ReturnResponses: boolean = false) { }
            Serialize(): string {
                var xml = "<l:ContinueOnError>" + this.ContinueOnError.toString() + "</l:ContinueOnError>";
                xml += "<l:ReturnResponses>" + this.ReturnResponses.toString() + "</l:ReturnResponses>";
                return xml;
            }
        }
        export class ExecuteMultipleResponse extends ExecuteResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["Responses"] = "ExecuteMultipleResponseItem";
                this.PropertyTypes["ResponseName"] = "s";
            }
            Responses: ExecuteMultipleResponseItemCollection;
        }
        type ExecuteMultipleResponseItemCollection = Array<ExecuteMultipleResponseItem>;

        export class ExecuteMultipleResponseItem {
            PropertyTypes = {
                Fault: "AttributeCollection",
                RequestIndex: "n",
                Response: "SoapResponse"
            }
            Fault: Fault;
            RequestIndex: number;
            Response: SoapResponse;
        }
        export class GrantAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   target      Record to receive access to.
             * @param   {EntityReference}   principal   The user or team to be granted access.
             * @param   {AccessRights}  accessMask      The type of access to be granted.
             */
            constructor(target: EntityReference, principal: EntityReference, accessMask: Soap.AccessRights) {
                super("GrantAccess");
                this.parameters["Target"] = target;
                this.parameters["PrincipalAccess"] = new PrincipalAccess(accessMask, principal);
            }
        }
        export class GrantAccessResponse extends ExecuteResponse { }
        export class ModifyAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   target      Record for which access will be modified.
             * @param   {EntityReference}   principal   The user or team for which the access will be modified.
             * @param   {AccessRights}  accessMask      The type of access to be granted.
             */
            constructor(target: EntityReference, principal: EntityReference, accessMask: AccessRights) {
                super("ModifyAccess");
                this.parameters["Target"] = target;
                this.parameters["PrincipalAccess"] = new PrincipalAccess(accessMask, principal);
            }
        }
        export class ModifyAccessResponse extends ExecuteResponse { }
        export class RevokeAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   target  Record for which access will be removed.
             * @param   {EntityReference}   revokee The user or team to remove access from.
             */
            constructor(target: EntityReference, revokee: EntityReference) {
                super("RevokeAccess");
                this.parameters["Target"] = target;
                this.parameters["Revokee"] = revokee;
            }
        }
        export class RevokeAccessResponse extends ExecuteResponse { }
        export class RetrievePrincipleAccessRequest extends ExecuteRequest {

            /**
             * Constructor.
             *
             * @param   {EntityReference}   target      Record for which the access will be retrieved.
             * @param   {EntityReference}   principal   The user or team for which the access will be checked.
             */
            constructor(target: EntityReference, principal: EntityReference) {
                super("RetrievePrincipalAccess");
                this.parameters["Target"] = target;
                this.parameters["Principal"] = principal;
            }
        }
        export class RetrievePrincipleAccessResponse extends ExecuteResponse {
            constructor(responseXML: string) {
                super(responseXML);
                this.PropertyTypes["AccessRights"] = "[n]";
            }
            AccessRights: Array<AccessRights>;
        }
        export class PrincipalAccess implements Soap.ISerializable {
            constructor(public accessMask: AccessRights, public principal: EntityReference) { }
            Serialize(): string {
                var xml = "<b:value i:type=\"g:PrincipalAccess\"><g:AccessMask>" + AccessRights[this.accessMask].toString() + "</g:AccessMask>";
                xml += "<g:Principal><a:Id>" + Soap.Entity.EncodeValue(this.principal.Id) + "</a:Id>";
                xml += "<a:LogicalName>" + Soap.Entity.EncodeValue(this.principal.LogicalName) + "</a:LogicalName><a:Name i:nil=\"true\" />";
                xml += "</g:Principal></b:value>";
                return xml;
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
        export class RetrieveEntityRequest extends ExecuteRequest {
            constructor(logicalName: string, filters: Array<EntityFilters>, retrieveAsIfPublished: boolean = true) {
                super("RetrieveEntity", "a:RetrieveEntityRequest");
                this.parameters["EntityFilters"] = new EntityFilterCollection(filters);
                this.parameters["MetadataId"] = new GuidValue("00000000-0000-0000-0000-000000000000");
                this.parameters["RetrieveAsIfPublished"] = new BooleanValue(true);
                this.parameters["LogicalName"] = new StringValue(logicalName);
            }
        }
        export class RetrieveEntityResponse extends ExecuteResponse {
            EntityMetadata: EntityMetadata;
        }
        export class EntityFilterCollection implements ISerializable {
            constructor(filters: Array<EntityFilters>) {
                if (filters) {
                    this.Filters = filters;
                }
            }
            Filters: Array<EntityFilters> = [];
            Serialize(): string {
                var filterNames = new Array<string>();
                var containsAllValue = $.inArray(EntityFilters.All, this.Filters) >= 0;
                if (containsAllValue == true) {
                    this.Filters = new Array<EntityFilters>();
                    this.Filters.push(EntityFilters.Attributes);
                    this.Filters.push(EntityFilters.Entity);
                    this.Filters.push(EntityFilters.Privileges);
                    this.Filters.push(EntityFilters.Relationships);
                }
                $.each(this.Filters, function (i, filter) {
                    filterNames.push(EntityFilters[filter].toString());
                });
                var xml = "<b:value i:type=\"h:EntityFilters\">" + filterNames.join(" ") + "</b:value>";
                return xml;
            }
        }
        export class HasPropertyTypes {
            PropertyTypes = new PropertyTypeCollection();
        }
        export class BaseMetaData extends HasPropertyTypes {
            MetadataId: string;
            HasChanged: boolean;
        }
        export class EntityMetadata extends BaseMetaData {
            constructor() {
                super();
                this.PropertyTypes["ActivityTypeMask"] = "n";
                this.PropertyTypes["Attributes"] = "[AttributeMetadata]";
                this.PropertyTypes["AutoCreateAccessTeams"] = "b";
                this.PropertyTypes["AutoRouteToOwnerQueue"] = "b";
                this.PropertyTypes["CanBeInManyToMany"] = "ManagedProperty";
                this.PropertyTypes["CanBePrimaryEntityInRelationship"] = "ManagedProperty";
                this.PropertyTypes["CanBeRelatedEntityInRelationship"] = "aaaaManagedProperty";
                this.PropertyTypes["CanCreateAttributes"] = "ManagedProperty";
                this.PropertyTypes["CanCreateCharts"] = "aaaaManagedProperty";
                this.PropertyTypes["CanCreateForms"] = "ManagedProperty";
                this.PropertyTypes["CanCreateViews"] = "ManagedProperty";
                this.PropertyTypes["CanModifyAdditionalSettings"] = "ManagedProperty";
                this.PropertyTypes["CanTriggerWorkflow"] = "b";
                this.PropertyTypes["Description"] = "Label";
                this.PropertyTypes["DisplayCollectionName"] = "Label";
                this.PropertyTypes["DisplayName"] = "aaaaLabel";
                this.PropertyTypes["EnforceStateTransitions"] = "b";
                this.PropertyTypes["IconLargeName"] = "s";
                this.PropertyTypes["IconMediumName"] = "s";
                this.PropertyTypes["IconSmallName"] = "s";
                this.PropertyTypes["IsAIRUpdated"] = "b";
                this.PropertyTypes["IsActivity"] = "b";
                this.PropertyTypes["IsActivityParty"] = "b";
                this.PropertyTypes["IsAuditEnabled"] = "ManagedProperty";
                this.PropertyTypes["IsAvailableOffline"] = "b";
                this.PropertyTypes["IsBusinessProcessEnabled"] = "b";
                this.PropertyTypes["IsChildEntity"] = "b";
                this.PropertyTypes["IsConnectionsEnabled"] = "ManagedProperty";
                this.PropertyTypes["IsCustomEntity"] = "b";
                this.PropertyTypes["IsCustomizable"] = "ManagedProperty";
                this.PropertyTypes["IsDocumentManagementEnabled"] = "b";
                this.PropertyTypes["IsDuplicateDetectionEnabled"] = "ManagedProperty";
                this.PropertyTypes["IsEnabledForCharts"] = "b";
                this.PropertyTypes["IsEnabledForTrace"] = "aaaa";
                this.PropertyTypes["IsImportable"] = "aaaa";
                this.PropertyTypes["IsIntersect"] = "aaaa";
                this.PropertyTypes["IsMailMergeEnabled"] = "ManagedProperty";
                this.PropertyTypes["IsManaged"] = "b";
                this.PropertyTypes["IsMappable"] = "ManagedProperty";
                this.PropertyTypes["IsQuickCreateEnabled"] = "b";
                this.PropertyTypes["IsReadOnlyInMobileClient"] = "ManagedProperty";
                this.PropertyTypes["IsReadingPaneEnabled"] = "b";
                this.PropertyTypes["IsRenameable"] = "ManagedProperty";
                this.PropertyTypes["IsStateModelAware"] = "b";
                this.PropertyTypes["IsValidForAdvancedFind"] = "b";
                this.PropertyTypes["IsValidForQueue"] = "ManagedProperty";
                this.PropertyTypes["IsVisibleInMobile"] = "ManagedProperty";
                this.PropertyTypes["IsVisibleInMobileClient"] = "ManagedProperty";
                this.PropertyTypes["LogicalName"] = "s";
                this.PropertyTypes["ManyToManyRelationships"] = "[ManyToManyRelationshipMetadata]";
                this.PropertyTypes["ManyToOneRelationships"] = "[OneToManyRelationshipMetadata]";
                this.PropertyTypes["OwnershipType"] = "n";
                this.PropertyTypes["PrimaryIdAttribute"] = "s";
                this.PropertyTypes["PrimaryNameAttribute"] = "s";
                this.PropertyTypes["Privileges"] = "[n]";
                this.PropertyTypes["RecurrenceBaseEntityLogicalName"] = "s";
                this.PropertyTypes["ReportViewName"] = "s";
                this.PropertyTypes["SchemaName"] = "s";
                this.PropertyTypes["IntroducedVersion"] = "s";
                this.PropertyTypes["PrimaryImageAttribute"] = "s";
                this.PropertyTypes["CanChangeHierarchicalRelationship"] = "ManagedProperty";
                this.PropertyTypes["EntityHelpUrl"] = "s";
                this.PropertyTypes["EntityHelpUrlEnabled"] = "b";
            }
            ActivityTypeMask: number = null;
            Attributes: Array<AttributeMetadata> = [];
            AutoCreateAccessTeams: boolean;
            AutoRouteToOwnerQueue: boolean;
            CanBeInManyToMany: ManagedProperty;
            CanBePrimaryEntityInRelationship: ManagedProperty;
            CanBeRelatedEntityInRelationship: ManagedProperty;
            CanCreateAttributes: ManagedProperty;
            CanCreateCharts: ManagedProperty;
            CanCreateForms: ManagedProperty;
            CanCreateViews: ManagedProperty;
            CanModifyAdditionalSettings: ManagedProperty;
            CanTriggerWorkflow: boolean;
            Description: Label;
            DisplayCollectionName: Label;
            DisplayName: Label;
            EnforceStateTransitions: boolean;
            IconLargeName: string;
            IconMediumName: string;
            IconSmallName: string;
            IsAIRUpdated: boolean;
            IsActivity: boolean;
            IsActivityParty: boolean;
            IsAuditEnabled: ManagedProperty;
            IsAvailableOffline: boolean;
            IsBusinessProcessEnabled: boolean;
            IsChildEntity: boolean;
            IsConnectionsEnabled: ManagedProperty;
            IsCustomEntity: boolean;
            IsCustomizable: ManagedProperty;
            IsDocumentManagementEnabled: boolean;
            IsDuplicateDetectionEnabled: ManagedProperty;
            IsEnabledForCharts: boolean;
            IsEnabledForTrace: boolean;
            IsImportable: boolean;
            IsIntersect: boolean;
            IsMailMergeEnabled: ManagedProperty;
            IsManaged: boolean;
            IsMappable: ManagedProperty;
            IsQuickCreateEnabled: boolean;
            IsReadOnlyInMobileClient: ManagedProperty;
            IsReadingPaneEnabled: boolean;
            IsRenameable: ManagedProperty;
            IsStateModelAware: boolean;
            IsValidForAdvancedFind: boolean;
            IsValidForQueue: ManagedProperty;
            IsVisibleInMobile: ManagedProperty;
            IsVisibleInMobileClient: ManagedProperty;
            LogicalName: string;
            ManyToManyRelationships: Array<ManyToManyRelationshipMetadata> = [];
            ManyToOneRelationships: Array<OneToManyRelationshipMetadata> = [];
            ObjectTypeCode: number;
            OneToManyRelationships: Array<OneToManyRelationshipMetadata> = [];
            OwnershipType: OwnershipTypes;
            PrimaryIdAttribute: string;
            PrimaryNameAttribute: string;
            Privileges: Array<SecurityPrivilegeMetadata>;
            RecurrenceBaseEntityLogicalName: string;
            ReportViewName: string;
            SchemaName: string;
            IntroducedVersion: string;
            PrimaryImageAttribute: string;
            CanChangeHierarchicalRelationship: ManagedProperty;
            EntityHelpUrl: string;
            EntityHelpUrlEnabled: boolean;
        }
        export class AttributeMetadata extends BaseMetaData {
            constructor() {
                super();
                this.PropertyTypes["AttributeOf"] = "s";
                this.PropertyTypes["AttributeType"] = "n";
                this.PropertyTypes["CanBeSecuredForCreate"] = "b";
                this.PropertyTypes["CanBeSecuredForRead"] = "b";
                this.PropertyTypes["CanBeSecuredForUpdate"] = "b";
                this.PropertyTypes["CanModifyAdditionalSettings"] = "ManagedProperty";
                this.PropertyTypes["ColumnNumber"] = "n";
                this.PropertyTypes["DeprecatedVersion"] = "s";
                this.PropertyTypes["Description"] = "Label";
                this.PropertyTypes["DisplayName"] = "Label";
                this.PropertyTypes["EntityLogicalName"] = "s";
                this.PropertyTypes["IsAuditEnabled"] = "ManagedProperty";
                this.PropertyTypes["IsCustomAttribute"] = "b";
                this.PropertyTypes["IsCustomizable"] = "ManagedProperty";
                this.PropertyTypes["IsManaged"] = "b";
                this.PropertyTypes["IsPrimaryId"] = "b";
                this.PropertyTypes["IsPrimaryName"] = "b";
                this.PropertyTypes["IsRenameable"] = "ManagedProperty";
                this.PropertyTypes["IsSecured"] = "b";
                this.PropertyTypes["IsValidForAdvancedFind"] = "ManagedProperty";
                this.PropertyTypes["IsValidForCreate"] = "b";
                this.PropertyTypes["IsValidForRead"] = "b";
                this.PropertyTypes["IsValidForUpdate"] = "b";
                this.PropertyTypes["LinkedAttributeId"] = "s";
                this.PropertyTypes["LogicalName"] = "s";
                this.PropertyTypes["RequiredLevel"] = "ManagedProperty";
                this.PropertyTypes["SchemaName"] = "s";
                this.PropertyTypes["AttributeTypeName"] = "s";
                this.PropertyTypes["IntroducedVersion"] = "s";
                this.PropertyTypes["IsLogical"] = "b";
                this.PropertyTypes["SourceType"] = "n";
            }
            AttributeOf: string;
            AttributeType: AttributeTypeCode;
            CanBeSecuredForCreate: boolean;
            CanBeSecuredForRead: boolean;
            CanBeSecuredForUpdate: boolean;
            CanModifyAdditionalSettings: ManagedProperty;
            ColumnNumber: number;
            DeprecatedVersion: string;
            Description: Label;
            DisplayName: Label;
            EntityLogicalName: string;
            IsAuditEnabled: ManagedProperty;
            IsCustomAttribute: boolean;
            IsCustomizable: ManagedProperty;
            IsManaged: boolean;
            IsPrimaryId: boolean;
            IsPrimaryName: boolean;
            IsRenameable: ManagedProperty;
            IsSecured: boolean;
            IsValidForAdvancedFind: ManagedProperty;
            IsValidForCreate: boolean;
            IsValidForRead: boolean;
            IsValidForUpdate: boolean;
            LinkedAttributeId: string;
            LogicalName: string;
            RequiredLevel: ManagedProperty;
            SchemaName: string;
            AttributeTypeName: string;
            IntroducedVersion: string;
            IsLogical: boolean;
            SourceType: number;
        }
        export class BooleanAttributeMetadata extends AttributeMetadata {
            DefaultValue: boolean;
            OptionSet: OptionSetMetadata;
            FormulaDefinition: string;
            SourceTypeMask: number;
        }
        export class DateTimeAttributeMetadata extends AttributeMetadata {
            Format: DateTimeFormat;
            ImeMode: ImeMode;
            FormulaDefinition: string;
            SourceTypeMask: number;
        }
        export class DecimalAttributeMetadata extends AttributeMetadata {
            ImeMode: ImeMode;
            MaxValue: number;
            MinValue: number;
            Precision: number;
            FormulaDefinition: string;
            SourceTypeMask: number;
        }
        export class DoubleAttributeMetadata extends AttributeMetadata {
            ImeMode: ImeMode;
            MaxValue: number;
            MinValue: number;
            Precision: number;
        }
        export class ImageAttributeMetadata extends AttributeMetadata {
            IsPrimaryImage: boolean;
            MaxHeight: number;
            MaxWidth: number;
        }
        export class IntegerAttributeMetadata extends AttributeMetadata {
            Format: IntegerFormat;
            MaxValue: number;
            MinValue: number;
            FormulaDefinition: string;
            SourceTypeMask: number;
        }
        export class LookupAttributeMetadata extends AttributeMetadata {
            Targets: Array<string> = [];
        }
        export class MemoAttributeMetadata extends AttributeMetadata {
            Format: StringFormat;
            ImeMode: ImeMode;
            MaxLength: number;
            IsLocalizable: boolean;
        }
        export class MoneyAttributeMetadata extends AttributeMetadata {
            CalculationOf: string;
            ImeMode: ImeMode;
            MaxValue: number;
            MinValue: number;
            Precision: number;
            PrecisionSource: number;
            FormulaDefinition: string;
            IsBaseCurrency: boolean;
            SourceTypeMask: number;
        }
        export class PicklistAttributeMetadata extends AttributeMetadata {
            DefaultFormValue: number;
            OptionSet: OptionSetMetadata;
        }
        export class StringAttributeMetadata extends AttributeMetadata {
            Format: StringFormat;
            ImeMode: ImeMode;
            MaxLength: number;
            YomiOf: string;
            FormatName: StringFormatName;
            FormulaDefinition: string;
            IsLocalizable: boolean;
            SourceTypeMask: number;
        }
        export class RelationshipMetadataBase extends BaseMetaData {
            IsCustomRelationship: boolean;
            IsCustomizable: ManagedProperty;
            IsManaged: boolean;
            IsValidForAdvancedFind: boolean;
            SchemaName: string;
            SecurityTypes: SecurityTypes;
            IntroducedVersion: string;
            RelationshipType: RelationshipType;
        }
        export class ManyToManyRelationshipMetadata extends RelationshipMetadataBase {
            Entity1AssociatedMenuConfiguration: AssociatedMenuConfiguration;
            Entity1IntersectAttribute: string;
            Entity1LogicalName: string;
            Entity2AssociatedMenuConfiguration: AssociatedMenuConfiguration;
            Entity2IntersectAttribute: string;
            Entity2LogicalName: string;
            IntersectEntityName: string;
        }
        export class OneToManyRelationshipMetadata extends RelationshipMetadataBase {
            constructor() {
                super();
                this.PropertyTypes["AssociatedMenuConfiguration"] = "AssociatedMenuConfiguration";
                this.PropertyTypes["CascadeConfiguration"] = "CascadeConfiguration";
            }
            AssociatedMenuConfiguration: AssociatedMenuConfiguration;
            CascadeConfiguration: CascadeConfiguration;
            IsHierarchical: boolean;
            ReferencedAttribute: string;
            ReferencedEntity: string;
            ReferencingAttribute: string;
            ReferencingEntity: string;
        }
        export class OptionSetMetadata extends BaseMetaData {
            Description: Label;
            DisplayName: Label;
            IsCustomOptionSet: boolean;
            IsCustomizable: ManagedProperty;
            IsGlobal: boolean;
            IsManaged: boolean;
            Name: string;
            OptionSetType: OptionSetType;
            IntroducedVersion: string;
            Options: Array<OptionMetadata> = [];
            FormulaDefinition: string;
            SourceTypeMask: number;
        }
        export class OptionMetadata extends BaseMetaData {
            Description: Label;
            IsManaged: boolean;
            Label: Label;
            Value: number;
        }
        export class AssociatedMenuConfiguration {
            Behavior: AssociatedMenuBehavior;
            Group: AssociatedMenuGroup;
            Label: Label;
            Order: number;
        }
        export class CascadeConfiguration extends HasPropertyTypes {
            constructor() {
                super();
                this.PropertyTypes["Assign"] = "n";
                this.PropertyTypes["Delete"] = "n";
                this.PropertyTypes["Merge"] = "n";
                this.PropertyTypes["Reparent"] = "n";
                this.PropertyTypes["Share"] = "n";
                this.PropertyTypes["Unshare"] = "n";
            }
            PropertyTypes: PropertyTypeCollection;
            Assign: CascadeType;
            Delete: CascadeType;
            Merge: CascadeType;
            Reparent: CascadeType;
            Share: CascadeType;
            Unshare: CascadeType;
        }
        export class SecurityPrivilegeMetadata {
            CanBeBasic: boolean;
            CanBeDeep: boolean;
            CanBeGlobal: boolean;
            CanBeLocal: boolean;
            Name: string;
            PrivilegeId: string;
            PrivilegeType: PrivilegeType;
        }
        export class Label extends HasPropertyTypes {
            constructor() {
                super();
                this.PropertyTypes["LocalizedLabels"] = "Array";
                this.PropertyTypes["UserLocalizedLabel"] = "LocalizedLabel";
            }
            LocalizedLabels: Array<LocalizedLabel>;
            UserLocalizedLabel: LocalizedLabel;
        }
        export class StringFormatName {
            Value: string;
        }
        export class LocalizedLabel {
            MetadataId: string;
            HasChanged: boolean;
            IsManaged: boolean;
            Label: string;
            LanguageCode: number;
        }
        export class ManagedProperty {
            constructor(public CanBeChanged: boolean, public ManagedPropertyLogicalName: boolean, public Value: boolean) { }
        }

        export enum AssociatedMenuBehavior {
            // Summary:
            //     Use the collection name for the associated menu. Value = 0.
          
            UseCollectionName = 0,
            //
            // Summary:
            //     Use the label for the associated menu. Value = 1.
          
            UseLabel = 1,
            //
            // Summary:
            //     Do not show the associated menu. Value = 2.
            
            DoNotDisplay = 2
        }
        export enum AssociatedMenuGroup {
            // Summary:
            //     Show the associated menu in the details group. Value = 0.
            
            Details = 0,
            //
            // Summary:
            //     Show the associated menu in the sales group. Value = 1.
           
            Sales = 1,
            //
            // Summary:
            //     Show the associated menu in the service group. Value = 2.
            
            Service = 2,
            //
            // Summary:
            //     Show the associated menu in the marketing group. Value = 3.
           
            Marketing = 3
        }
        export enum AttributeTypeCode {
            // Summary:
            //     A Boolean attribute. Value = 0.
            
            Boolean = 0,
            //
            // Summary:
            //     An attribute that represents a customer. Value = 1.
           
            Customer = 1,
            //
            // Summary:
            //     A date/time attribute. Value = 2.
           
            DateTime = 2,
            //
            // Summary:
            //     A decimal attribute. Value = 3.
           
            Decimal = 3,
            //
            // Summary:
            //     A double attribute. Value = 4.
           
            Double = 4,
            //
            // Summary:
            //     An integer attribute. Value = 5.
          
            Integer = 5,
            //
            // Summary:
            //     A lookup attribute. Value = 6.
           
            Lookup = 6,
            //
            // Summary:
            //     A memo attribute. Value = 7.
           
            Memo = 7,
            //
            // Summary:
            //     A money attribute. Value = 8.
            
            Money = 8,
            //
            // Summary:
            //     An owner attribute. Value = 9.
           
            Owner = 9,
            //
            // Summary:
            //     A partylist attribute. Value = 10.
            
            PartyList = 10,
            //
            // Summary:
            //     A picklist attribute. Value = 11.
            
            Picklist = 11,
            //
            // Summary:
            //     A state attribute. Value = 12.
           
            State = 12,
            //
            // Summary:
            //     A status attribute. Value = 13.
           
            Status = 13,
            //
            // Summary:
            //     A string attribute. Value = 14.
           
            String = 14,
            //
            // Summary:
            //     An attribute that is an ID. Value = 15.
           
            Uniqueidentifier = 15,
            //
            // Summary:
            //     An attribute that contains calendar rules. Value = 0x10.
            
            CalendarRules = 16,
            //
            // Summary:
            //     An attribute that is created by the system at run time. Value = 0x11.
           
            Virtual = 17,
            //
            // Summary:
            //     A big integer attribute. Value = 0x12.
           
            BigInt = 18,
            //
            // Summary:
            //     A managed property attribute. Value = 0x13.
           
            ManagedProperty = 19,
            //
            // Summary:
            //     An entity name attribute. Value = 20.
           
            EntityName = 20
        }
        export enum CascadeType {
            NoCascade,
            RemoveLink
        }
        export enum DateTimeFormat {
            // Summary:
            //     Display the date only. Value = 0.
           
            DateOnly = 0,
            //
            // Summary:
            //     Display the date and time. Value = 1.
            DateAndTime = 1,
        }
        export enum EntityFilters {
            // Summary:
            //     Use this to retrieve only entity information. Equivalent to EntityFilters.Entity.
            //     Value = 1.
            Default = 1,
            //
            // Summary:
            //     Use this to retrieve only entity information. Equivalent to EntityFilters.Default.
            //     Value = 1.
       
            Entity = 1,
            //
            // Summary:
            //     Use this to retrieve entity information plus attributes for the entity. Value
            //     = 2.
       
            Attributes = 2,
            //
            // Summary:
            //     Use this to retrieve entity information plus privileges for the entity. Value
            //     = 4.
        
            Privileges = 4,
            //
            // Summary:
            //     Use this to retrieve entity information plus entity relationships for the
            //     entity. Value = 8.
       
            Relationships = 8,
            //
            // Summary:
            //     Use this to retrieve all data for an entity. Value = 15.
            All = 15
        }
        export enum ImeMode {
            // Summary:
            //     Specifies that the IME mode is chosen automatically. Value =0.
           
            Auto = 0,
            //
            // Summary:
            //     Specifies that the IME mode is inactive. Value = 1.
            
            Inactive = 1,
            //
            // Summary:
            //     Specifies that the IME mode is active. Value = 2.
            
            Active = 2,
            //
            // Summary:
            //     Specifies that the IME mode is disabled. Value = 3.
            
            Disabled = 3
        }
        export enum IntegerFormat {
            // Summary:
            //     Specifies to display an edit field for an integer. Value = 0.
        
            None = 0,
            //
            // Summary:
            //     Specifies to display the integer as a drop down list of durations. Value
            //     = 1.
          
            Duration = 1,
            //
            // Summary:
            //     Specifies to display the integer as a drop down list of time zones. Value
            //     = 2.
            
            TimeZone = 2,
            //
            // Summary:
            //     Specifies the display the integer as a drop down list of installed languages.
            //     Value = 3.
           
            Language = 3,
            //
            // Summary:
            //     Specifies a locale. Value = 4.
           
            Locale = 4,
        }
        export enum OwnershipTypes {
            // Summary:
            //     The entity does not have an owner. internal Value = 0.
            None = 0,
            //
            // Summary:
            //     The entity is owned by a system user. Value = 1.
           
            UserOwned = 1,
            //
            // Summary:
            //     The entity is owned by a team. internalValue = 2.
          
            TeamOwned = 2,
            //
            // Summary:
            //     The entity is owned by a business unit. internal Value = 4.
           
            BusinessOwned = 4,
            //
            // Summary:
            //     The entity is owned by an organization. Value = 8.
           
            OrganizationOwned = 8,
            //
            // Summary:
            //     The entity is parented by a business unit. internal Value = 16.
           
            BusinessParented = 16
        }
        export enum OptionSetType {
            // Summary:
            //     The option set provides a list of options. Value = 0.
           
            Picklist = 0,
            //
            // Summary:
            //     The option set represents state options for a Microsoft.Xrm.Sdk.Metadata.StateAttributeMetadata
            //     attribute. Value = 1.
           
            State = 1,
            //
            // Summary:
            //     The option set represents status options for a Microsoft.Xrm.Sdk.Metadata.StatusAttributeMetadata
            //     attribute. Value = 2.
            
            Status = 2,
            //
            // Summary:
            //     The option set provides two options for a Microsoft.Xrm.Sdk.Metadata.BooleanAttributeMetadata
            //     attribute. Value = 3.
           
            Boolean = 3,
        }
        export enum PrivilegeType {
            // Summary:
            //     Specifies no privilege. Value = 0.
        
            None = 0,
            //
            // Summary:
            //     The create privilege. Value = 1.
           
            Create = 1,
            //
            // Summary:
            //     The read privilege. Value = 2.
          
            Read = 2,
            //
            // Summary:
            //     The write privilege. Value = 3.
           
            Write = 3,
            //
            // Summary:
            //     The delete privilege. Value = 4.
           
            Delete = 4,
            //
            // Summary:
            //     The assign privilege. Value = 5.
           
            Assign = 5,
            //
            // Summary:
            //     The share privilege. Value = 6.
           
            Share = 6,
            //
            // Summary:
            //     The append privilege. Value = 7.
           
            Append = 7,
            //
            // Summary:
            //     The append to privilege. Value = 8.
           
            AppendTo = 8
        }
        export enum RelationshipType {
            // Summary:
            //     The entity relationship is a One-to-Many relationship. Value = 0.
            OneToManyRelationship = 0,
            //
            // Summary:
            //     The default value. Equivalent to OneToManyRelationship. Value = 0.
            Default = 0,
            //
            // Summary:
            //     The entity relationship is a Many-to-Many relationship. Value = 1.
            ManyToManyRelationship = 1
        }
        export enum SecurityTypes {
            // Summary:
            //     No security privileges are checked during create or update operations. Value
            //     = 0.
           
            None = 0,
            //
            // Summary:
            //     The Microsoft.Xrm.Sdk.Metadata.PrivilegeType.Append and Microsoft.Xrm.Sdk.Metadata.PrivilegeType.AppendTo
            //     privileges are checked for create or update operations. Value = 1.
           
            Append = 1,
            //
            // Summary:
            //     Security for the referencing entity record is derived from the referenced
            //     entity record. Value = 2.
           
            ParentChild = 2,
            //
            // Summary:
            //     Security for the referencing entity record is derived from a pointer record.
            //     Value = 4.
           
            Pointer = 4,
            //
            // Summary:
            //     The referencing entity record inherits security from the referenced security
            //     record. Value = 8.
           
            Inheritance = 8
        }
        export enum StringFormat {
            // Summary:
            //     Specifies to display the string as an e-mail. Value = 0.
           
            Email = 0,
            //
            // Summary:
            //     Specifies to display the string as text. Value = 1.
          
            Text = 1,
            //
            // Summary:
            //     Specifies to display the string as a text area. Value = 2.
           
            TextArea = 2,
            //
            // Summary:
            //     Specifies to display the string as a URL. Value = 3.
           
            Url = 3,
            //
            // Summary:
            //     Specifies to display the string as a ticker symbol. Value = 4.
           
            TickerSymbol = 4,
            //
            // Summary:
            //     Specifies to display the string as a phonetic guide. Value = 5.
            
            PhoneticGuide = 5,
            //
            // Summary:
            //     Specifies to display the string as a version number. Value = 6.
           
            VersionNumber = 6,
            //
            // Summary:
            //     internal
            
            Phone = 7
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