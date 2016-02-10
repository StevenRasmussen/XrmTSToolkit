namespace UMC.FormScripts {
    export class AccountFormScripts {
        constructor() { }

        public OnFormLoad(): void {
            (<Xrm.Page.LookupAttribute>Xrm.Page.getAttribute("new_parentaccountcopied")).setValue(null);

            var query = new XrmTSToolkit.Soap.Query.QueryExpression("account");
            query.Columns = new XrmTSToolkit.Soap.ColumnSet(["parentaccountid"]);
            var retrieveMultiplePromise = XrmTSToolkit.Soap.Retrieve(Xrm.Page.data.entity.getId(), Xrm.Page.data.entity.getEntityName(), new XrmTSToolkit.Soap.ColumnSet(["parentaccountid"]));
            retrieveMultiplePromise.done(function (response: XrmTSToolkit.Soap.RetrieveSoapResponse) {
                debugger;
                var parentAccountRef = response.RetrieveResult.Attributes["parentaccountid"] as XrmTSToolkit.Soap.EntityReference;
                if (parentAccountRef) {
                    (<Xrm.Page.LookupAttribute>Xrm.Page.getAttribute("new_parentaccountcopied")).setValue([parentAccountRef]);
                    alert("XrmTSToolkit Test: Successfully set the 'Parent Account Copied' field");
                }
                else
                    alert("XrmTSToolkit Test: Please set the parent account in order to test the copy of the 'Parent Account' to the 'Parent Account Copied' field");
            }).fail(function () {
                alert("XrmTSToolkit Test: Error setting the 'Parent Account Copied' field");
            });
        }
    }
}

var accountFormScripts = new UMC.FormScripts.AccountFormScripts();
