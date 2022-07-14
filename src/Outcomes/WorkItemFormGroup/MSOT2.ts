import {
  IWorkItemFormService,
  WorkItemTrackingServiceIds,
} from "azure-devops-extension-api/WorkItemTracking";
import { IListBoxItem} from "azure-devops-ui/ListBox";
import * as SDK from "azure-devops-extension-sdk";

let results = GetAllMyDatasources()

async function GetAllMyDatasources() {
  let response = await ParseT1Data()
  return response
}
async function RetrieveT1Data(){
      const workItemFormService = await SDK.getService<IWorkItemFormService>(
        WorkItemTrackingServiceIds.WorkItemFormService
      )
        let outcomeT1Data = (await workItemFormService.getFieldValue("Custom.MSOutcomeT2Data")).toString();
        return outcomeT1Data
    }

async function ParseT1Data(){
  let dataplaceholder = new Array<IListBoxItem<{}>>();
  const data = await RetrieveT1Data()
  return data
}
export const MSOutcomeT2Data = results


