import {
  IWorkItemFormService,
  WorkItemTrackingServiceIds,
} from "azure-devops-extension-api/WorkItemTracking";
import "./WorkItemFormGroup.scss";
import * as SDK from "azure-devops-extension-sdk";
import { Dropdown } from "azure-devops-ui/Dropdown";
import * as React from "react";
import { showRootComponent } from "../../Common";
import { IListBoxItem, ListBoxItemType } from "azure-devops-ui/ListBox";
import { Link } from "azure-devops-ui/Link";
import { Observer } from "azure-devops-ui/Observer";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { DropdownSelection } from "azure-devops-ui/Utilities/DropdownSelection";
import { MSOutcomeT1Data } from "./MSOT1"
import { MSOutcomeT2Data } from "./MSOT2";
import { MSOutcomeT3Data } from "./MSOT3";
import { MSOP1Data } from "./MSOSP1";
import { MSOP2Data } from "./MSOSP2";
import { Card } from "azure-devops-ui/Card";

interface WorkItemFormGroupComponentState {
  MSOutcome1: string;
  MSOutcome1Data: Array<IListBoxItem<{}>>;
  MSOutcome1DataTestString: string;
  MSOutcome2: string;
  MSOutcome2Data: Array<IListBoxItem<{}>>;
  MSOutcome3: string;
  MSOutcome3Data: Array<IListBoxItem<{}>>;
  MSOSP1: string;
  MSOSP2: string;
  MSOutcomeStatement: string;
  MSOutcomeStatementP1Data: Array<IListBoxItem<{}>>;
  MSOutcomeStatementP2Data: Array<IListBoxItem<{}>>;
  MSOutcomeRetrievedCount: number;
  IsRenderReady: boolean;
  MSOutcomeT3DDIsDisabled: boolean;
  outcomeDDItemsT3: Array<IListBoxItem<{}>>;
  selectionDDT3: DropdownSelection;
  GenerateOutcomeBTNIsDisabled: boolean;
  ShowOutcomeCard: boolean;
}

export class WorkItemFormGroupComponent extends React.Component<{},  WorkItemFormGroupComponentState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      // eventContent: "",
      MSOutcome1: "",
      MSOutcome1Data: [],
      MSOutcome1DataTestString: "",
      MSOutcome2: "",
      MSOutcome2Data: [],
      MSOutcome3: "",
      MSOutcome3Data: [],
      MSOSP1: "",
      MSOSP2: "",
      MSOutcomeStatement: "",
      MSOutcomeStatementP1Data: [],
      MSOutcomeStatementP2Data: [],
      MSOutcomeRetrievedCount: 0,
      IsRenderReady: false,
      MSOutcomeT3DDIsDisabled: true,
      outcomeDDItemsT3: [],
      selectionDDT3: new DropdownSelection(),
      GenerateOutcomeBTNIsDisabled: true,
      ShowOutcomeCard: false,
  }
}

  private selectedItemIDT1 = new ObservableValue<string>("");
  private selectedItemTextT1 = new ObservableValue<string>("");
  private selectedItemIDT2 = new ObservableValue<string>("");
  private selectedItemTextT2 = new ObservableValue<string>("");
  private selectedItemIDT3 = new ObservableValue<string>("");
  private selectedItemTextT3 = new ObservableValue<string>("");


  public componentDidMount() {
    SDK.init().then(() => {
    this.fetchAllJSONData().then(() => {
    this.getAllMSOutcomeFieldValue();
    })
    });
  }

  public increment() {
    this.setState(MSOutcomeRetrievedCount=>({
      MSOutcomeRetrievedCount:MSOutcomeRetrievedCount.MSOutcomeRetrievedCount +1
      }));
    if (this.state.MSOutcomeRetrievedCount == 3) {
      this.setState({
        IsRenderReady: true
     });
    }
  }

  public async fetchAllJSONData(){
    let dataplaceholderT1 = new Array<IListBoxItem<{}>>();
    let dataplaceholderT2 = new Array<IListBoxItem<{}>>();
    let dataplaceholderT3 = new Array<IListBoxItem<{}>>();
    let dataplaceholderOSP1 = new Array<IListBoxItem<{}>>();
    let dataplaceholderOSP2 = new Array<IListBoxItem<{}>>();
    const T1 = (await MSOutcomeT1Data);
    const T2 = (await MSOutcomeT2Data)
    const T3 = (await MSOutcomeT3Data)
    const OSP1 = (await MSOP1Data)
    const OSP2 = (await MSOP2Data)
  //   this.setState({
  //     MSOutcome1DataTestString: T1
  //  });
    var parseDataT1 = JSON.parse(T1)
    for (let entry of parseDataT1.items) {
      dataplaceholderT1.push({ "id": entry.id, "text": entry.text})
    }
    var parseDataT2 = JSON.parse(T2)
    for (let entry of parseDataT2.items) {
      dataplaceholderT2.push({ "id": entry.id, "text": entry.text})
    }
    var parseDataT3 = JSON.parse(T3)
    for (let entry of parseDataT3.items) {
      dataplaceholderT3.push({ "id": entry.id, "text": entry.text, "groupId": entry.groupId})
    }
    var parseDataOSP1 = JSON.parse(OSP1)
    for (let entry of parseDataOSP1.items) {
      dataplaceholderOSP1.push({ "id": entry.id, "text": entry.text, "groupId": entry.groupId})
    }
    var parseDataOSP2 = JSON.parse(OSP2)
    for (let entry of parseDataOSP2.items) {
      dataplaceholderOSP2.push({ "id": entry.id, "text": entry.text, "groupId": entry.groupId})
    }
    const MSOT1DataSorted = dataplaceholderT1.sort(function(a,b) {
        var nameA = a.text?.toUpperCase(); // ignore upper and lowercase
        var nameB = b.text?.toUpperCase(); // ignore upper and lowercase
        if (nameA! < nameB!) {
          return -1; //nameA comes first
        }
        if (nameA! > nameB!) {
          return 1; // nameB comes first
        }
        return 0;  // names must be equal
      });
      const MSOT2DataSorted = dataplaceholderT2.sort(function(a,b) {
        var nameA = a.text?.toUpperCase(); // ignore upper and lowercase
        var nameB = b.text?.toUpperCase(); // ignore upper and lowercase
        if (nameA! < nameB!) {
          return -1; //nameA comes first
        }
        if (nameA! > nameB!) {
          return 1; // nameB comes first
        }
        return 0;  // names must be equal
      });
      const MSOT3DataSorted = dataplaceholderT3.sort(function(a,b) {
        var nameA = a.text?.toUpperCase(); // ignore upper and lowercase
        var nameB = b.text?.toUpperCase(); // ignore upper and lowercase
        if (nameA! < nameB!) {
          return -1; //nameA comes first
        }
        if (nameA! > nameB!) {
          return 1; // nameB comes first
        }
        return 0;  // names must be equal
      });

  for (let entry of MSOT1DataSorted) {
      this.state.MSOutcome1Data.push({ "id": entry.id, "text": entry.text})
    }
    this.state.MSOutcome1Data.push({ id: "divider", type: ListBoxItemType.Divider });
    this.state.MSOutcome1Data.push({ id: "0", text: "Other / Not Listed"});
  
  for (let entry of MSOT2DataSorted) {
    this.state.MSOutcome2Data.push({ "id": entry.id, "text": entry.text})
  }
  this.state.MSOutcome2Data.push({ id: "divider", type: ListBoxItemType.Divider });
  this.state.MSOutcome2Data.push({ id: "0", text: "Other / Not Listed"});
  
  for (let entry of MSOT3DataSorted) {
    this.state.MSOutcome3Data.push({ "id": entry.id, "text": entry.text, "groupId": entry.groupId})
  }
  this.state.MSOutcome3Data.push({ id: "divider", type: ListBoxItemType.Divider, groupId: "0" });
  this.state.MSOutcome3Data.push({ id: "0", text: "Other / Not Listed", groupId: "0"});

  for (let entry of dataplaceholderOSP1) {
    this.state.MSOutcomeStatementP1Data.push({ "id": entry.id, "text": entry.text, "groupId": entry.groupId})
  }

  for (let entry of dataplaceholderOSP2) {
    this.state.MSOutcomeStatementP2Data.push({ "id": entry.id, "text": entry.text, "groupId": entry.groupId})
  }

  }

  public fetchT1Outcomes() {
  return(this.state.MSOutcome1Data)
  }
  public fetchT2Outcomes() {
  return(this.state.MSOutcome2Data)

}
public fetchT3Outcomes() {
  this.state.selectionDDT3.clear();
  this.setState({
    outcomeDDItemsT3: this.state.MSOutcome3Data.filter((val) => val.groupId == this.selectedItemIDT2.value || val.groupId == "0")
 });
  return ""

}

public resetT3Outcomes() {
  this.setState({outcomeDDItemsT3:[]}, this.fetchT3Outcomes)

}

  public async getAllMSOutcomeFieldValue() {
    const workItemFormService = await SDK.getService<IWorkItemFormService>(
      WorkItemTrackingServiceIds.WorkItemFormService
    );
    var outcomeonetext = (await workItemFormService.getFieldValue("Custom.MSOutcomeT1Text")).toString();
    this.setState({
        MSOutcome1: outcomeonetext
     });
     this.increment();
    var outcometwotext = (await workItemFormService.getFieldValue("Custom.MSOutcomeT2Text")).toString();
    this.setState({
        MSOutcome2: outcometwotext
    });
    this.increment();
    var outcomethreetext = (await workItemFormService.getFieldValue("Custom.MSOutcomeT3Text")).toString();
    this.setState({
        MSOutcome3: outcomethreetext
    });
    this.increment();
    var outcomestatement = (await workItemFormService.getFieldValue("Custom.MSOutcomeStatement")).toString();
    this.setState({
        MSOutcomeStatement: outcomestatement
  });
  if (outcomestatement){
    this.setState({
      ShowOutcomeCard: true,
      });
  }
}

  public render(): JSX.Element {
    if (this.state.IsRenderReady){
    return (
        <div style={{ width: "calc(100% - 30px)", alignItems: "center" }}>
          <div className="outcome-dd">
            <Dropdown 
                ariaLabel="Basic"
                className="MSOCT1"
                placeholder="Select Microsoft Outcome Tier 1"
                items={this.fetchT1Outcomes()}
                onSelect={this.onSelectT1}
            />
            </div>
              <Observer selectedItemIDT1={this.selectedItemIDT1}>
                {(props: { selectedItem: string}) => {
                  return null;
               }}
            </Observer>
            {this.state.MSOutcome1 ? <div className="current-outcome" >{this.state.MSOutcome1}</div> : <div className="required-field" >Microsoft Outcome Tier 1 cannot be empty.</div> }
            <div className="outcome-dd">
            <Dropdown
                ariaLabel="Basic"
                className="MSOCT2"
                placeholder="Select Microsoft Outcome Tier 2"
                items={this.fetchT2Outcomes()}
                onSelect={this.onSelectT2}
            />
            </div>
            <Observer selectedItemIDT2={this.selectedItemIDT2}>
                {(props: { selectedItem: string}) => {
              return null;
               }}
            </Observer>
            {this.state.MSOutcome2 ? <div className="current-outcome" >{this.state.MSOutcome2}</div> : <div className="required-field" >Microsoft Outcome Tier 2 cannot be empty.</div> }
            <div className="outcome-dd">
            <Dropdown 
                ariaLabel="Basic"
                className="MSOCT3"
                selection={this.state.selectionDDT3}
                placeholder="Select Microsoft Outcome Tier 3"
                disabled = {this.state.MSOutcomeT3DDIsDisabled}
                items={this.state.outcomeDDItemsT3}
                onSelect={this.onSelectT3}
            />
            </div>
            <Observer selectedItemTextT3={this.selectedItemTextT3}>
                {(props: { selectedItem: string}) => {
                    return null;
                }}
            </Observer>
            {
            this.state.MSOutcome3 ? <div className="current-outcome" >{this.state.MSOutcome3}</div> : <div className="required-field" >Microsoft Outcome Tier 3 cannot be empty.</div> 
            }
            <div className="learn-more">
            <Link href="https://microsoft.sharepoint.com/teams/SDMinADO/SitePages/How-To--Contribute-To-Outcomes.aspx" target="_blank" subtle={true} className="missingSomethingLink">
                Are we missing something?
            </Link>
            </div>
            {
                  this.state.ShowOutcomeCard? 
                  <div className="outcome-card">
                  <Card>
                   {this.state.MSOutcomeStatement}
                  </Card>
                  </div> : null
            }
            </div>
      );
    }
      else {
        return (<div className="flex-row"></div>)

      }
}

private onSelectT1 = async (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) => {
  this.selectedItemIDT1.value = item.id || "";
  this.selectedItemTextT1.value = item.text || "";
  const workItemFormService = await SDK.getService<IWorkItemFormService>(
    WorkItemTrackingServiceIds.WorkItemFormService
  );
   workItemFormService.setFieldValues({"Custom.MSOutcomeStatement": ""});
   workItemFormService.setFieldValues({"Custom.MSOutcomeT1ID": this.selectedItemIDT1.value, "Custom.MSOutcomeT1Text": this.selectedItemTextT1.value});
   this.getAllMSOutcomeFieldValue();
   this.createOutcomeStatement();

};
private onSelectT2 = async (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) => {
  this.selectedItemIDT2.value = item.id || "";
  this.selectedItemTextT2.value = item.text || "";
  const workItemFormService = await SDK.getService<IWorkItemFormService>(
    WorkItemTrackingServiceIds.WorkItemFormService
  );
   workItemFormService.setFieldValues({"Custom.MSOutcomeStatement": ""});
   workItemFormService.setFieldValues({"Custom.MSOutcomeT2ID": this.selectedItemIDT2.value, "Custom.MSOutcomeT2Text": this.selectedItemTextT2.value});
   this.getAllMSOutcomeFieldValue();
   workItemFormService.setFieldValues({"Custom.MSOutcomeT3ID": "", "Custom.MSOutcomeT3Text": ""});
   this.setState({
    MSOutcome3: "",
    ShowOutcomeCard: false
  });
  this.setState({
    MSOutcomeT3DDIsDisabled: false
    });
    this.fetchT3Outcomes();
    this.createOutcomeStatement();
};
private onSelectT3 = async (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) => {
  this.selectedItemIDT3.value = item.id || "";
  this.selectedItemTextT3.value = item.text || "";
  const workItemFormService = await SDK.getService<IWorkItemFormService>(
    WorkItemTrackingServiceIds.WorkItemFormService
  );
   workItemFormService.setFieldValues({"Custom.MSOutcomeStatement": ""});
   workItemFormService.setFieldValues({"Custom.MSOutcomeT3ID": this.selectedItemIDT3.value, "Custom.MSOutcomeT3Text": this.selectedItemTextT3.value});
   this.getAllMSOutcomeFieldValue();
   this.createOutcomeStatement();
};
public async createOutcomeStatement(){
    const workItemFormService = await SDK.getService<IWorkItemFormService>(
      WorkItemTrackingServiceIds.WorkItemFormService
    );
    var outcomestringbuilder = ""
    var outcomeoneID = (await workItemFormService.getFieldValue("Custom.MSOutcomeT1ID")).toString();
    var outcometwoID = (await workItemFormService.getFieldValue("Custom.MSOutcomeT2ID")).toString();
    var outcomethreeID = (await workItemFormService.getFieldValue("Custom.MSOutcomeT3ID")).toString();
    var outcomeidbuilder = outcomeoneID + outcometwoID + outcomethreeID
    if (outcomeoneID != "0" && outcometwoID != "0" && outcomethreeID != "0") {
    const a = this.state.MSOutcomeStatementP1Data
    const b = this.state.MSOutcomeStatementP2Data
    let MSOutcomeP1 = a.filter((val) => val.groupId == this.selectedItemIDT1.value)[0].text
    let MSOutcomeP2 = b.filter((val) => val.groupId == this.selectedItemIDT1.value)[0].text
    var concatedoutcome = outcomestringbuilder.concat(MSOutcomeP1 as string, this.selectedItemTextT2.value, MSOutcomeP2 as string, this.selectedItemTextT3.value )
    workItemFormService.setFieldValues({"Custom.MSOutcomeStatement" : concatedoutcome, "Custom.MSOutcomeStatementID" : outcomeidbuilder});
    this.setState({
      MSOutcomeStatement: (await workItemFormService.getFieldValue("Custom.MSOutcomeStatement")).toString()
      });
    } else {
      workItemFormService.setFieldValues({"Custom.MSOutcomeStatement" : "No outcome statement for this record.", "Custom.MSOutcomeStatementID" : "0"});
      this.setState({
        MSOutcomeStatement: (await workItemFormService.getFieldValue("Custom.MSOutcomeStatement")).toString()
        });
    }
    
    if(outcomeoneID && outcometwoID && outcomethreeID){
      this.setState({
        ShowOutcomeCard: true,
        });

    } else {
      this.setState({
        ShowOutcomeCard: false,
        });
    }

}

 }
 

export default WorkItemFormGroupComponent;

showRootComponent(<WorkItemFormGroupComponent />);




