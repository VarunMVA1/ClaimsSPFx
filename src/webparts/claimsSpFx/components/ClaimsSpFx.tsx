import * as React from 'react';
import styles from './ClaimsSpFx.module.scss';
import { IClaimsSpFxProps } from './IClaimsSpFxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IClaimsSpFx } from '../Model/IClaimsSpFx';
import { default as pnp, ItemAddResult } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';


export default class ClaimsSpFx extends React.Component<IClaimsSpFxProps, IClaimsSpFx> {
  constructor(props: IClaimsSpFxProps) {   
    super(props);
    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this._onCheckboxChange = this._onCheckboxChange.bind(this);
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.createItem = this.createItem.bind(this);
    this.onTaxPickerChange = this.onTaxPickerChange.bind(this);
    this._getManager = this._getManager.bind(this);
    this.state = {
      disputeClaim:"",
      Claimdescription:"",
      selectedItems: [],
      hideDialog: true,
      showPanel: false,
      gpselectedItem: undefined,
      gpselectedItems: [],  
      disableToggle:false,
      defaultChecked:false,
      termKey: undefined,
      userIDs: [],
      userManagerIDs: [],
      pplPickerType: "",
      status:"",
      isChecked: false,
      required:"This is required",
      onSubmission:false,
      termnCond:false
    };
  }
  
  public render(): React.ReactElement<IClaimsSpFxProps> {
    const { gpselectedItem, gpselectedItems } = this.state;
    const { disputeClaim, Claimdescription } = this.state;   
    pnp.setup({
      spfxContext: this.props.context
    });

    return (
      <form>
        <div className={styles.claimsSpFx}>
          <div className={styles.container}>
        <div className={`ms-Grid-row ms-bgColor-neutralLight ms-fontColor-white ${styles.row}`}>
        <div className="ms-Grid-col ms-u-sm4 block">
            <label className="ms-Label">Dispute Claim</label>             
        </div>
          <div className="ms-Grid-col ms-u-sm8 block">
              <TextField value={this.state.disputeClaim} required={true} onChanged={this.handleTitle}
          errorMessage={(this.state.disputeClaim.length === 0 && this.state.onSubmission === true) ? this.state.required : ""}/>
          </div>
            <div className="ms-Grid-col ms-u-sm4 block">
              <label className="ms-Label">Claim Description</label>
            </div>
            <div className="ms-Grid-col ms-u-sm8 block">
              <TextField multiline autoAdjustHeight value={this.state.Claimdescription} onChanged={this.handleDesc} />
            </div>
            <div className="ms-Grid-col ms-u-sm4 block">
              <label className="ms-Label">Dispute Location</label><br/>
            </div>
            <div className="ms-Grid-col ms-u-sm8 block">
            <TaxonomyPicker
              allowMultipleSelections={false}
              termsetNameOrID="Countries"
              panelTitle="Select Location"
              label=""
              context={this.props.context}
              onChange={this.onTaxPickerChange}
              isTermSetSelectable={false} />
              <p className={(this.state.termKey === undefined && this.state.onSubmission === true)? styles.fontRed : styles.hideElement}>This is required</p>
            </div>
            <div className="ms-Grid-col ms-u-sm4 block">
              <label className="ms-Label">Claims Group</label><br/>
            </div>
            <div className="ms-Grid-col ms-u-sm8 block">
              <Dropdown
                placeHolder="Select an Option"
                label=""
                id="component"
                selectedKey={gpselectedItem ? gpselectedItem.key : undefined}
                ariaLabel="Group dropdown"
                options={[
                  { key: 'Medical Care', text: 'Medical Care' },
                  { key: 'Health Care', text: 'Health Care' },
                  { key: 'Employee Refunds', text: 'Employee Refunds' }
                ]}
                onChanged={this._changeState}
                onFocus={this._log('onFocus called')}
                onBlur={this._log('onBlur called')}
                />
          </div>
          <div className="ms-Grid-col ms-u-sm4 block">
            <label className="ms-Label">Escalate To Higher Level?</label>
          </div>
          <div className="ms-Grid-col ms-u-sm8 block">
          <Toggle
            disabled={this.state.disableToggle}
            checked={this.state.defaultChecked}
            label=""
            onAriaLabel="This toggle is checked. Press to uncheck."
            offAriaLabel="This toggle is unchecked. Press to check."
            onText="On"
            offText="Off"
            onChanged={(checked) =>this._changeSharing(checked)}
            onFocus={() => console.log('onFocus called')}
            onBlur={() => console.log('onBlur called')}         
          />
          </div>
          <div className="ms-Grid-col ms-u-sm4 block">
            <label className="ms-Label">Reporting Manager</label>
          </div>
          <div className="ms-Grid-col ms-u-sm8 block">
            <PeoplePicker
              context={this.props.context}
              titleText=" "
              personSelectionLimit={1}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={false}
              isRequired={true}
              disabled={false}
              ensureUser={true}
              selectedItems={this._getManager}
              errorMessage={(this.state.userManagerIDs.length === 0 && this.state.onSubmission === true) ? this.state.required : " "} />
          </div>
          <div className={`ms-Grid-col ms-u-sm1 block ${styles.customFont}`}>
            <br/><Checkbox onChange={this._onCheckboxChange} ariaDescribedBy={'descriptionID'} color={`${styles.customFont}`} label="I have read and agree to the terms & condition" />
          </div>
          <div className="ms-Grid-col ms-u-sm11 block">
            {/* <span className={`${styles.customFont}`}>I have read and agree to the terms & condition</span><br/> */}
            <p className={(this.state.termnCond === false && this.state.onSubmission === true)? styles.fontRed : styles.hideElement}>Please check the Terms & Condition</p>
          </div>        
          <div className="ms-Grid-col ms-lg8 ms-md8 ms-u-sm6 block marginRight">
              <br/><PrimaryButton text="Create" onClick={() => { this.validateForm(); }} /> &nbsp;
              <DefaultButton text="Cancel" onClick={() => { this.setState({}); }} />
          </div>
          <div>
          <Panel
            isOpen={this.state.showPanel}
            type={PanelType.smallFixedFar}
            onDismiss={this._onClosePanel}
            isFooterAtBottom={false}
            headerText="Are you sure you want to create dispute claim ?"
            closeButtonAriaLabel="Close"
            onRenderFooterContent={this._onRenderFooterContent}
          ><span>Please check the details filled and click on Confirm button to create dispute claim.</span>
          </Panel>
        </div>
        <Dialog
            hidden={this.state.hideDialog}
            onDismiss={this._closeDialog}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: 'Request Submitted Successfully',
              subText: "" }}
              modalProps={{
              titleAriaId: 'myLabelId',
              subtitleAriaId: 'mySubTextId',
              isBlocking: false,
              containerClassName: 'ms-dialogMainOverride'            
              }}>
            <div dangerouslySetInnerHTML={{__html:this.state.status}}/>    
          <DialogFooter>
          <PrimaryButton onClick={()=>this.gotoHomePage()} text="Okay" />
          </DialogFooter>
        </Dialog>
        </div>
        </div>
        </div>
      </form>
    );
  }

  ​private onTaxPickerChange(terms : IPickerTerms) {
      this.setState({ termKey: terms[0].key.toString() });
      console.log("Terms", terms);
  }
  
  private _getManager(items: any[]) {
    this.state.userManagerIDs.length = 0;
    for (let item in items)
    {   
      this.state.userManagerIDs.push(items[item].id);
      console.log(items[item].id);
    }
  }
  
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this.createItem} style={{ marginRight: '8px' }}>
          Confirm
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }
  
  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }
  
  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }
  
  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }
  
  private _changeSharing(checked:any):void{
    this.setState({defaultChecked: checked});
  }
  
  private _changeState = (item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ gpselectedItem: item });
    if(item.text == "Dispute Claim")
    {
      this.setState({defaultChecked: false});
      this.setState({disableToggle: true});     
    }
    else
    {
      this.setState({disableToggle:false});
    }
  }
  
  private handleTitle(value: string): void {
    return this.setState({
      disputeClaim: value
    });
  }
  
  private handleDesc(value: string): void {
    return this.setState({
      Claimdescription: value
    });
  }
  
  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    this.setState({termnCond: (isChecked)?true:false});
  }
  
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  
  private _showDialog = (status:string): void => {   
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }
  
  private validateForm():void{
    let allowCreate: boolean = true;
    this.setState({ onSubmission : true });
    
    if(this.state.disputeClaim.length === 0)
    {
      allowCreate = false;
    }
    if(this.state.termKey === undefined)
    {
      allowCreate = false;
    }   
    
    if(allowCreate)
    {
      this._onShowPanel();
    }
    else
    {
      //do nothing
    } 
  }

  private createItem():void { 
    this._onClosePanel(); 
    this._showDialog("Submitting Request");
    console.log(this.state.termKey);
    pnp.sp.web.lists.getByTitle("Claims").items.add({
      'Title': this.state.disputeClaim,
      'Description': this.state.Claimdescription,
      'Group': this.state.gpselectedItem.key,
      'Escalate_To_Higher_Level': this.state.isChecked,
      'Reporting_ManagerId': this.state.userManagerIDs[0],
      'Location': {
        __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
        Label: "1",
        TermGuid: this.state.termKey,
        WssId: -1 
    },
  }).then((result: ItemAddResult) => {
      this.setState({ status: "Your request has been submitted sucessfully " });
  }, (error: any): void => {  
    this.setState({ status: "Error while creating the item: " + error});  
  });
  }
  
  private gotoHomePage():void{
    window.location.replace(this.props.siteUrl);
  }
}
