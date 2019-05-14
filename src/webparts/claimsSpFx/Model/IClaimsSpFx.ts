import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
  
export interface IClaimsSpFx {
    selectedItems: any[];
    disputeClaim: string;
    Claimdescription: string;
    gpselectedItem?: { key: string | number | undefined };
    termKey?: string | number;
    gpselectedItems: IDropdownOption[];
    disableToggle: boolean;
    defaultChecked: boolean;
    pplPickerType:string;
    userIDs: number[];
    userManagerIDs: number[];
    hideDialog: boolean;
    status: string;
    isChecked: boolean;
    showPanel: boolean;
    required:string;
    onSubmission:boolean;
    termnCond:boolean;
}