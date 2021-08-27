import * as React from 'react';
import styles from './FileUpload1.module.scss';
import { IFileUpload1Props } from './IFileUpload1Props';
import { escape, isEmpty } from '@microsoft/sp-lodash-subset';

import { TextField, Stack, IDropdownOption, Dropdown, IDropdownStyles, 
  IStackStyles, DatePicker, Toggle, PrimaryButton, Label } from '@fluentui/react';
  import Service from './Service';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

const stackTokens = { childrenGap: 50 };
  const stackStyles: Partial<IStackStyles> = { root: { padding: 10 } };
  const stackButtonStyles: Partial<IStackStyles> = { root: { Width: 20 } };
  
  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 },
  };
  


  const options: IDropdownOption[] = [

  {key:'Select',text:'Select Operation'},
    { key: 'Add', text: 'Add' },
    { key: 'Update', text: 'Update' },
    { key: 'Remove', text: 'Remove' },
  ];


  export interface IFieldUpload1ControlFieldsState{
    operation:any;
    file:any;
    status:any;
    
  }



export default class FileUploadControl extends React.Component<IFileUpload1Props, IFieldUpload1ControlFieldsState> {

  public _service:any;
  public constructor(props:IFileUpload1Props){
    super(props);
    this.state={
      
      operation:null,
      file:null,
      status:null
      

    };
    this._service = new Service(this.props.url,this.props.context);
  }

  private changeChoice(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void{
    this.setState({operation:item});
  }

  private fileChangeHandler(event:any):void{
    //this.setState({date:data});
    this.setState({file:event.target.files[0]});
    

  }

  private OnBtnClick():void{

    
    if(this.state.operation==null || this.state.operation.key=='Select')
    {

      alert('please select any value');

    }

    else if(this.state.file==null)
    {
     
      alert('please select any file');
    }

    else if(this.state.file.name!="SC_Upload_Template.xlsx")
    {

      alert('please select file name with SC_Upload_Template.xlsx');


    }

    else
    {

    console.log(this.state);
    let inputData:any=
    {
      
      Operation: this.state.operation.text,
      Status:'NotCopied',
      //currentDateTime:Date().toLocaleString()
      
      
    };

    this._service.uploadFile(this.state.file,inputData);

  }
  }

 
  public render(): React.ReactElement<IFileUpload1Props> {
    return (

      <Stack tokens={stackTokens} styles={stackStyles}>
        <Stack>
          <b><div>Operations</div></b><br></br>

            <Dropdown
                placeholder="Select an option"
                options={options}
                styles={dropdownStyles}
                selectedKey={this.state.operation ? this.state.operation.key : undefined}
                onChange={this.changeChoice.bind(this)} 
            />
              <br/>

               <b><div id="test">File Name</div></b>
               <br/>

              <input type="file" name="file" onChange={this.fileChangeHandler.bind(this)} accept=".xlsx"  />


        </Stack>
        <Stack >
        <PrimaryButton text="Submit" onClick={this.OnBtnClick.bind(this)} styles={stackButtonStyles} className={styles.button}/>
        </Stack>
        </Stack>
      
    );
  }

}