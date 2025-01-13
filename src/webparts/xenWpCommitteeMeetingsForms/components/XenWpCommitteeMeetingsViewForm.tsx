/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./XenWpCommitteeMeetingsForms.module.scss";
import "./CustomStyles/custom.css";
import type { IXenWpCommitteeMeetingsFormsProps } from "./IXenWpCommitteeMeetingsFormsProps";
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Icon,
  IconButton,
  Link,
  mergeStyleSets,
  Modal,
  PrimaryButton,
  SelectionMode,
  Spinner,
  SpinnerSize,
  TextField,
  // Toggle,
} from "@fluentui/react";
import { RichText } from "@pnp/spfx-controls-react/lib/controls/richText";
import { format } from "date-fns";
import PasscodeModal from "./passCode/passCode";

interface CommtteeMeetingsState {
  MeetingNumber: string;
  MeetingDate: string;
  MeetingLink: string;
  MeetingMode: string;
  MeetingSubject: string;
  MeetingStatus: string;
  Department: string;
  ConsolidatedPDFPath: string;
  CommitteeName: string;
  Chairman: any;
  chairmanObjectAfterFilter:any;
  CommitteeMeetingGuestMembersDTO: any;
  CommitteeMeetingMembersDTO: any;
  CommitteeMeetingNoteDTO: any;
  CommitteeMeetingMembers: any;
  CommitteeMeetingGuests: any;
  AuditTrail: any;
  StatusNumber: string;
  CurrentApprover: any;
  FinalApprover: any;
  PreviousApprover: any;
  Confirmation: any;
  actionBtn: string;
  hideCnfirmationDialog: boolean;
  hideSuccussDialog: boolean;
  hideWarningDialog: boolean;
  hideApprovalDialog:boolean;
  SuccussMsg: string;
  CommitteeMeetingMemberCommentsDT: any;
  comments: string;
  isRturn: boolean;
  isApproverBtn:any;
  Created: any;
  departmentAlias: any;
  meetingId: any;
   isPasscodeModalOpen: boolean;
   isPasscodeValidated: boolean;
   passCodeValidationFrom: any;
   isLoading:any;
   consolidatePdf:any;
   hideParellelActionAlertDialog:any;
   parellelActionAlertMsg:any;

}
const getIdFromUrl = (): any => {
  const params = new URLSearchParams(window.location.search);
  const Id = params.get("itemId");
  return Number(Id);
};


const Cutsomstyles = mergeStyleSets({
  modal: {
    padding: "10px",
    minWidth: "300px",
    maxWidth: "80vw",
    width: "100%",
    "@media (min-width: 768px)": {
      maxWidth: "580px",
    },
    "@media (max-width: 767px)": {
      maxWidth: "290px",
    },
    margin: "auto",
    backgroundColor: "white",
    borderRadius: "4px",
    boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",

    borderBottom: "1px solid #ddd",
    height: "50px",
  },
  headerTitle: {
    margin: "5px",
    marginLeft: "0px",
    fontSize: "16px",
    fontWeight: "400",
  },
  body: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
    padding: "20px 0",
    height: "100%",
    "@media (min-width: 768px)": {
      marginLeft: "20px",
      marginRight: "20px",
    },
    "@media (max-width: 767px)": {
      marginLeft: "20px",
      marginRight: "20px",
    },
  },
  footer: {
    display: "flex",
    alignItem: "center",
    justifyContent: "flex-end",

    borderTop: "1px solid #ddd",
    paddingTop: "12px",
    height: "50px",
  },
  button: {
    maxHeight: "32px",
  },
});

export default class XenWpCommitteeMeetingsViewForm extends React.Component<
  IXenWpCommitteeMeetingsFormsProps,
  CommtteeMeetingsState
> {
  private _listName;

  private _currentUserEmail = this.props.context.pageContext.user.email;

  constructor(props: any) {
    super(props);
    this.state = {
      departmentAlias: "",
      meetingId: "",

      MeetingNumber: "",
      MeetingDate: "",
      MeetingLink: "",
      MeetingMode: "",
      MeetingSubject: "",
      MeetingStatus: "",
      Department: "",
      ConsolidatedPDFPath: "",
      CommitteeName: "",
      Chairman: null,
      chairmanObjectAfterFilter:[],
      CommitteeMeetingGuestMembersDTO: [],
      CommitteeMeetingMembersDTO: [],
      CommitteeMeetingNoteDTO: [],
      CommitteeMeetingMembers: [],
      CommitteeMeetingGuests: [],
      AuditTrail: [],
      StatusNumber: "",
      CurrentApprover: null,
      FinalApprover: null,
      PreviousApprover: null,
      Confirmation: {
        Confirmtext: "",
        Description: "",
      },
      actionBtn: "",
      hideCnfirmationDialog: true,
      hideSuccussDialog: true,
      hideWarningDialog: true,
      hideApprovalDialog:true,
      SuccussMsg: "",
      CommitteeMeetingMemberCommentsDT: [],
      comments: "",
      isRturn: false,
      isApproverBtn:true,
      Created: null,
        isPasscodeModalOpen: false,
        isPasscodeValidated: false, 
        passCodeValidationFrom: "",
        isLoading:true,
        consolidatePdf:[],

      hideParellelActionAlertDialog:false,
      parellelActionAlertMsg:''
    };
    const listName = this.props.listName;
    this._listName = listName?.title;
    this._getItemBy();
    console.log(this._currentUserEmail)
  }

  public componentDidMount() {
   
    
    setTimeout(() => {
      this.setState({isLoading:false})
    }, 3000);
  }
  private stylesModal = mergeStyleSets({
    modal: {
      minWidth: "300px",
      maxWidth: "80vw",
      width: "100%",
      "@media (min-width: 768px)": {
        maxWidth: "580px",
      },
      "@media (max-width: 767px)": {
        maxWidth: "290px", 
      },
      margin: "auto",
      padding: "10px",
      backgroundColor: "white",
      borderRadius: "4px",
      boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
    },
    header: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      borderBottom: "1px solid #ddd",
      minHeight: "50px",
      padding: "5px",
    },
    headerTitle: {
      margin: "5px",
      marginLeft: "5px",
      fontSize: "16px",
      fontWeight: "400",
    },
    headerIcon: {
      paddingRight: "0px", 
    },
    body: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      textAlign: "center",
      padding: "20px 0",
      height: "100%",
      "@media (min-width: 768px)": {
        marginLeft: "20px",
        marginRight: "20px", 
      },
      "@media (max-width: 767px)": {
        marginLeft: "20px",
        marginRight: "20px",
      },
    },
    footer: {
      display: "flex",
      justifyContent: "space-between", 

      borderTop: "1px solid #ddd",
      paddingTop: "10px",
      minHeight: "50px",
    },
    button: {
      maxHeight: "32px",
      flex: "1 1 50%",
      margin: "0 5px", 
    },
    buttonContent: {
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
    },
    buttonIcon: {
      marginRight: "4px", 
    },

    removeTopMargin: {
      marginTop: "4px",
      marginBottom: "14px",
      fontWeight: "400",
    },
  });


  private _filterChairmanDataFromCommitteeMembersDTO = (data:any):any=>{
    const committeeMemberDTOFromList = JSON.parse(data)

    committeeMemberDTOFromList.filter((each:any)=>{
      if (each.isChairman === true){
        console.log(each)
        this.setState({chairmanObjectAfterFilter:each})
      }
    })
    const committeeMemberDTOWithOutChairmansData = committeeMemberDTOFromList.filter((each:any)=>each.isChairman === false)
  
    return committeeMemberDTOWithOutChairmansData



  }

  private _getItemDataFromList = async (itemId:any)=>{
    const item = await this.props.sp.web.lists
    .getByTitle(this._listName)
    .items.getById(Number(itemId))
    .select(`*,Created,Author/Title,Author/EMail,
      Editor/Title,
      CurrentApprover/Title,
      CurrentApprover/EMail,
      CurrentApprover/JobTitle,
      FinalApprover/Title,
      FinalApprover/EMail,
      FinalApprover/JobTitle,
      PreviousApprover/Title,
      Chairman/Title,
      Chairman/EMail,
      PreviousApprover/EMail`).expand(`Author,Editor,
   CurrentApprover,PreviousApprover,FinalApprover,Chairman`)();
   return item
  }

  private _getItemBy = async () => {
     await this.props.sp?.web.currentUser();
    const itemId = getIdFromUrl();
    const item = await this._getItemDataFromList(itemId)


    console.log(item,"Item fetched from the List")
  

    if (item) {
      if(item.StatusNumber === "9000"){
    await  this. _getConslidatePdf(item.Title)
  }

      this.setState({
        meetingId: item.Title,
        MeetingNumber: item.MeetingNumber,
        MeetingDate: item.MeetingDate
          ? new Date(item.MeetingDate).toLocaleDateString()
          : "",
        MeetingLink: item.MeetingLink,
        MeetingMode: item.MeetingMode,
        MeetingSubject: item.MeetingSubject,
        MeetingStatus: item.MeetingStatus,
        Department: item.Department,
        ConsolidatedPDFPath: item.MeetingNumber,
        CommitteeName: item.CommitteeName,
        Chairman:
          item.Chairman === null && item.ChairmanId === null
            ? null
            : item.Chairman,
        CommitteeMeetingGuestMembersDTO:
          item.CommitteeMeetingGuestMembersDTO === null
            ? []
            : JSON.parse(item.CommitteeMeetingGuestMembersDTO),
        CommitteeMeetingMembersDTO:
          item.CommitteeMeetingMembersDTO === null
            ? []
            : this._filterChairmanDataFromCommitteeMembersDTO(item.CommitteeMeetingMembersDTO), //CommitteeMeetingMemberCommentsDTO
        CommitteeMeetingMemberCommentsDT:
          item.CommitteeMeetingMemberCommentsDT === null
            ? []
            : JSON.parse(item.CommitteeMeetingMemberCommentsDT), //CommitteeMeetingMemberCommentsDTO
        CommitteeMeetingNoteDTO:
          item.CommitteeMeetingNoteDTO === null
            ? []
            : JSON.parse(item.CommitteeMeetingNoteDTO),
        CommitteeMeetingMembers:
          item.CommitteeMeetingMembers === null
            ? []
            : item.CommitteeMeetingGuestMembersDTO,
        CommitteeMeetingGuests: [],
        AuditTrail: item.AuditTrail === null ? [] : JSON.parse(item.AuditTrail),
        StatusNumber: item.StatusNumber,
        CurrentApprover:
          item.CurrentApprover === null && item.CurrentApproverId === null
            ? null
            : item.CurrentApprover,
        FinalApprover:
          item.FinalApprover === null && item.FinalApproverId === null
            ? null
            : item.FinalApprover,
        PreviousApprover:
          item.PreviousApprover === null && item.PreviousApproverId === null
            ? null
            : item.PreviousApprover,
        Created:
          new Date(item.Created).toLocaleDateString() +
          " " +
          new Date(item.Created).toLocaleTimeString(),
      });

      return item.StatusNumber;
    }
  };

  private _getConslidatePdf=async (folderName:string)=>{
    const foldername=folderName.replace(/\//g, "-");
    const url = `${this.props.context.pageContext.web.serverRelativeUrl}/CommitteeMeetingDocuments/${foldername}`;
    try {
      const folderItemsPdf = await this.props.sp.web
        .getFolderByServerRelativePath(`${url}`)
        .files.select("*")
        .expand("Author", "Editor")()
        .then((res:any) => res);
     
      this.setState({consolidatePdf:folderItemsPdf})
  }catch(err){
    return err
  }
 
}




  private columnsCommitteeMembers: IColumn[] = [
    {
      key: "memberName",
      name: "Member Name",
      fieldName: "memberEmailName",
      minWidth: 60,
      maxWidth: 250,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 100,
      maxWidth: 180,
      isResizable: true,
    },


    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',
      minWidth: 100,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: any) => {
    
        let iconName = '';
       
        switch (item.status) {
         
          case "Pending": 
            iconName = 'AwayStatus';
            break;
          case 'Waiting':
            iconName = 'Refresh';
            break;
          case 'Approved':
            iconName = 'CompletedSolid';
            break;
         
          case 'Returned':
            iconName = 'ReturnToSession';
            break;
      
          default:
            iconName = 'AwayStatus';
            break;
        }
    
        return (
          <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'center' }}>
            <Icon iconName={iconName} />
            <span style={{ marginLeft: '8px', lineHeight: '24px' }}>{item.status}</span>
          </div>
        );
      },
    },
    {
      key: "actionDate",
      name: "Action Date",
      fieldName: "actionDate",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
  ];


  private isReturnChecked = (
    event: React.ChangeEvent<HTMLInputElement>,
    checked?: boolean
  ) => {
    const isChecked = event.target.checked;

  this.setState({
    isRturn: isChecked,  // Updates isRturn based on checkbox state
    isApproverBtn: !isChecked  // If isRturn is true, hide the Approver button, else show it
  });
  };

  private columnsCommitteeGuestMembers: IColumn[] = [
    {
      key: "guestMemberName",
      name: "Guest Members Name",
      fieldName: "memberEmailName",
      minWidth: 150,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 290,
      isResizable: true,
    },
  ];


  private columnsCommitteeMeetingMinutes: IColumn[] = [
    {
      key: "serialNo",
      name: "S.No",
      fieldName: "serialNo",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      onRender: (_item: any, _index?: number) => (
        <span>{(_index !== undefined ? _index : 0) + 1}</span>
      ),
    },
    {
      key: "noteTitle",
      name: "Note#",
      fieldName: "noteTitle",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "committeeName",
      name: "Committee Name",
      fieldName: "committeeName",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "department",
      name: "Department",
      fieldName: "department",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "meetingMinutes",
      name: "Meeting Minutes",
      fieldName: "meetingMinutes",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: any) => (
        <RichText
          value={item.mom}
          isEditMode={false}
          style={{ minHeight: "auto", padding: "8px" }} 
        />
      ),
    },
    {
      key: "noteLink",
      name: "Note Link",
      fieldName: "noteLink",
      minWidth: 150,
      maxWidth: 400,
      isResizable: true,
      onRender(item, index, column) {
        return (

          <a
          href={item.noteLink} 
          target="_blank"
          rel="noopener noreferrer"
          data-interception="off"
          className={styles.notePdfCustom}
        >
          {item?.link}
        </a>
      
        );
      },
    },
  ];


  private columnsCommitteeComments: IColumn[] = [
    {
      key: "comments",
      name: "Comments",
      fieldName: "comments",
      minWidth: 200, 
      maxWidth: 550,
      isResizable: true,
    },
    {
      key: "commentedBy",
      name: "Commented by",
      fieldName: "commentedBy",
      minWidth: 200,
      maxWidth: 550,
      isResizable: true,
    },
  ];


  private columnsCommitteeWorkFlowLog: IColumn[] = [
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 100,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: "actionBy",
      name: "Action By",
      fieldName: "actionBy",
      minWidth: 100,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: "actionDate",
      name: "Action Date",
      fieldName: "actionDate",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
  ];

  private _formatDateTime = (date: string | number | Date) => {
    const formattedDate = format(new Date(date), "dd-MMM-yyyy");
    const formattedTime = format(new Date(date), "hh:mm a");
    return `${formattedDate} ${formattedTime}`;
  };

  private onClickMemberApprove = () => {
    this.setState({
      Confirmation: {
        Confirmtext: "Are you sure you want to approve this meeting?",
        Description: "Please click on Confirm button to approve meeting.",
      },
      hideCnfirmationDialog: false,

      actionBtn: "mbrApprove",
    });
  };
  private onClickMemberReturn = () => {
    if (this.state.comments===''){
      this.setState({
        hideWarningDialog: false,
      });

    }else if (!this.state.isPasscodeValidated) {
      this.setState({
        isPasscodeModalOpen: true,
        passCodeValidationFrom: "7000",
      }); 
      return; 
    }

    

   
   
   
  };
  private onClickChairman = () => {
    this.setState({
      Confirmation: {
        Confirmtext: "Are you sure you want to approve this meeting?",
        Description: "Please click on Confirm button to approve meeting.",
      },
      hideCnfirmationDialog: false,
      actionBtn: "chairmanApprove",
    });
  };

  private handleApproveByMembers = async () => {
    this.setState((prevState) => ({
      isLoading: true,
      hideCnfirmationDialog: !prevState.hideCnfirmationDialog,
    }));
    
    const itemId = getIdFromUrl();
    const itemFromList = await this._getItemDataFromList(itemId);
    console.log(itemFromList);

    const _CommitteeMemberDTO = JSON.parse(itemFromList?.CommitteeMeetingMembersDTO).filter(
      (each:any)=>each.isChairman === false
    );
    console.log(_CommitteeMemberDTO)
    const updatedCurrentApprover = _CommitteeMemberDTO?.map(
      (obj:any,index:any) => {

        if (
          obj.memberEmail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase()
        ) {
          return {
            
            ...obj,
            status: "Approved",
            statusNumber: "9000",
            actionDate:  this. _formatDateTime(new Date())
          };
        } else {
          return obj;
        }
      }
    );
    const isApprovedByAll = updatedCurrentApprover?.every(
      (obj: { status: string }) => obj.status === "Approved"
    );
  
    const checkingIsChairmanAvailable = updatedCurrentApprover.some((each:any)=>each.isChairman === true
    )
    console.log(checkingIsChairmanAvailable)

    if (!checkingIsChairmanAvailable){
      updatedCurrentApprover.push(this.state.chairmanObjectAfterFilter)
    }

    
    console.log(updatedCurrentApprover)
   

    
    const auditTrail = JSON.parse(itemFromList?.AuditTrail)
    const comments =  itemFromList.CommitteeMeetingMemberCommentsDT === null
    ? []
    : JSON.parse(itemFromList.CommitteeMeetingMemberCommentsDT)
   
   

    auditTrail.push({
      action: `Committee meeting approved by ${this.props.userDisplayName}`,
      actionBy: this.props.userDisplayName,
      actionDate: this. _formatDateTime(new Date()),
    });
    comments.push({
      comments: this.state.comments,
      commentedBy: this.props.userDisplayName,
      createdDate: new Date().toLocaleDateString(),
    });
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(getIdFromUrl())
      .update({
        startProcessing:isApprovedByAll,
        AuditTrail: JSON.stringify(auditTrail),
        PreviousApproverId:(await this.props.sp?.web.currentUser())?.Id,
        MeetingStatus: isApprovedByAll
          ? "Pending Chairman Approval"
          : this.state.MeetingStatus,
        StatusNumber: isApprovedByAll ? "6000" : itemFromList?.StatusNumber,
        CommitteeMeetingMembersDTO: JSON.stringify(updatedCurrentApprover),
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    if (item) {
      this.setState((prevState) => ({
        isLoading: false,
        hideSuccussDialog: !prevState.hideSuccussDialog,
        SuccussMsg: "Committee meeting has been approved successfully",
      }));
      
    }
  };

  private handleReturnByMembers = async () => {
    this.setState((prevState) => ({
      isLoading: true,
      hideCnfirmationDialog: !prevState.hideCnfirmationDialog,
    }));
    

    const itemId = getIdFromUrl();
    const itemFromList = await this._getItemDataFromList(itemId);
    console.log(itemFromList);

    const _CommitteeMemberDTO = JSON.parse(itemFromList?.CommitteeMeetingMembersDTO).filter(
      (each:any)=>each.isChairman === false
    );
    
    console.log(_CommitteeMemberDTO)
    const updatedCurrentApprover =_CommitteeMemberDTO?.map(
      (obj: any) => {
        if (
          obj.memberEmail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase() 
        ) {
          return {
            ...obj,
            status: "Returned",
            actionDate:  this. _formatDateTime(new Date())
          };
        } else {
          return obj;
        }
      }
    );
    const checkingIsChairmanAvailable = updatedCurrentApprover.some((each:any)=>each.isChairman === true
  )
  console.log(checkingIsChairmanAvailable)

  if (!checkingIsChairmanAvailable){
    updatedCurrentApprover.push(this.state.chairmanObjectAfterFilter)
  }

    console.log(updatedCurrentApprover)
    const auditTrail = JSON.parse(itemFromList?.AuditTrail)
 const comments = this.state.CommitteeMeetingMemberCommentsDT;

    auditTrail.push({
      action: `Committee meeting returned by ${this.props.userDisplayName}`,
      actionBy: this.props.userDisplayName,
      actionDate:  this. _formatDateTime(new Date())
    });
    comments.push({
      comments: this.state.comments,
      commentedBy: this.props.userDisplayName,
      createdDate: this. _formatDateTime(new Date()),
    });
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(getIdFromUrl())
      .update({
        startProcessing: true,
        AuditTrail: JSON.stringify(auditTrail),
        CommitteeMeetingMemberCommentsDT: this.state.comments
          ? JSON.stringify(comments)
          : null,
        MeetingStatus: "Returned",
        StatusNumber: "7000",
        CommitteeMeetingMembersDTO: JSON.stringify(updatedCurrentApprover),
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    if (item) {
      this.setState((prevState) => ({
        hideSuccussDialog: !prevState.hideSuccussDialog,
        isLoading: false,
        SuccussMsg: "Committee meeting has been returned successfully",
      }));
      
    }
  };

  private handleApproveByChairman = async () => {
    this.setState((prevState) => ({
      isLoading: true,
      hideCnfirmationDialog: !prevState.hideCnfirmationDialog,
    }));
    

    const itemId = getIdFromUrl();
    const itemFromList = await this._getItemDataFromList(itemId);
    console.log(itemFromList);
    const auditTrail = JSON.parse(itemFromList?.AuditTrail)
    const comments = this.state.CommitteeMeetingMemberCommentsDT;

    auditTrail.push({
      action: `Committee meeting approved by Chairman`,
      actionBy: this.props.userDisplayName,
      actionDate:  this. _formatDateTime(new Date())
    });
    comments.push({
      comments: this.state.comments,
      commentedBy: this.props.userDisplayName,
      createdDate: new Date().toLocaleDateString(),
    });
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(getIdFromUrl())
      .update({
        startProcessing: true,
        AuditTrail: JSON.stringify(auditTrail),
        CommitteeMeetingMemberCommentsDT: this.state.comments
          ? JSON.stringify(comments)
          : null,

        MeetingStatus: "Approved",
        StatusNumber: "9000",
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    if (item) {
      this.setState((prevState) => ({
        hideSuccussDialog: !prevState.hideSuccussDialog,
        isLoading: false,
        SuccussMsg: "Committee meeting has been approved successfully",
      }));
      
    }
  };

  private handleComments = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: any
  ) => {
    this.setState({
      comments: newValue,
    });
  };

  private onConfirmation = async () => {
    const itemId = getIdFromUrl();
    const _StatusNumber = await this._getItemDataFromList(itemId);
    if(_StatusNumber !=="7000"){
    switch (this.state.actionBtn) {
      case "mbrApprove":
        this.handleApproveByMembers();
        break;
      case "mbrReturn":
        this.handleReturnByMembers();

        break;
      case "chairmanApprove":
        this.handleApproveByChairman();

        break;

      default:
        break;
    }
  }else{
    this.setState({hideApprovalDialog:false,hideCnfirmationDialog:true});
  }
  };

  


  public handlePasscodeSuccess = (): void => {
    this.setState(
        { isPasscodeValidated: true, isPasscodeModalOpen: false },
        () => {
            if (this.state.passCodeValidationFrom === "7000") { 
                if (this.state.comments) {
                    this.setState({
                        Confirmation: {
                            Confirmtext: "Are you sure you want to return this meeting?",
                            Description: "Please click on Confirm button to return meeting.",
                        },
                        hideCnfirmationDialog: false,
                        actionBtn: "mbrReturn",
                    });
                }
            }
        }
    );
};


  private _openDocumentinTab=(url:string)=>{
    const fileUrl = window.location.protocol + "//" + window.location.host+url
window.location.href=fileUrl;

  }
  public render(): React.ReactElement<IXenWpCommitteeMeetingsFormsProps> {
    console.log(this.state)
    return (
      <div>
        <div className={styles.titleContainer}>
          <div className={`${styles.noteTitle}`}>
            <div className={styles.statusContainer}>
              {
                <p className={styles.status}>
                  Status: {this.state.MeetingStatus}{" "}
                </p>
              }
            </div>
            <h1 className={styles.title}>
              {getIdFromUrl()
                ? `eCommittee Meeting -${this.state.meetingId}`
                : `eCommittee Meeting -${this.props.formType}`}
            </h1>

            <p className={styles.titleDate}>Created : {this.state.Created}</p>
          </div>
        </div>
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            General Section
          </h1>
        </div>

        <div
          className={`${styles.generalSection}`}
          style={{
            flexGrow: 1,
            margin: "10 10px",
            boxSizing: "border-box",
          }}
        >
          {/* Meeting ID: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_MeetinngId">
              Meeting ID :<span className={styles.warning}>*</span>
            </label>
            <TextField
            id="_MeetinngId"
              type="text"
              className={styles.textField}
              value={this.state.meetingId}
              readOnly
            />
          </div>

          {/* Committee Name Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_CommitteeName">
              Committee Name :<span className={`${styles.warning}`}>*</span>
            </label>
            <TextField
            id="_CommitteeName"
              type="text"
              className={styles.textField}
              value={this.state.CommitteeName}
              readOnly
            />
          </div>

          {/* Convenor Department : Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_ConvenorDpt">
              Convenor Department :<span className={styles.warning}>*</span>
            </label>
            <TextField
            id="_ConvenorDpt"
              type="text"
              className={styles.textField}
              value={this.state.Department}
            />
          </div>

          {/* Chairman: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_Chairman">
              Chairman :<span className={styles.warning}>*</span>
            </label>
            <TextField
            id="_Chairman"
              type="text"
              className={styles.textField}
              value={this.state.Chairman?.Title || ""}
            />
          </div>

          {/* Meeting Date: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_MeetingDate">
              Meeting Date :<span className={styles.warning}>*</span>
            </label>
            <TextField
              type="text"
              id="_MeetingDate"
              className={styles.textField}
              value={this.state.MeetingDate}
              readOnly
            />

          </div>

          {/* Meeting Subject: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_MeetingSubject">
              Meeting Subject :<span className={styles.warning}>*</span>
            </label>
            <textarea
            id="_MeetingSubject"
              className={styles.textarea}
              value={this.state.MeetingSubject}
              readOnly
            >
              {" "}
            </textarea>
          </div>

          {/* Meeting Mode : Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_MeetingMode">
              Meeting Mode :<span className={`${styles.warning}`}>*</span>
            </label>
            <TextField
            id="_MeetingMode"
              type="text"
              className={styles.textField}
              value={this.state.MeetingMode}
              readOnly
            />
          </div>

          {/* Meeting Link: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_MeetingLink">
              Meeting Link :<span className={styles.warning}>*</span>
            </label>
            <div className={styles.parentContainer}>
              <Link
              id="_MeetingLink"
                className={styles.meetingLink}
                onClick={() => window.open(this.state.MeetingLink, "_blank")}
              >
                {this.state.MeetingLink}
              </Link>
            </div>
          </div>
        </div>

        {/* Committee Members section */}

        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Committee Members
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.CommitteeMeetingMembersDTO} // Data for the table
                columns={this.columnsCommitteeMembers} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>
        {/* Committee Guest  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Committee Guest Members
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.CommitteeMeetingGuestMembersDTO} // Data for the table
                columns={this.columnsCommitteeGuestMembers} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>
        {this.state.StatusNumber ==="9000"  &&this.state.StatusNumber&&(
          <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Consolidated Pdf Link
          </h1>
        </div>
        )}
         {this.state.StatusNumber ==="9000" &&this.state.StatusNumber &&(
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
           <div>Consolidated pdf link: {this.state.consolidatePdf.length>0 &&(<Link onClick={()=>this._openDocumentinTab(this.state.consolidatePdf[0].ServerRelativeUrl)} style={{"wordBreak":"break-all"}}>{this.state.consolidatePdf[0].Name}</Link>)}</div>
          </div>
        </div>
        )}
   
        {/* Meeting Minutes  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Meeting Minutes
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            <div style={{ overflowX: "auto", width: "100%" }}>
              <DetailsList
                items={this.state.CommitteeMeetingNoteDTO} 
                columns={this.columnsCommitteeMeetingMinutes} 
                layoutMode={DetailsListLayoutMode.fixedColumns} 
                selectionMode={SelectionMode.none} 
                isHeaderVisible={true} 
              />
            </div>
          </div>
        </div>
        {this.state.CommitteeMeetingMembersDTO.some(
          (obj:any) =>
            obj.memberEmail.toLowerCase() ===
            this.props.context.pageContext.user.email.toLowerCase() && ( obj.status ==="Pending" ||  obj.status ==="Waiting")
        ) &&
          this.state.StatusNumber === "5000" && (
            <div
              className={`${styles.generalSectionMainContainer}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <h1 className={styles.viewFormHeaderSectionContainer}>
                Comments section
              </h1>
            </div>
          )}
        {this.state.CommitteeMeetingMembersDTO.some(
          (obj:any) =>
            obj.memberEmail.toLowerCase() ===
            this.props.context.pageContext.user.email.toLowerCase() && ( obj.status ==="Pending" ||  obj.status ==="Waiting")
        ) &&
          this.state.StatusNumber === "5000" && (
            <div
              className={`${styles.generalSectionApproverDetails}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
          <div>
        
        <label htmlFor="returnCheckbox">Do you want to return?</label>
        <br />
        <div className={styles.sliderContainer}>
        <input
          type="checkbox"
          id="returnCheckbox"
          checked={this.state.isRturn}
          onChange={this.isReturnChecked}
          aria-checked={this.state.isRturn}
          className={styles.hiddenCheckbox}  
        />  
        <label htmlFor="returnCheckbox" className={styles.toggleSwitch} aria-label="Toggle return checkbox">
          <span className={styles.slider} />
        </label>

      
        </div>
      </div>


              
              {this.state.isRturn && (
                <div>
                  <label className={styles.label} htmlFor="_ApproversReturnCmt">Comments <span className={styles.warning}>*</span> :</label>
                  <TextField
                  id="_ApproversReturnCmt"
                    multiline
                    value={this.state.comments}
                    onChange={this.handleComments}
                    placeholder="Add Comment"
                  ></TextField>
                </div>
              )}
            </div>
          )}

        {this.state.Chairman?.EMail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase() &&
          this.state.StatusNumber === "6000" && (
            <div
              className={`${styles.generalSectionMainContainer}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <h1 className={styles.viewFormHeaderSectionContainer}>
                Comments section
              </h1>
            </div>
          )}
        {this.state.Chairman?.EMail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase() &&
          this.state.StatusNumber === "6000" && (
            <div
              className={`${styles.generalSectionApproverDetails}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <label className={styles.label} htmlFor="_ApproversCmt">Comments :</label>

              <TextField
              id="_ApproversCmt"
                multiline
                value={this.state.comments}
                onChange={this.handleComments}
                placeholder="Add Comment"
              ></TextField>
            </div>
          )}

        {/* Comments section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>Comments</h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.CommitteeMeetingMemberCommentsDT} // Data for the table
                columns={this.columnsCommitteeComments} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>

        {/* WorkFlow  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Workflow Log
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.AuditTrail} // Data for the table
                columns={this.columnsCommitteeWorkFlowLog} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>

        {/*  Buttons Section */}

        <div className={styles.buttonSectionContainer}>
            <span
              hidden={
                !(
                  this.state.CommitteeMeetingMembersDTO.some(
                    (obj:any) =>
                      obj.memberEmail.toLowerCase() ===this.props.context.pageContext.user.email.toLowerCase() && ( obj.status ==="Pending" ||  obj.status ==="Waiting")
                  ) && this.state.StatusNumber === "5000" && !this.state.isRturn
                )
              }
            >
              <PrimaryButton
                onClick={async ()=>{
                  const itemId = getIdFromUrl();
                  if (itemId){
                    
                    const item = await this._getItemDataFromList(itemId);
                    console.log(item);

                    const _CommitteeMemberDTO = JSON.parse(item?.CommitteeMeetingMembersDTO);
                    console.log(_CommitteeMemberDTO)


                   const committeMemberDTO =  _CommitteeMemberDTO.filter(
                      (each:any)=>each.memberEmail === this._currentUserEmail
                    )[0]
                  
                    if (item?.StatusNumber === '7000'){

                      this.setState({
                     
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This request has been ${item?.MeetingStatus.toLowerCase()}.`,
                      });


                    return

                    }



                    if (committeMemberDTO.status !== 'Pending'){

                      this.setState({
                     
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This request has been ${item?.MeetingStatus.toLowerCase()}.`,
                      });


                    return

                    }

                  }

                  

                  this.onClickMemberApprove()

                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "DocumentApproval" }}
              >
                Approve
              </PrimaryButton>
            </span>
          {/* )} */}

          <span
            hidden={
              !(
                this.state.CommitteeMeetingMembersDTO.some(
                  (obj:any) =>
                    obj.memberEmail.toLowerCase() ===this.props.context.pageContext.user.email.toLowerCase() && 
                    ( obj.status ==="Pending" ||  obj.status ==="Waiting")
                ) &&
                this.state.StatusNumber === "5000" &&
                this.state.isRturn
              )
            }
          >
            <PrimaryButton
              onClick={
                async ()=>{
                  const itemId = getIdFromUrl();
                  if (itemId){
                    
                    const item = await this._getItemDataFromList(itemId);
                    console.log(item);

                    const _CommitteeMemberDTO = JSON.parse(item?.CommitteeMeetingMembersDTO);
                    console.log(_CommitteeMemberDTO)


                   const committeMemberDTO =  _CommitteeMemberDTO.filter(
                      (each:any)=>each.memberEmail === this._currentUserEmail
                    )[0]


                    if (item?.StatusNumber === '7000'){

                      this.setState({
                     
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This request has been ${item?.MeetingStatus.toLowerCase()}.`,
                      });


                    return

                    }

                    if (committeMemberDTO.status !== 'Pending'){

                      this.setState({
                     
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This request has been ${item?.MeetingStatus.toLowerCase()}.`,
                      });


                    return

                    }


                    

                   
                    

                  }


                  this.onClickMemberReturn()
                }
                
               }
              className={`${styles.responsiveButton} `}
              iconProps={{ iconName: "ReturnToSession" }}
            >
              Return
            </PrimaryButton>
          </span>

          
          <span
          hidden={
            !(
              this.state.Chairman?.EMail.toLowerCase() ===
                this.props.context.pageContext.user.email.toLowerCase() &&
              this.state.StatusNumber === "6000" && !this.state.isRturn 
              
            )
          }
        >
          <PrimaryButton
            onClick={
              async ()=>{
                const itemId = getIdFromUrl();
                if (itemId){
                  
                  const item = await this._getItemDataFromList(itemId);
                  console.log(item);

                  if (item?.StatusNumber !== '6000'){

                    this.setState({
                   
                      hideParellelActionAlertDialog: true,
                      parellelActionAlertMsg:
                        `This request has been ${item?.MeetingStatus.toLowerCase()}.`,
                    });


                  return

                  }
                  this.onClickChairman()

                }
              }
              
              
              }
            className={`${styles.responsiveButton} `}
            iconProps={{ iconName: "DocumentApproval" }}
          >
            Approve
          </PrimaryButton>
        </span>
          <DefaultButton
            // type="button"
            onClick={() => {
              const pageURL: string = this.props.homePageUrl;
              window.location.href = `${pageURL}`;
            }}
            className={`${styles.responsiveButton} `}
            iconProps={{ iconName: "Cancel" }}
          >
            Exit
          </DefaultButton>
        </div>
        <Modal
          isOpen={!this.state.hideCnfirmationDialog}
          onDismiss={() =>
            this.setState({
              hideCnfirmationDialog: false,
            })
          }
          isBlocking={true}
          containerClassName={this.stylesModal.modal}
        >
          <>
            <div className={this.stylesModal.header}>
              <div style={{ display: "flex", alignItems: "center" }}>
                <IconButton iconProps={{ iconName: "WaitlistConfirm" }} />
                <h4 className={this.stylesModal.headerTitle}>Confirmation</h4>
              </div>
              <IconButton
                iconProps={{ iconName: "Cancel" }}
                onClick={() =>
                  this.setState({
                    hideCnfirmationDialog: true,
                  })
                }
              />
            </div>
            {this.state.Confirmation && (
              <div className={this.stylesModal.body}>
                <p className={`${this.stylesModal.removeTopMargin}`}>
                  {this.state.Confirmation.Confirmtext}
                </p>
                <br />
                <p className={`${this.stylesModal.removeTopMargin}`}>
                  {this.state.Confirmation.Description}
                </p>
              </div>
            )}
            <div className={this.stylesModal.footer}>
              <PrimaryButton
                iconProps={{
                  iconName: "SkypeCircleCheck",
                  styles: { root: this.stylesModal.buttonIcon },
                }}
                onClick={this.onConfirmation}
                text="Confirm"
                className={this.stylesModal.button}
                styles={{ root: this.stylesModal.buttonContent }}
              />
              <DefaultButton
                iconProps={{
                  iconName: "ErrorBadge",
                  styles: { root: this.stylesModal.buttonIcon },
                }}
                onClick={() =>
                  this.setState({
                    hideCnfirmationDialog: true,
                  })
                }
                text="Cancel"
                className={this.stylesModal.button}
                styles={{ root: this.stylesModal.buttonContent }}
              />
            </div>
          </>
        </Modal>
        <Modal
          isOpen={!this.state.hideSuccussDialog}
          onDismiss={() =>
            this.setState({
              hideSuccussDialog: true,
            })
          }
          isBlocking={true}
          containerClassName={this.stylesModal.modal}
        >
          <>
            <div className={styles.header}>
              <div style={{ display: "flex", alignItems: "center" }}>
                <IconButton iconProps={{ iconName: "Info" }} />
                <h4 className={this.stylesModal.headerTitle}>Alert</h4>
              </div>
              <IconButton
                iconProps={{ iconName: "Cancel" }}
                onClick={() => {
                  const pageURL: string = this.props.homePageUrl;
                  window.location.href = `${pageURL}`;
                  this.setState({
                    hideSuccussDialog: true,
                  });
                }}
              />
            </div>
            <div className={styles.body}>
              <p>{this.state.SuccussMsg}</p>
            </div>
            <div className={styles.footer}>
              <PrimaryButton
                className={styles.button}
                iconProps={{ iconName: "ReplyMirrored" }}
                onClick={() => {
                  const pageURL: string = this.props.homePageUrl;
                  window.location.href = `${pageURL}`;
                  this.setState({
                    hideSuccussDialog: true,
                  });
                }}
                text="Ok"
              />
            </div>
          </>
        </Modal>

        <Modal
          isOpen={!this.state.hideWarningDialog}
          onDismiss={() =>
            this.setState({
              hideWarningDialog: true,
            })
          }
          isBlocking={true}
          containerClassName={this.stylesModal.modal}
        >
          <>
            <div className={styles.header}>
            <div style={{ display: "flex", alignItems: "center" }}>
                <IconButton iconProps={{ iconName: "Info" }} />
                <h4 className={this.stylesModal.headerTitle}>Alert</h4>
              </div>
              <IconButton
                iconProps={{ iconName: "Cancel" }}
                onClick={() =>
                  this.setState({
                    hideWarningDialog: true,
                  })
                }
              />
            </div>
            <div className={styles.body}>
              <p>Please fill in comments then click on return</p>
            </div>
            <div className={styles.footer}>
              <PrimaryButton
                className={styles.button}
                iconProps={{ iconName: "ReplyMirrored" }}
                onClick={() =>
                  this.setState({
                    hideWarningDialog: true,
                  })
                }
                text="Ok"
              />
            </div>
          </>
        </Modal>
        <Modal
          isOpen={!this.state.hideApprovalDialog}
          onDismiss={() =>
            this.setState({
              hideApprovalDialog: true,
            })
          }
          isBlocking={true}
          containerClassName={this.stylesModal.modal}
        >
          <>
            <div className={styles.header}>
            <div style={{ display: "flex", alignItems: "center" }}>
                <IconButton iconProps={{ iconName: "Info" }} />
                <h4 className={this.stylesModal.headerTitle}>Alert</h4>
              </div>
              <IconButton
                iconProps={{ iconName: "Cancel" }}
                onClick={() => {
                  const pageURL: string = this.props.homePageUrl;
                  window.location.href = `${pageURL}`;
                  this.setState({
                    hideApprovalDialog: true,
                  });
                }
                }
              />
            </div>
            <div className={styles.body}>
              <p>This Committee meeting has been returned</p>
            </div>
            <div className={styles.footer}>
              <PrimaryButton
                className={styles.button}
                iconProps={{ iconName: "ReplyMirrored" }}
                onClick={() => {
                  const pageURL: string = this.props.homePageUrl;
                  window.location.href = `${pageURL}`;
                  this.setState({
                    hideApprovalDialog: true,
                  });
                }}
                text="Ok"
              />
            </div>
          </>
        </Modal>

          {/* Loading Section */}

          {this.state.isLoading && (
              <div>
                <Modal
                  isOpen={this.state.isLoading}
                  containerClassName={styles.spinnerModalTranparency}
                  styles={{
                    main: {
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      background: "transparent", // Removes background color
                      boxShadow: "none", // Removes box shadow
                    },
                  }}
                >
                  <div className="spinner">
                    <Spinner
                      label="still loading..."
                      ariaLive="assertive"
                      size={SpinnerSize.large}
                    />
                  </div>
                </Modal>
              </div>
            )}


             {/* duplicate validation */}
            
                    <Modal
                                      isOpen={this.state.hideParellelActionAlertDialog}
                                      onDismiss={() => {
                                        console.log("close triggered");
                                        this.setState((prevState) => ({
                                          hideParellelActionAlertDialog: !prevState.hideParellelActionAlertDialog,
                                        }));
                                        
                                      }}
                                      isBlocking={true}
                                      containerClassName={Cutsomstyles.modal}
                                    >
                                      <div className={Cutsomstyles.header}>
                                        <div style={{ display: "flex", alignItems: "center" }}>
                                          <IconButton iconProps={{ iconName: "Info" }} />
                                          <h4 className={Cutsomstyles.headerTitle}>Alert</h4>
                                        </div>
                                        <IconButton
                                          iconProps={{ iconName: "Cancel" }}
                                          onClick={() => {
                                            console.log("close triggered");
                                            window.location.reload();
                                            this.setState({ hideParellelActionAlertDialog: false });
                                          }}
                                        />
                                      </div>
                                      <div className={Cutsomstyles.body}>
                                        <p>{this.state.parellelActionAlertMsg}</p>
                                      </div>
                                      <div className={Cutsomstyles.footer}>
                                        <PrimaryButton
                                          className={Cutsomstyles.button}
                                          iconProps={{ iconName: "ReplyMirrored" }}
                                          onClick={() =>{
                                            this.setState({ hideParellelActionAlertDialog: false })
                                            window.location.reload();
            
                                          }}
                                          text="OK"
                                        />
                                      </div>
                                    </Modal>

            

            {/* PassCode Section */}

        <form>
              <PasscodeModal
            createPasscodeUrl={this.props.passCodeUrl}
            isOpen={this.state.isPasscodeModalOpen}
            onClose={() => this.setState({
              isPasscodeModalOpen: false,
              isPasscodeValidated: false,
            })}
            sp={this.props.sp}
            user={this.props.context.pageContext.user}
            onSuccess={this.handlePasscodeSuccess}              />
            </form>
      </div>
    );
  }
}
