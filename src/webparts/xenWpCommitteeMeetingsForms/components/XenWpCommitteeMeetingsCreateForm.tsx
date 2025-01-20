/* eslint-disable prefer-const */
/* eslint-disable react/self-closing-comp */
/* eslint-disable max-lines */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./XenWpCommitteeMeetingsForms.module.scss";
import "./CustomStyles/custom.css";
import type { IXenWpCommitteeMeetingsFormsProps } from "./IXenWpCommitteeMeetingsFormsProps";
import {
  DatePicker,
  DefaultButton,
  defaultDatePickerStrings,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  Icon,
  IconButton,
  IDetailsFooterProps,
  IDropdownOption,
  Link,
  mergeStyleSets,
  Modal,
  PrimaryButton,
  SelectionMode,
  Spinner,
  SpinnerSize,
  TextField,
} from "@fluentui/react";
import { format } from "date-fns";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import * as he from 'he';
// import DOMPurify from "dompurify";
// import { decode, encode } from 'he';
// import { encode, decode } from "html-entities";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/files";
import "@pnp/sp/profiles";
import "@pnp/sp/site-groups";
import DateTime from "./dateComponent";

interface CommtteeMeetingsStateProps {
  committename: any;
  committeeNameFeildValue: any;
  charimanFeildValue: any;
  convernorFeildValue: any;
  convernorData: any;
  charimanData: any;
  selectedCommitteeMembers: any;
  selectedCommitteeGuestMembers: any;
  selectedCommitteeNoteRecords: any;
  committeeMembersData: any;
  committeeGuestMembersData: any;
  committeeNoteRecordsData: any;
  committeeMemberskey: any;
  committeeGuestMemberskey: any;
  committeeNoteRecordskey: any;
  isModalOpen: any;
  isModalMOMOpen: any;
  modalMessage: any;
  meetingId: any;
  meetingDate: any;
  meetingSubject: any;
  meetingMode: any;
  meetingLink: any;
  statusNumber: any;
  auditTrail: any;
  _IsAuthor:boolean;

  committeeNoteRecordDropDownData: any;
  committeeNoteRecordSelectedValue: any;
  CommitteeMeetingMemberCommentsDT: any;
  pdfLink: any;
  committeNoteRecordsDropDownDataWithAllProperties: any;
  selectedMOMNoteRecord: any;
  invalidFields: any;
  MeetingStatus: any;

  isMomDraftDialogOpen: boolean;
  dialogType: any;
  draftResolutionFieldValue: any;
  btnType: any;
  departmentAlias: any;
  isWarningCommitteeName: boolean;
  isWarningConvenor: boolean;
  isWarningChairman: boolean;
  isWarningMeetingDate: boolean;
  isWarningMeetingSubject: boolean;
  isWarningMeetingMode: boolean;
  isWarningMeetingLink: boolean;

  isSmallScreen: boolean;
  isLoading: any;
  hideParellelActionAlertDialog:any;
  parellelActionAlertMsg:any;


}

const getIdFromUrl = (): any => {
  const params = new URLSearchParams(window.location.search);
  const Id = params.get("itemId");

  return Id;
};

export interface IFileDetails {
  name?: string;
  content?: File;
  index?: number;
  fileUrl?: string;
  ServerRelativeUrl?: string;
  isExists?: boolean;
  Modified?: string;
  isSelected?: boolean;
}

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

export default class XenWpCommitteeMeetingsCreateForm extends React.Component<
  IXenWpCommitteeMeetingsFormsProps,
  CommtteeMeetingsStateProps
> {
  private _peopplePicker: IPeoplePickerContext;
  private _libraryName: any;
  private _listName: any;
  private _committeNameList: any;

  private _itemId: number = Number(getIdFromUrl());
  private _absUrl: any = this.props.context.pageContext.web.serverRelativeUrl;
  private _homePageUrl: any = this.props.homePageUrl;


  private _currentUserEmail =( this.props.context.pageContext.user.email).toLocaleLowerCase();

  constructor(props: any) {
    super(props);
    this.state = {
      committename: [],
      committeeNameFeildValue: "",
      charimanFeildValue: "",
      convernorFeildValue: "",
      convernorData: {},
      charimanData: {},
      selectedCommitteeMembers: [],
      selectedCommitteeGuestMembers: [],
      selectedCommitteeNoteRecords: [],
      committeeMembersData: [],
      committeeGuestMembersData: [],
      committeeNoteRecordsData: [],
      committeeMemberskey: [],
      committeeGuestMemberskey: [],
      committeeNoteRecordskey: [],
      CommitteeMeetingMemberCommentsDT: [],
      isModalOpen: false,
      isModalMOMOpen: false,
      modalMessage: "",
      meetingId: "To be generated",
      meetingDate: null,
      meetingSubject: "",
      meetingMode: "",
      meetingLink: "",
      statusNumber: "",
      MeetingStatus: "",
      auditTrail: [],
      committeeNoteRecordDropDownData: [],
      committeeNoteRecordSelectedValue: "",
      pdfLink: "",
      committeNoteRecordsDropDownDataWithAllProperties: [],
      dialogType: "",
      isMomDraftDialogOpen: false,
      draftResolutionFieldValue: "",
      selectedMOMNoteRecord: "",
      btnType: "",
      departmentAlias: "",
      invalidFields: [],

      isWarningCommitteeName: false,
      isWarningConvenor: false,
      isWarningChairman: false,
      isWarningMeetingDate: false,
      isWarningMeetingSubject: false,
      isWarningMeetingMode: false,
      isWarningMeetingLink: false,
      isLoading: true,
      _IsAuthor:false,

      isSmallScreen: window.innerWidth < 768,

      hideParellelActionAlertDialog:false,
      parellelActionAlertMsg:''

    };

    this._peopplePicker = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    const committeName = this.props.committeeMeetingNameList;
    this._committeNameList = committeName?.title;
    const libraryTilte = this.props.libraryId;
    this._libraryName = libraryTilte?.title;
    
    const listName = this.props.listName;
    this._listName = listName?.title;
    this.getfield();
    this._fetchDepartmentAlias();

    window.addEventListener("resize", this.handleResize);
    
    this._itemId && this._getItemData(this._itemId);
    
  }

  public componentDidMount() {
    window.addEventListener("resize", this.handleResize);

    setTimeout(() => {
      this.setState({ isLoading: false });
    }, 3000);
  }

  // public componentWillUnmount() {
  //   window.removeEventListener("resize", this.handleResize);
  // }

  private handleResize = () => {
    console.log(window.innerWidth < 768,"is Small Screen")
    this.setState({ isSmallScreen: window.innerWidth < 768 });
  };

  private _getECommitteRequestBasedOnSelectedItem = async (
    committeeName: any
  ): Promise<any> => {

    const currentUserId=(await this.props.sp?.web.currentUser())?.Id;
    const filterQury=  `AuthorId eq ${currentUserId} and CommitteeType eq 'Committee' and CommitteeName eq '${committeeName}' and isMapped eq 'false' and StatusNumber eq '9000' `;
    try {
      const fieldDetails = await this.props.sp.web.lists
        .getByTitle("EcommiteeRequests")
        .items.select(
          "*",
          "CommitteeType",
          "Title",
          "Department",
          "CommitteeName"
        )
        .filter(
          filterQury
        )();
       

      const dropDownDataListing = fieldDetails
        .filter((each: any) => {
          return each.isMapped === false;
        })
        .map((each: any) => {
          return {
            ...each,
            key: each.Title,
            text: each.Title,
            id: each.Title,
            ...each
          };
        });

      const committeNoteRecordsDropDownDataWithAllProperties =
        await Promise.all(
          fieldDetails.map(async (each: any) => {
            const link = await this._getItemDocumentsData(each.Title);

            return {
              ...each,
              key: each.Title,
              text: each.Title,
              id: each.Id,
              noteTitle: each.Title,
              committeeName: each.CommitteeName,
              department: each.Department,
              noteLink: link[0],
              link: link[1],
              mom: "",
            };
          })
        );

      this.setState({
        committeeNoteRecordDropDownData: dropDownDataListing,
        committeNoteRecordsDropDownDataWithAllProperties:
          committeNoteRecordsDropDownDataWithAllProperties,
      });
    } catch (error) {
      console.error("Error fetching field details: ", error);
    }
  };

  private _getFileObj = (data: any): any => {
    const tenantUrl = window.location.protocol + "//" + window.location.host;
    const result = data.TimeCreated;
    const filesObj = {
      name: data.Name,
      content: data,
      index: 0,
      fileUrl: tenantUrl + data.ServerRelativeUrl,
      ServerRelativeUrl: "",
      isExists: true,
      Modified: "",
      isSelected: false,
      size: parseInt(data.Length),
      type: `application/${data.Name.split(".")[1]}`,
      modifiedBy: data.Author.Title,
      createData: result,
    };
    return filesObj;
  };

  private _getItemDocumentsData = async (tilte: any): Promise<any> => {
    const folderName = tilte.replace(/\//g, "-");
    const url = `${this._absUrl}/${this._libraryName}/${folderName}`;
    try {
      const folderItemsPdf = await this.props.sp.web
        .getFolderByServerRelativePath(`${url}/Pdf`)
        .files.select("*")
        .expand("Author", "Editor")();
      let pdfLink;
      
      folderItemsPdf.forEach((values: any) => {
     
        pdfLink = this._getFileObj(values).fileUrl;
      });
      return [pdfLink, `${this._libraryName}/${folderName}`];
    } catch (err){
      return err
    }
  };

  private _fetchDepartmentAlias = async (): Promise<void> => {
    try {
      const profile = await this.props.sp.profiles.myProperties();
      profile.UserProfileProperties.filter(async (element: any) => {
        if (element.Key === "Department") {
          const items: any[] = await this.props.sp.web.lists
            .getByTitle("Departments")
            .items.filter(`Department eq '${element.Value}'`)
            .select(
              "*",
              "Department",
              "DepartmentAlias",
              "Admin/EMail",
              "Admin/Title"
            ) 
            .expand("Admin")();

          this.setState({
            departmentAlias: items[0].DepartmentAlias, 
          });
        }
      });
    } catch (error) {
    
      return error
    }
  };

  private _getTitle = (id: any): any => {
    const currentyear = new Date().getFullYear();
    const nextYear = (currentyear + 1).toString().slice(-2);

    let meetingID = "";
   

    this.setState((prevState) => ({
      meetingId: `${prevState.departmentAlias}/${currentyear}-${nextYear}/${id}`,
    }));
    

    meetingID = `${this.state.departmentAlias}/${currentyear}-${nextYear}/${id}`;

    return meetingID;
  };


  private _filterChairmanDataFromCommitteeMembersDTO = (data:any):any=>{
    const committeeMemberDTOFromList = JSON.parse(data)
    const committeeMemberDTOWithOutChairmansData = committeeMemberDTOFromList.filter((each:any)=>each.isChairman === false)
   
    return committeeMemberDTOWithOutChairmansData

  }

  private _getItemDataFromList = async (id:any)=>{
    const item: any = await this.props.sp.web.lists
    .getByTitle(this._listName)
    .items.getById(id)
    .select(
      "*",
      "Chairman",
      "Chairman/Title",
      "Chairman/EMail",

      "CurrentApprover",

      "CurrentApprover/Title",

      "CurrentApprover/EMail"
    )
    .expand("Chairman", "CurrentApprover")();

   

    return item

  }


  // private escapeHtml = (input: string): string => {
  //   return input.replace(/<p>(.*?)<\/p>/g, (match, content) => {
  //     const div = document.createElement('div');
  //     div.innerText = content; // Escape only the content within the <p> tags
  //     return `<p>${div.innerHTML}</p>`; // Return the <p> tags with escaped content
  //   });
  // };


  // private _getEncodedCommitteeMeetingNoteDTO = (input:any)=>{
  //   return input.map((each:any)=>({...each,mom:this.escapeHtml(each.mom)}))
  // }

  private _getItemData = async (id: any) => {
    const item: any = await this._getItemDataFromList(id)
    const _chairmanProfile = await this.props.sp.web.getUserById(item.ChairmanId)();
    const _chairmanDesignation = await this._getUserProperties( _chairmanProfile?.LoginName);
    

    this.setState({
      _IsAuthor:item.AuthorId ===  (await this.props.sp?.web.currentUser())?.Id,
      meetingId: item.Title,
      meetingDate: new Date(item.MeetingDate),
      meetingSubject: item.MeetingSubject,
      meetingMode: item.MeetingMode,
      meetingLink: item.MeetingLink,
      MeetingStatus: item.MeetingStatus,
      statusNumber: item.StatusNumber,
      committeeNameFeildValue: item.CommitteeName,
      committeeMembersData:this._filterChairmanDataFromCommitteeMembersDTO(item.CommitteeMeetingMembersDTO),
      committeeGuestMembersData: JSON.parse(
        item.CommitteeMeetingGuestMembersDTO
      ),
      committeeNoteRecordsData: JSON.parse(item.CommitteeMeetingNoteDTO),
      CommitteeMeetingMemberCommentsDT:
        item.CommitteeMeetingMemberCommentsDT === null
          ? []
          : JSON.parse(item.CommitteeMeetingMemberCommentsDT),
      convernorFeildValue: item.Department,
      charimanFeildValue: item.Chairman.Title,
      auditTrail: JSON.parse(item.AuditTrail),
      isLoading: false,
      convernorData: JSON.parse(item.ConvenerDTO),

      charimanData: { ...item.Chairman, chairmanId: item.ChairmanId ,designation:_chairmanDesignation[0]},
    });

    return item;
  };

  private getfield = async () => {
    try {
      
      
     

      //  const convernorDetails = await this.props.sp.web.lists
      // .getByTitle(this._committeNameList)
      // .items.filter(`Convener/EMail eq '${this._currentUserEmail}'`)
      // .select(
        

      //   "Convener/EMail"
      // )
      // .expand("Convener")();
      // console.log(convernorDetails)

      const convernorDetails = await this.props.sp.web.lists
      .getByTitle(this._committeNameList)
      .items.filter(`Convener/EMail eq '${this._currentUserEmail}'`)
      .select(
        
       
        "Chairman/EMail",
       
      )
      .expand("Chairman")();
     

      const isConvernorExists = convernorDetails.length > 0
     

      if (isConvernorExists){

        const fieldDetails = await this.props.sp.web.lists
        .getByTitle(this._committeNameList)
        .fields.filter("Hidden eq false and ReadOnlyField eq false")();


    

     

      const filtering = fieldDetails.map(
        (_x: { TypeDisplayName: string; InternalName: any; Choices: any }) => {
          if (_x.TypeDisplayName === "Choice") {
            return [_x.InternalName, _x.Choices];
          }
        }
      );

      const finalList = filtering?.filter((each: any) => {
        if (typeof each !== "undefined") {
          return each;
        }
      });

      finalList?.map((each: string | any[] | undefined) => {
        if (
          each !== undefined &&
          Array.isArray(each) &&
          each.length > 1 &&
          Array.isArray(each[1])
        ) {
          if (each[0] === "CommitteeNames") {
            const committenameArray = each[1].map((item, index) => {
              return { key: item, text: item };
            });

            this.setState({ committename: committenameArray });
          }
        }
      });
        
      }else{
        this.setState({ committename: [] });

      }

      


      

    


    } catch (error) {
      console.error("Error fetching field details: ", error);
    }
  };


  // private getfield = async () => {
  //   try {
      

  //     const fieldDetails = await this.props.sp.web.lists
  //       .getByTitle("CommitteeMeetingApprovers")
  //       .fields.filter("Hidden eq false and ReadOnlyField eq false")();

     

  //     const filtering = fieldDetails.map(
  //       (_x: { TypeDisplayName: string; InternalName: any; Choices: any }) => {
  //         if (_x.TypeDisplayName === "Choice") {
  //           return [_x.InternalName, _x.Choices];
  //         }
  //       }
  //     );

  //     const finalList = filtering?.filter((each: any) => {
  //       if (typeof each !== "undefined") {
  //         return each;
  //       }
  //     });

  //     finalList?.map((each: string | any[] | undefined) => {
  //       if (
  //         each !== undefined &&
  //         Array.isArray(each) &&
  //         each.length > 1 &&
  //         Array.isArray(each[1])
  //       ) {
  //         if (each[0] === "CommitteeName") {
  //           const committenameArray = each[1].map((item, index) => {
  //             return { key: item, text: item };
  //           });

  //           this.setState({ committename: committenameArray });
  //         }
  //       }
  //     });


  //   } catch (error) {
  //     console.error("Error fetching field details: ", error);
  //   }
  // };
  private handleDeleteCommitteeMemberData = (index: number): void => {



    this.setState((prevState) => ({
      committeeMembersData: prevState.committeeMembersData.filter(
        (_: any, i: number) => i !== index
      ),
    }));
  };

  private _columnsCommitteeMembersSNOObjOnRender = (_item: any, _index?: number):any => (
    <span>{(_index !== undefined ? _index : 0) + 1}</span>
  )


  private _columnsCommitteeMembersActionObjOnRender =(_item: any, index?: number) => (
    <IconButton
      iconProps={{ iconName: "Delete" }}
      title="Delete"
      ariaLabel="Delete"
      onClick={() => this.handleDeleteCommitteeMemberData(index!)}
      styles={{
        root: { paddingBottom: "16px", background: "transparent" },
      }}
    />
  )

  private columnsCommitteeMembers: IColumn[] = [
    {
      key: "sNo",
      name: "S.No",
      fieldName: "sNo",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      onRender: this._columnsCommitteeMembersSNOObjOnRender
    },
    {
      key: "memberEmailName",
      name: "Members",
      fieldName: "memberEmailName",
      minWidth: 150,
      maxWidth: 230,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: this._columnsCommitteeMembersActionObjOnRender
    },
  ];

  private handleDeleteGuestMember = (index: number): void => {
    this.setState((prevState) => ({
      committeeGuestMembersData: prevState.committeeGuestMembersData.filter(
        (_: any, i: number) => i !== index
      ),
    }));
  };

  private columnsCommitteeGuestMembers: IColumn[] = [
    {
      key: "sNo",
      name: "S.No",
      fieldName: "sNo",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      onRender: (_item: any, _index?: number) => (
        <span>{(_index !== undefined ? _index : 0) + 1}</span>
      ),
    },
    {
      key: "memberEmailName", 
      name: "Guest Members",
      fieldName: "memberEmailName",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (_item: any, index?: number) => (
        <IconButton
          iconProps={{ iconName: "Delete" }}
          title="Delete"
          ariaLabel="Delete"
          onClick={() => this.handleDeleteGuestMember(index!)}
          styles={{
            root: { paddingBottom: "16px", background: "transparent" },
          }}
        />
      ),
    },
  ];

  private _deleteRecord = (id: number) => {

    this.setState((prevState)=>{
      const updatedData = prevState.committeeNoteRecordsData.filter(
        (record: { id: number }) => record.id !== id
      );
      return {committeeNoteRecordsData: updatedData }
      
    })
  };

  private columnsCommitteeNoteRecords: IColumn[] = [
    {
      key: "noteTitle",
      name: "Note Title",
      fieldName: "noteTitle",
      minWidth: 120,
      maxWidth: 180,
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
      key: "link",
      name: "Note Link",
      fieldName: "link",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
      onRender: (item) => {

        return (
          <Link
            href={item.noteLink}
            download={item.noteTitle}
          >
            {item.link}
          </Link>
        );
      },
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item) =>
        this.state.statusNumber !== "1000" &&
        this.state.statusNumber !== "2000" &&
        this.state.statusNumber !== "" ? (
          <IconButton
            iconProps={{ iconName: "Edit" }}
            title="Edit"
            ariaLabel="Edit"
            onClick={() => {
              
              this.setState({
                dialogType: "mom",
                isModalMOMOpen: true,
                selectedMOMNoteRecord: item.key,
                draftResolutionFieldValue: item.meetingMinutesForRichTextComponent,
              });
            }}
          />
        ) : (
          <IconButton

            iconProps={{ iconName: "Delete" }}
            title="Delete"
            ariaLabel="Delete"
            onClick={() => this._deleteRecord(item.id)}
            styles={{
              root: { paddingBottom: "16px", background: "transparent" },
            }}
          />
        ),
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

  private _clearCommitteeMembersPeoplePicker = () => {
    this.setState(   (prevState) => ({
      selectedCommitteeMembers: [],
      committeeMemberskey: prevState.committeeMemberskey + 1,
    }))
 
  };

  private _clearCommitteeGuestMembersPeoplePicker = () => {
    this.setState(   (prevState) => ({
      selectedCommitteeGuestMembers: [],
      committeeGuestMemberskey: prevState.committeeGuestMemberskey + 1,
    }))
 
  };

  private _clearCommitteeNoteRecordsPeoplePicker = () => {
 
    this.setState(   (prevState) => ({
      selectedCommitteeNoteRecords: [],
      committeeNoteRecordskey: prevState.committeeNoteRecordskey + 1,
    }))
  };

  //Committee Members Validation

  private checkSelectedCommitteMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember =
    (): boolean => {
      const committeeMemberEmail = this.state.committeeMembersData.map(
        (each: any) => each.email || each.memberEmail
      );

      const committeeGuestMembersEmail =
        this.state.committeeGuestMembersData.map(
          (each: any) => each.email || each.memberEmail
        );

      const selectedCommitteeMembers = this.state.selectedCommitteeMembers[0];
      const selectedCommitteeMembersEmail =
        selectedCommitteeMembers.email ||
        selectedCommitteeMembers.secondaryText;
      const selectedMemberIsAChairman =
        selectedCommitteeMembersEmail === this.state.charimanData.EMail;

      const selectedMemberIsACon =
        selectedCommitteeMembersEmail === this.state.convernorData.EMail;

     
      const iscommitteeMemberOrcommitteeGuestMembers =
        committeeGuestMembersEmail.includes(selectedCommitteeMembersEmail) ||
        committeeMemberEmail.includes(selectedCommitteeMembersEmail);

      const isCurrentUserCommitteMember =
        this._currentUserEmail === selectedCommitteeMembersEmail;

      return (
        iscommitteeMemberOrcommitteeGuestMembers ||
        isCurrentUserCommitteMember ||
        selectedMemberIsAChairman ||
        selectedMemberIsACon
      );
    };

/*   //Guest Members Validation */

  private checkSelectedGuestMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember =
    (): boolean => {
      const committeeMemberEmail = this.state.committeeMembersData.map(
        (each: any) => each.email || each.memberEmail
      );

      const committeeGuestMembersEmail =
        this.state.committeeGuestMembersData.map(
          (each: any) => each.email || each.memberEmail
        );

      const selectedGuestMembers = this.state.selectedCommitteeGuestMembers[0];
      const selectedGuestMembersEmail =
        selectedGuestMembers.email || selectedGuestMembers.secondaryText;

      const selectedMemberIsAChairman =
        selectedGuestMembersEmail === this.state.charimanData.EMail;
  
      const iscommitteeMemberOrcommitteeGuestMembers =
        committeeGuestMembersEmail.includes(selectedGuestMembersEmail) ||
        committeeMemberEmail.includes(selectedGuestMembersEmail);

      const isCurrentUserGuestMember =
        this._currentUserEmail === selectedGuestMembersEmail;

      const selectedGuestMemberIsACon =
        selectedGuestMembersEmail === this.state.convernorData.EMail;

      return (
        iscommitteeMemberOrcommitteeGuestMembers ||
        isCurrentUserGuestMember ||
        selectedMemberIsAChairman ||
        selectedGuestMemberIsACon
      );
    };

  private handleOnAdd = async (event: any, type: string): Promise<void> => {
    if (this.state.committeeNameFeildValue === "") {
      this.setState({
        isModalOpen: true,
        modalMessage: "Please select Committee Name and click Add.",
      });
      this._clearCommitteeMembersPeoplePicker();
      return;
    }

    if (type === "committeeMembers") {
      if (
        this.checkSelectedCommitteMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember()
      ) {
        this.setState({
          isModalOpen: true,
          modalMessage:
            "The selected member cannont be same as existing Member/Convenor/Chairman.",
        });
        this._clearCommitteeMembersPeoplePicker();
        this._clearCommitteeGuestMembersPeoplePicker();
        return;
      }

      this.setState(
        (prevState) => ({
          committeeMembersData: [
            ...prevState.committeeMembersData,
            ...prevState.selectedCommitteeMembers,
          ],
        }))
      this._clearCommitteeMembersPeoplePicker();
    } else if (type === "committeeGuestMembers") {
      if (
        this.checkSelectedGuestMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember()
      ) {
        this.setState({
          isModalOpen: true,
          modalMessage:
            "The selected member cannont be same as existing Member/Convenor/Chairman.",
        });
        this._clearCommitteeMembersPeoplePicker();
        this._clearCommitteeGuestMembersPeoplePicker();
        return;
      }
  
      this.setState(
        (prevState) => ({
          committeeGuestMembersData: [
            ...prevState.committeeGuestMembersData,
            ...prevState.selectedCommitteeGuestMembers,
          ],
        }))
      this._clearCommitteeGuestMembersPeoplePicker();
    } else if (type === "committeeNoteRecords") {
 
      this.setState(
        (prevState) => ({
          committeeNoteRecordsData: [
            ...prevState.committeeNoteRecordsData,
            ...prevState.selectedCommitteeNoteRecords,
          ],
        }))
      this._clearCommitteeNoteRecordsPeoplePicker();
    }
  };

  private _getUserProperties = async (loginName: any): Promise<any> => {
    let designation ;
    let email ;
    const profile = await this.props.sp.profiles.getPropertiesFor(loginName);
    designation = profile.Title;
    email = profile.Email;
    const props: any = {};
    profile.UserProfileProperties.forEach(
      (prop: { Key: string | number; Value: any }) => {
        props[prop.Key] = prop.Value;
      }
    );

    profile.userProperties = props;
    return [designation, email];
  };

  private _getPeoplePickerItemsCommitteeMembers = async (
    items: any[]
  ): Promise<any> => {
    const dataRec = await this._getUserProperties(items[0].loginName);
    if (typeof dataRec[0]?.toString() === "undefined") {
      const newItemsDataNA = items.map((obj: any) => {
        return {
          ...obj,
          optionalText: "N/A",

          email: obj.secondaryText,
        };
      });
      this.setState({
        selectedCommitteeMembers: [
          {
            memberEmailName: newItemsDataNA[0].text,
            srNo: newItemsDataNA[0].srNo,
            designation: newItemsDataNA[0].optionalText,
            email: newItemsDataNA[0].email,
            memberEmail: newItemsDataNA[0].email,
            userId: newItemsDataNA[0].id,
            isChairman:false
          },
        ],
      });
    } else {
      const newItemsData = items.map((obj: any) => {
        return {
          ...obj,
          optionalText: dataRec[0],

          email: dataRec[1],
          srNo: dataRec[1].split("@")[0],
        };
      });
      this.setState({
        selectedCommitteeMembers: [
          {
            memberEmailName: newItemsData[0].text,
            srNo: newItemsData[0].srNo,
            designation: newItemsData[0].optionalText,
            email: newItemsData[0].email,
            memberEmail: newItemsData[0].email,
            userId: newItemsData[0].id,
          },
        ],
      });
    }
  };

  private _getPeoplePickerItemsCommitteeGuestMembers = async (
    items: any[]
  ): Promise<any> => {
    const dataRec = await this._getUserProperties(items[0].loginName);
    if (typeof dataRec[0]?.toString() === "undefined") {
      const newItemsDataNA = items.map((obj: any) => {
        return {
          ...obj,
          optionalText: "N/A",

          email: obj.secondaryText,
        };
      });
      this.setState({
        selectedCommitteeGuestMembers: [
          {
            memberEmailName: newItemsDataNA[0].text,
            srNo: newItemsDataNA[0].srNo,
            designation: newItemsDataNA[0].optionalText,
            email: newItemsDataNA[0].email,
            memberEmail: newItemsDataNA[0].email,
            userId: newItemsDataNA[0].id,
          },
        ],
      });
    } else {
      const newItemsData = items.map((obj: any) => {
        return {
          ...obj,
          optionalText: dataRec[0],

          email: dataRec[1],
          srNo: dataRec[1].split("@")[0],
        };
      });
      this.setState({
        selectedCommitteeGuestMembers: [
          {
            memberEmailName: newItemsData[0].text,
            srNo: newItemsData[0].srNo,
            designation: newItemsData[0].optionalText,
            email: newItemsData[0].email,
            memberEmail: newItemsData[0].email,
            userId: newItemsData[0].id,
          },
        ],
      });
    }
  };

  private onRenderCaretDowncommitteeNameFeildValue = (): JSX.Element => {
    return this.state.committeeNameFeildValue ? (
      <Icon
        iconName="Cancel"
        onClick={() => {
          this.setState({ committeeNameFeildValue: "" });
        }}
      />
    ) : (
      <Icon
        iconName="ChevronDown"
        onClick={() => {
          this.setState({ committeeNameFeildValue: "" });
        }}
      />
    );
  };

  private handleCommittename = async (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> => {
    const value = option ? option.text : "";

    const fieldDetails = await this.props.sp.web.lists
    .getByTitle(this._committeNameList)
    .items.filter(`CommitteeNames eq '${value}'`).select(
      "*",
      "Convener",
      "Convener/Title",
      "Chairman",
      "Chairman/Title",
      "Chairman/EMail",
      "Convener/EMail"
    )
    .expand("Convener", "Chairman")();
    

    const each = fieldDetails[0]


   

  

    if (fieldDetails.length > 0){
      const convenorDepartment = await this.getUserDepartmentByEmail(
        each.ConvenerId
      );
     

      const _chairmanProfile = await this.props.sp.web.getUserById(each.ChairmanId)();

     
      const _chairmanDesignation = await this._getUserProperties( _chairmanProfile?.LoginName);
     

    

      this.setState({
        charimanFeildValue: each.Chairman.Title,
        convernorData: {
          ...each.Convener,
          convenerId: each.ConvenerId,
          department: convenorDepartment,
        },
        charimanData: { ...each.Chairman, chairmanId: each.ChairmanId,designation:_chairmanDesignation[0]},
      });

    }else{
      this.setState({
        charimanFeildValue: '',
        convernorData: '',
        charimanData: '',
      });

    }
  

  
    this._getECommitteRequestBasedOnSelectedItem(value);

    this.setState({
      committeeNameFeildValue: value,
      committeeGuestMembersData: [],
      committeeMembersData: [],
      committeeNoteRecordsData: [],
    });
  };

  private onDateChange = (date: any): void => {
    this.setState({ meetingDate: date }); // Update the state with the selected date, or null if the user clears the selection
  };

  private _closeModal = (): void => {
    this.setState({ isModalOpen: false, isModalMOMOpen: false });
    this.setState({ dialogType: "" });
  };

  private _getCommitteeMeetingMembersDTO = (data: any): any => {
    const makeCommitteeMeetingMemberDTO = data.map((each: any, index: any) => {
      return {
        createdDate: new Date(),
        userId: each.userId,
        srNo: each.srNo,
        designation: each.designation,
        memberEmail: each.email || each.memberEmail,
        memberEmailName: each.memberEmailName,
        status: "Pending",
        statusNumber: "",
        isChairman:false
      };
    });

    const _chairmanIncluded = {
      createdDate:this._formatDateTime(new Date()) ,
      userId: 0,
      srNo:this.state.charimanData.EMail.split("@")[0],
      designation:this.state.charimanData.designation,
      memberEmail:this.state.charimanData.EMail,
      memberEmailName:this.state.charimanData.Title,
      status: "",
      statusNumber: "",
      isChairman:true

    }

    makeCommitteeMeetingMemberDTO.push(_chairmanIncluded)
  

    return JSON.stringify(makeCommitteeMeetingMemberDTO);
  };

  private _getCommitteeGuestMembersDTO = (data: any): any => {
    const makeCommitteeMeetingMemberDTO = data.map((each: any, index: any) => {
      return {
        createdDate: new Date(),
        userId: each.userId,
        srNo: each.srNo,
        designation: each.designation,
        memberEmail: each.email || each.memberEmail,
        memberEmailName: each.memberEmailName,
        status: "Pending",
        statusNumber: "",
        isChairman:false
      };
    });

   

    return JSON.stringify(makeCommitteeMeetingMemberDTO);
  };

  private _getMemeberId = (data: any) => {
    const getId = data.map((each: any) => each.userId);
    return getId;
  };

  private _getCurrrentApproverId = (data: any): any => {


    return data[0].userId;
  };


  private eCommitteeMeetingsParamsObject = async (status:any,statusNumber:any,auditTrail:any):Promise<any> =>{
    const ecommitteObject: any = {
      MeetingDate: this.state.meetingDate,
      MeetingLink: this.state.meetingLink,

      MeetingMode: this.state.meetingMode,
      MeetingSubject: this.state.meetingSubject,
      CommitteeName: this.state.committeeNameFeildValue,
      ChairmanId: this.state.charimanData.chairmanId,
      FinalApproverId:
        this.state.committeeMembersData[
          this.state.committeeMembersData.length - 1
        ].userId,
      CommitteeMeetingMembersDTO: this._getCommitteeMeetingMembersDTO(
        this.state.committeeMembersData
      ),
      CommitteeMeetingGuestMembersDTO: this._getCommitteeGuestMembersDTO(
        this.state.committeeGuestMembersData
      ),
      CommitteeMeetingNoteDTO: JSON.stringify(
        this.state.committeeNoteRecordsData
      ),
 
      CurrentApproverId: (await this.props.sp?.web.currentUser())?.Id,
      CommitteeMeetingMembersId: this._getMemeberId(
        this.state.committeeMembersData
      ),
      CommitteeMeetingGuestsId: this._getMemeberId(
        this.state.committeeGuestMembersData
      ),
      MeetingStatus: status,
      ConvenerDTO: JSON.stringify(this.state.convernorData),
      Department: this.state.convernorFeildValue,
      StatusNumber: statusNumber,
      AuditTrail: JSON.stringify(auditTrail),
      startProcessing: true,
      PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
    };
     
    return ecommitteObject

  }

   private _formatDateTime = (date: string | number | Date) => {
      const formattedDate = format(new Date(date), "dd-MMM-yyyy");
      const formattedTime = format(new Date(date), "hh:mm a");
      return `${formattedDate} ${formattedTime}`;
    };

  private createEcommitteeMeetingObject = async (
    status: string,
    statusNumber: any
  ): Promise<any> => {


    

    const auditTrail = this.state.auditTrail;

    const auditLog = {
      actionBy: this.props.context.pageContext.user.displayName,

      action: `Committee Meeting ${status}`,

      actionDate: this. _formatDateTime(new Date()),
    };
    auditTrail.push(auditLog);

    const ecommitteObject: any = await this.eCommitteeMeetingsParamsObject(status,statusNumber,auditTrail)
    return ecommitteObject;
  };



  private _handleCreateMeeting = async (): Promise<void> => {
    this.setState({ isModalOpen: false, isLoading: true });

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.add(await this.createEcommitteeMeetingObject("Created", "1000"))
      .then(async (res: any) => {
        const title = this._getTitle(res.Id);

        await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(res.Id)
          .update({
            Title: title,
          });
      });


    this.setState({
      isLoading: false,
      isModalOpen: true,
      dialogType: "success",
    });
  };

  private _handlePulbicMeeting = async (): Promise<void> => {
    this.setState({ isModalOpen: false, isLoading: true });


    

    const auditTrail = {
      action: `Committee Meeting Published`,
      actionBy: this.props.userDisplayName,
      actionDate:this. _formatDateTime(new Date()),
    };

    this._itemId
      ? await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(this._itemId)
          .update({
            StatusNumber: "2000",
            MeetingStatus: "Published",
            FinalApproverId:
              this.state.committeeMembersData[
                this.state.committeeMembersData.length - 1
              ].userId,
            CommitteeMeetingMembersDTO: this._getCommitteeMeetingMembersDTO(
              this.state.committeeMembersData
            ),
            CommitteeMeetingGuestMembersDTO:
              this._getCommitteeGuestMembersDTO(
                this.state.committeeGuestMembersData
              ),
            CommitteeMeetingNoteDTO: JSON.stringify(
              this.state.committeeNoteRecordsData
            ),
       
            CurrentApproverId: (await this.props.sp?.web.currentUser())?.Id,
            CommitteeMeetingMembersId: this._getMemeberId(
              this.state.committeeMembersData
            ),
            CommitteeMeetingGuestsId: this._getMemeberId(
              this.state.committeeGuestMembersData
            ),

            AuditTrail: JSON.stringify([...this.state.auditTrail, auditTrail]),
            startProcessing: true,
            PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
          })
      : await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.add(
            await this.createEcommitteeMeetingObject("Published", "2000")
          )
          .then(async (res: any) => {
            const title = this._getTitle(res.Id);

            await this.props.sp.web.lists
              .getByTitle(this._listName)
              .items.getById(res.Id)
              .update({
                Title: title,
                MeetingNumber:title

              });
          });
    this.setState({
      isLoading: false,
      isModalOpen: true,
      dialogType: "success",
    });
  };

  private _handleReturnBack = async (): Promise<void> => {
    this.setState({ isModalOpen: false, isLoading: true });

    
    const itemFromList = await this._getItemDataFromList(this._itemId);

    // const _CommitteeMemberDTO = JSON.parse(itemFromList?.CommitteeMeetingMembersDTO).filter(
    //   (each:any)=>each.isChairman === false
    // );
    const auditTrail = JSON.parse(itemFromList?.AuditTrail);


    auditTrail.push({
      action: `Meeting Returned Back`,
      actionBy: this.props.userDisplayName,
      actionDate: this. _formatDateTime(new Date()),
    });

    // const encodeMOMDTO = this._encodeMOMDTO(this.state.selectedCommitteeNoteRecords)
    // console.log(encodeMOMDTO,"encode MOM DTO")

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({

    
        
        CommitteeMeetingMembersDTO: this._getCommitteeMeetingMembersDTO(
          this.state.committeeMembersData
        ),
        CommitteeMeetingNoteDTO: JSON.stringify(
          this.state.committeeNoteRecordsData
        ),

        CurrentApproverId: this._getCurrrentApproverId(
          this.state.committeeMembersData
        ),

        MeetingStatus: "Return Back",
        StatusNumber: "5000",
        AuditTrail: JSON.stringify(auditTrail),
        startProcessing: true,
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
       
      });

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        MeetingStatus: "Pending Approval",
      
        startProcessing: true,
        CurrentApproverId: this._getCurrrentApproverId(
          this.state.committeeMembersData
        ),
      });

    this.setState({
      isLoading: false,
      isModalOpen: true,
      dialogType: "success",
    });
  };

  private _handleMeetingOver = async (): Promise<void> => {
    this.setState({ isModalOpen: false, isLoading: true });
    const auditTrail = this.state.auditTrail || [];

    auditTrail.push({
      action: `Meeting Over`,
      actionBy: this.props.userDisplayName,
      actionDate: this. _formatDateTime(new Date()),
    });

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        StatusNumber: "3000",
        AuditTrail: JSON.stringify(auditTrail),
        startProcessing: true,
        MeetingStatus: "Meeting Over",
        FinalApproverId:
          this.state.committeeMembersData[
            this.state.committeeMembersData.length - 1
          ].userId,
        CommitteeMeetingMembersDTO: this._getCommitteeMeetingMembersDTO(
          this.state.committeeMembersData
        ),
        CommitteeMeetingGuestMembersDTO: this._getCommitteeGuestMembersDTO(
          this.state.committeeGuestMembersData
        ),
        CommitteeMeetingNoteDTO: JSON.stringify(
          this.state.committeeNoteRecordsData
        ),
        CurrentApproverId: (await this.props.sp?.web.currentUser())?.Id,
        CommitteeMeetingMembersId: this._getMemeberId(
          this.state.committeeMembersData
        ),
        CommitteeMeetingGuestsId: this._getMemeberId(
          this.state.committeeGuestMembersData
        ),
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    this.setState({
      isLoading: false,
      isModalOpen: true,
      dialogType: "success",
    });
  };


  // private decodeAndSanitize=(encodedString: any): any=> {
  // //   const parser = new DOMParser();
  // // const doc = parser.parseFromString(encodedString, "text/html");

  // // // Extract the inner HTML, which will decode the entities
  // // return doc.body.innerHTML;

  // const entityMap: { [key: string]: string } = {
  //   "&amp;": "&",
  //   "&lt;": "<",
  //   "&gt;": ">",
  //   "&quot;": '"',
  //   "&#39;": "'",
  //   "&apos;": "'",
  //   "&nbsp;": " ",
  //   "&cent;": "¢",
  //   "&pound;": "£",
  //   "&yen;": "¥",
  //   "&euro;": "€",
  //   "&sect;": "§",
  //   "&copy;": "©",
  //   "&reg;": "®",
  //   "&trade;": "™",
  //   "&deg;": "°",
  //   "&plusmn;": "±",
  //   "&para;": "¶",
  //   "&middot;": "·",
  //   "&ndash;": "–",
  //   "&mdash;": "—",
  //   "&hellip;": "…",
  //   "&laquo;": "«",
  //   "&raquo;": "»",
  //   "&lsquo;": "‘",
  //   "&rsquo;": "’",
  //   "&ldquo;": "“",
  //   "&rdquo;": "”",
  //   "&bull;": "•",
  //   "&divide;": "÷",
  //   "&times;": "×",
  //   "&sup2;": "²",
  //   "&sup3;": "³",
  //   "&frac12;": "½",
  //   "&frac14;": "¼",
  //   "&frac34;": "¾",
  //   "&#8203;": "",
  // };

  // // Replace all the entities with their corresponding characters
  // return encodedString.replace(/&[a-zA-Z0-9#]+;/g, (entity: any) => entityMap[entity] || entity);
  // }

  // private _encodeMOMDTO =(data:any):any=>{
  //   const encodeMOMDTO = this.state.committeeNoteRecordsData.map(
  //     (each:any)=>{

  //       const decodedHtml = he.decode(each.mom)
  //       console.log(decodedHtml)

  //       return {...each,mom:decodedHtml}
  //     }

      
  //   )
  //   console.log(encodeMOMDTO,"encode MOM DTO")
  //   return encodeMOMDTO

  // }

  private _handleMOMPublished = async (): Promise<void> => {
    this.setState({ isModalOpen: false, isLoading: true });
    const auditTrail = this.state.auditTrail || [];

    
    

    auditTrail.push({
      action: `Meeting MOM Published`,
      actionBy: this.props.userDisplayName,
      actionDate: this. _formatDateTime(new Date()),
    });

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        StatusNumber: "4000",

        MeetingStatus: "MOM Published",
        AuditTrail: JSON.stringify(auditTrail),
        CommitteeMeetingNoteDTO: JSON.stringify(
          this.state.committeeNoteRecordsData
        ),
        CurrentApproverId: this._getCurrrentApproverId(
          this.state.committeeMembersData
        ),
      
        startProcessing: true,
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        StatusNumber: "5000",
        MeetingStatus: "Pending Approval",
        CurrentApproverId: this._getCurrrentApproverId(
          this.state.committeeMembersData
        ),

        startProcessing: true,
      });

    await this.state.committeeNoteRecordsData.map(async (each: any) => {
      await this.props.sp.web.lists
        .getByTitle("EcommiteeRequests")
        .items.getById(each.id)
        .update({
          isMapped: true,
        });
    });
    await this._CreateFolder();
   
    this.setState({
      isLoading: false,
      isModalOpen: true,
      dialogType: "success",
    });
  };

  private options: IDropdownOption[] = [
    { key: "teams", text: "Teams" },
    { key: "webex", text: "WebEx" },
  ];

  private  _CreateFolder=async(): Promise<void>=> {
    const MeetingNumber =this.state.meetingId.replace(/\//g, "-");
    try {

      const absUrl = this.props.context.pageContext.web.serverRelativeUrl;


      const parentFolderPath = `${absUrl}/CommitteeMeetingDocuments`;
      const folderPath = `${parentFolderPath}/${MeetingNumber}`;


      try {
        const folder = await this.props.sp.web.getFolderByServerRelativePath(folderPath).select("Exists")();
        if (folder.Exists) {
          return;
        }else{
          await this.props.sp.web.folders.addUsingPath(folderPath)
        }
      } catch (error) {
     
        return error
      }
  
    } catch (error) {
      
      return error
    }
  }

  /* // Event handler for the dropdown */
  private handleMeetingModeChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ meetingMode: option.key });
    }
  };

 /*  // Event handler for handling changes in both textareas */
  private handleInputChange = (
    event: React.ChangeEvent<HTMLTextAreaElement>,
    field: "meetingSubject" | "meetingLink"
  ): void => {
    const { value } = event.target;
    this.setState({ [field]: value } as Pick<
      CommtteeMeetingsStateProps,
      "meetingSubject" | "meetingLink"
    >);
  };

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

  private handleCommitteeNoteRecordsDropdownChange = (
    event: any,
    option: any
  ) => {
    this.setState({ committeeNoteRecordSelectedValue: option.key });
  };

  private _handleOnAddCommitteeNoteRecords = (): any => {
    this.state.committeeNoteRecordDropDownData.filter(
      (each: any) => each.key === this.state.committeeNoteRecordSelectedValue
    );
    const fiterSelected = this.state.committeeNoteRecordsData.filter(
      (each: any) => {
        if (each.key === this.state.committeeNoteRecordSelectedValue) {
          return each;
        }
      }
    );

    if (fiterSelected[0]?.key === this.state.committeeNoteRecordSelectedValue) {
      this.setState({
        isModalOpen: true,
        modalMessage:
          "Please select Another Note Record, which is already Selected",
      });

      return;
    }
    this.setState(
      (prevState) => ({
        committeeNoteRecordsData: [
          ...prevState.committeeNoteRecordsData,
          ...prevState.committeNoteRecordsDropDownDataWithAllProperties.filter(
            (each: any) => each.key === prevState.committeeNoteRecordSelectedValue
          ),
        ],
      }),)
  };

  private _checkAllNoteRecord = (): boolean => {
    const isAnyMOMEmpty = this.state.committeeNoteRecordsData.some(
      (each: any) => {

        const htmlContent = each.mom
    const parser = new DOMParser();
    const doc:any = parser.parseFromString(htmlContent, 'text/html');
    const contentWithoutPTags = doc.body.textContent.trim();
  
        return each.mom === "" || contentWithoutPTags === "";
      }
    );

    return isAnyMOMEmpty;
  };

  private stylesModal = mergeStyleSets({
    modal: {
      padding: "10px",
      minWidth: "300px",
      maxWidth: "80vw",
      width: "100%",
      "@media (min-width: 768px)": {
        maxWidth: "580px", // Adjust width for medium screens
      },
      "@media (max-width: 767px)": {
        maxWidth: "290px", // Adjust width for smaller screens
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
      minHeight: "50px",
      marginBottom: "20px",
    },
    headerTitle: {
      margin: "5px",
      marginLeft: "5px",
      fontSize: "16px",
      fontWeight: "400",
    },
    peoplePickerAndAddCombo: {
      display: "flex",
      gap: "5px",
      width: "60%",
    },
    body: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      textAlign: "center",
      padding: "20px 0",
    },
    footer: {
      display: "flex",
      justifyContent: "flex-end",
      marginTop: "20px",
      borderTop: "1px solid #ddd", // Added border to the top of the footer
      paddingTop: "10px",
    },
  });

  private stylesMOMModal = mergeStyleSets({
    modal: {
      padding: "10px",
      width: "580px",

      "@media (min-width: 768px)": {
        maxWidth: "580px", // Adjust width for medium screens
      },
      "@media (max-width: 767px)": {
        maxWidth: "290px", // Adjust width for smaller screens
      },

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
    },
    headerTitle: {
      margin: "5px",
      marginLeft: "5px",
      fontSize: "16px",
      fontWeight: "400",
    },
    body: {
      marginTop: "48px",
      position: "relative",
      height: "150px",
    },

    footer: {
      display: "flex",
      justifyContent: "flex-end",
      marginTop: "20px",
      borderTop: "1px solid #ddd", // Added border to the top of the footer
      padding: "10px",
    },
  });

  private _getAlertDialogContent = (): any => {
    return (
      <>
        <div className={this.stylesModal.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Info" }} />
            <h4 className={this.stylesModal.headerTitle}>Alert</h4>
          </div>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={() => {
              this._closeModal();
            }}
          />
        </div>
        <div className={this.stylesModal.body}>
          <p>{this.state.modalMessage}</p>
        </div>
        <div className={this.stylesModal.footer}>
          <PrimaryButton
            iconProps={{ iconName: "ReplyMirrored" }}

            onClick={() => {
       
              this._closeModal();
            }}
            text="OK"
          />
        </div>
      </>
    );
  };

  private _onRichTextChangeForMom = (newText: string) => {
   

   
    // console.log(newText," ,newText")

    // const parser = new DOMParser();
    //   const doc = parser.parseFromString(newText, 'text/html');
    //   const plainText = doc.body.textContent || "";
    //   console.log(plainText,"")

    //   let encoded = encodeURIComponent(newText);
    //   console.log(encoded, "encoded..........")
    //   let decoded = decodeURIComponent(encoded);
    //   console.log(decoded,"decode .............")
    // const decodedHtml = he.decode(newText)
    // console.log(decodedHtml)
 
    this.setState({ draftResolutionFieldValue: newText });
    return newText;
  };

  private _getConfirmationtDialogContent = (): any => {
    const styles = mergeStyleSets({
      modal: {
        minWidth: "300px",
        maxWidth: "80vw",
        width: "100%",
        "@media (min-width: 768px)": {
          maxWidth: "580px", // Adjust width for medium screens
        },
        "@media (max-width: 767px)": {
          maxWidth: "290px", // Adjust width for smaller screens
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
        paddingRight: "0px", // Reduced space between the icon and the title
      },
      body: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        textAlign: "center",

        "@media (min-width: 768px)": {
          marginLeft: "20px", // Adjust width for smaller screens
          marginRight: "20px", // Adjust width for medium screens
          height: "160px",
        },
        "@media (max-width: 767px)": {
          marginLeft: "20px", // Adjust width for smaller screens
          marginRight: "20px",
          height: "190px",
        },
      },
      footer: {
        display: "flex",
        justifyContent: "space-between", // Adjusted to space between

        borderTop: "1px solid #ddd",
        paddingTop: "10px",
        minHeight: "50px",
      },
      button: {
        maxHeight: "32px",
        flex: "1 1 50%", // Ensures each button takes up 50% of the footer width
        margin: "0 5px", // Adds some space between the buttons
      },
      buttonContent: {
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      },
      buttonIcon: {
        marginRight: "4px", // Adjust the space between the icon and text
      },

      removeTopMargin: {
        marginTop: "4px",
        marginBottom: "4px",
        fontWeight: "400",
      },
    });
    return (
      <>
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton
              iconProps={{ iconName: "WaitlistConfirm" }}
              className={styles.headerIcon}
            />
            <h4 className={styles.headerTitle}>Confirmation</h4>
          </div>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={this._closeModal}
          />
        </div>
        <div className={styles.body}>
          <p className={`${styles.removeTopMargin}`}>
            {(() => {
              switch (this.state.btnType) {
                case "createMeeting":
                  return "Are you sure you want to Create this Meeting?";
                case "publishMeeting":
                  return "Are you sure you want to Publish this Meeting?";
                case "meetingOver":
                  return "Are you sure you want to submit meeting over?";
                case "momPublished":
                  return "Are you sure you want to Publish MOM?";

                case "returnBack":
                  return "Are you sure you want to Return Back the Meeting?";
              }
            })()}
          </p>
          <br />
          <p className={`${styles.removeTopMargin}`}>
            {(() => {
              switch (this.state.btnType) {
                case "createMeeting":
                  return "Please check the details filled  and click on Confirm button to Create meeting.";
                case "publishMeeting":
                  return "Please check the details filled  and click on Confirm button to Publish meeting.";
                case "meetingOver":
                  return "Please check the details filled  and click on Confirm button to submit meeting over.";
                case "momPublished":
                  return "Please check the details filled  and click on Confirm button to Publish MOM.";

                case "returnBack":
                  return "Please check the details filled  and click on Confirm button to Return Back the meeting.";
              }
            })()}
          </p>
        </div>
        <div className={styles.footer}>
          <PrimaryButton
            onClick={() => {
              switch (this.state.btnType) {
                case "momPublished":
                  return this._handleMOMPublished();
                case "meetingOver":
                  return this._handleMeetingOver();
                case "publishMeeting":
                  return this._handlePulbicMeeting();

                case "returnBack":
                  return this._handleReturnBack();
                default:
                  return this._handleCreateMeeting();
              }
            }}
            text="Confirm"
            iconProps={{
              iconName: "SkypeCircleCheck",
              styles: { root: styles.buttonIcon },
            }}
            styles={{ root: styles.buttonContent }}
            className={styles.button}
          />
          <DefaultButton
            onClick={this._closeModal}
            text="Cancel"
            iconProps={{
              iconName: "ErrorBadge",
              styles: { root: styles.buttonIcon },
            }}
            styles={{ root: styles.buttonContent }}
            className={styles.button}
          />
        </div>
      </>
    );
  };

  private _geSuccessDialogContent = (): any => {
    return (
      <>
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Info" }} />
            <h4 className={styles.headerTitle}>Alert</h4>
          </div>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={() => {
              this._closeModal();
              const pageURL: string = this.props.homePageUrl;
              window.location.href = `${pageURL}`;
            }}
          />
        </div>
        <div className={styles.body}>
          <p>
            {(() => {
              switch (this.state.btnType) {
                case "createMeeting":
                  return "Meeting has been created successfully.";
                case "publishMeeting":
                  return "Meeting has been published successfully.";
                case "meetingOver":
                  return "Meeting over has been submitted successfully.";
                case "momPublished":
                  return "Meeting minutes has been published successfully.";

                case "returnBack":
                  return "Return back has been successfully.";
              }
            })()}
          </p>
        </div>
        <div className={styles.footer}>
          <PrimaryButton
            className={styles.button}
            iconProps={{ iconName: "ReplyMirrored" }}
            onClick={() => {
              this._closeModal();
              const pageURL: string = this.props.homePageUrl;
              window.location.href = `${pageURL}`;
            }}
            text="OK"
          />
        </div>
      </>
    );
  };

  private _getValidationDialog = (): any => {
    return (
      <>
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Info" }} />
            <h4 className={styles.headerTitle}>Alert</h4>
          </div>
          <Icon iconName="Cancel" onClick={this._closeModal} />
        </div>
        <div className={styles.body} style={{ alignItems: "flex-start" }}>
          <h4 className={styles.headerTitle}>
            Please fill up all the mandatory fields
          </h4>
          <ul>
            {this.state.invalidFields.length > 0 &&
              this.state.invalidFields.map((each: any) => (
                <li style={{ textAlign: "left" }} key={each}>
                  {each}
                </li>
              ))}
          </ul>
        </div>

        <div className={styles.footer}>
          <PrimaryButton
            text="OK"
            iconProps={{ iconName: "ReplyMirrored" }}
            onClick={this._closeModal}
          />
        </div>
      </>
    );
  };

  private _getBodyContentOfDialogBox = (): any => {
    switch (this.state.dialogType) {
    

      case "confirmation":
        return this._getConfirmationtDialogContent();

      case "success":
        return this._geSuccessDialogContent();

      case "validation":
        return this._getValidationDialog();
      default:
        return this._getAlertDialogContent();
    }
  };

  private _checkFields = (): boolean => {
    const fieldValues: any = {
      committeeName: [
        this.state.committeeNameFeildValue,
        "Committee Name",
        "isWarningCommitteeName",
      ],
      convenor: [
        this.state.convernorFeildValue,
        "Convenor",
        "isWarningConvenor",
      ],
      chairman: [
        this.state.charimanFeildValue,
        "Chairman",
        "isWarningChairman",
      ],
      meetingDate: [
        this.state.meetingDate,
        "Meeting Date",
        "isWarningMeetingDate",
      ],
      meetingSubject: [
        this.state.meetingSubject,
        "Meeting Subject",
        "isWarningMeetingSubject",
      ],
      meetingMode: [
        this.state.meetingMode,
        "Meeting Mode",
        "isWarningMeetingMode",
      ],
      meetingLink: [
        this.state.meetingLink,
        "Meeting Link",
        "isWarningMeetingLink",
      ],
      committeeMembers: [
        this.state.committeeMembersData,
        "Please select Committee Members",
        "isWarningCommitteeMembers",
      ],
 
      committeeNoteRecords: [
        this.state.committeeNoteRecordsData,
        "Please select Committee Note Records",
        "isWarningCommitteeNoteRecords",
      ],
    };

    const invalid: string[] = [];
    const warnings: Record<string, boolean> = {};

    /* // Check each field's value */
    Object.values(fieldValues).forEach(
      ([value, displayName, warningKey]: any) => {

        if (
          value === null ||
          value === undefined ||
          value === "" ||
          (Array.isArray(value) && value.length === 0)
        ) {
          invalid.push(displayName as string); // Push invalid field name
          warnings[warningKey as string] = true; // Set warning to true for the invalid field
        } else {
          warnings[warningKey as string] = false; // Clear warning for valid fields
        }
      }
    );

    this.setState({ invalidFields: invalid, ...warnings });

    const isValid = invalid.length === 0;
    return isValid;
  };

  private getUserDepartmentByEmail = async (
    id: any
  ): Promise<string | null> => {
    try {
      const userProfile = await this.props.sp.web.getUserById(id)();

      const profile = await this.props.sp.profiles.getPropertiesFor(
        userProfile?.LoginName
      );
      
    

      const departmentProperty = profile.UserProfileProperties.find(
        (element: any) => element.Key === "Department"
      );

      const department = departmentProperty?.Value || null;
      this.setState({ convernorFeildValue: department });

      return department;
    } catch (error) {
      console.error("Error fetching user profile:", error);
      return null;
    }
  };

  public render(): React.ReactElement<IXenWpCommitteeMeetingsFormsProps> {
  console.log(window.innerWidth)

    console.log(this.state.isSmallScreen)
    return (
      <div>
        <div className={styles.titleContainer}>
          <div
            className={`${styles.noteTitle}
          
          `}
          >
            <div className={styles.statusContainer}>
              {this._itemId ? (
                <p className={styles.status}>
                  Status:
                  {

                    this.state.MeetingStatus
                  }
                </p>
              ) : (
                ""
              )}
            </div>
            <h1 className={styles.title}>
              {this._itemId
                ? `eCommittee Meeting -${this.state.meetingId}`
                : `eCommittee Meeting -${this.props.formType}`}
            </h1>
            <p className={styles.titleDate}>
              {" "}
              <DateTime />
            </p>
          </div>

          <span className={styles.field}>
            All fields marked "*" are mandatory
          </span>
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
            <label className={styles.label} htmlFor="_MeetingId">
              Meeting ID :<span className={styles.warning}>*</span>
            </label>
            <TextField
            id="_MeetingId"
              type="text"
              className={styles.textField}
              value={this.state.meetingId}
            />
          </div>

          {/* Committee Name Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_CommitteeName">
              Committee Name :<span className={`${styles.warning}`}>*</span>
            </label>
            <Dropdown
            id="_CommitteeName"
              placeholder="Select an option"
              onRenderCaretDown={() =>
                this.onRenderCaretDowncommitteeNameFeildValue()
              }
              onChange={this.handleCommittename}
              className={styles.dropdown}
              options={this.state.committename}
              selectedKey={this.state.committeeNameFeildValue}
              styles={{
                dropdown: {
                  border:
                    this.state.committeeNameFeildValue === "" &&
                    this.state.isWarningCommitteeName
                      ? "2px solid red"
                      : "1px solid transparent",
                },
                title: {
                  borderColor:
                    this.state.committeeNameFeildValue === "" &&
                    this.state.isWarningCommitteeName
                      ? "transparent"
                      : undefined,
                },
              }}
            />
          </div>

          {/* Convenor Department : Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_ConvernorDpt">
              Convenor Department :<span className={styles.warning}>*</span>
            </label>
            <TextField
            id="_ConvernorDpt"
              type="text"
              className={styles.textField}
              value={this.state.convernorFeildValue}
              styles={{
                fieldGroup: {
                  border:
                    !this.state.convernorFeildValue &&
                    this.state.isWarningConvenor
                      ? "2px solid red"
                      : "1px solid",
                },
              }}
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
              value={this.state.charimanFeildValue}
              styles={{
                fieldGroup: {
                  border:
                    !this.state.charimanFeildValue &&
                    this.state.isWarningChairman
                      ? "2px solid red"
                      : "1px solid",
                },
              }}
            />
          </div>

          {/* Meeting Date: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_meetingDate">
              Meeting Date :<span className={styles.warning}>*</span>
            </label>
            <DatePicker
            id="_meetingDate"
              placeholder="Select a date..."
              value={this.state.meetingDate}
              onSelectDate={this.onDateChange}
              ariaLabel="Select a date"
              minDate={new Date()}
              strings={defaultDatePickerStrings}
              styles={{
                root: {
                  height: "35px",
                  border:
                    this.state.meetingSubject === "" &&
                    this.state.isWarningMeetingSubject
                      ? "2px solid red"
                      : "",
                },
              }}
            />
          </div>

          {/* Meeting Subject: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_meetingSubject">
              Meeting Subject :<span className={styles.warning}>*</span>
            </label>
            <textarea
            id="_meetingSubject"
              className={styles.textarea}
              value={this.state.meetingSubject}
              onChange={(event) =>
                this.handleInputChange(event, "meetingSubject")
              }
              style={{
                border:
                  this.state.meetingSubject === "" &&
                  this.state.isWarningMeetingSubject
                    ? "2px solid red"
                    : "",
              }}
            />
          </div>

          {/* Meeting Mode : Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label} htmlFor="_meetingMode">
              Meeting Mode :<span className={`${styles.warning}`}>*</span>
            </label>
            <Dropdown
            id="_meetingMode"
              placeholder="Select an option"
              options={this.options}
              selectedKey={this.state.meetingMode}
              onChange={this.handleMeetingModeChange}
              className={styles.dropdown}
              styles={{
                dropdown: {
                  // width: 300,

                  border:
                    this.state.meetingMode === "" &&
                    this.state.isWarningMeetingMode
                      ? "2px solid red"
                      : "1px solid transparent",
                },
                title: {
                  borderColor:
                    this.state.meetingMode === "" &&
                    this.state.isWarningMeetingMode
                      ? "transparent"
                      : undefined,
                },
              }}
            />
          </div>

          {/* Meeting Link: Section */}
          <div className={styles.halfWidth} style={{ paddingTop: "10px" }}>
            <label className={styles.label} htmlFor="_meetingLink">
              Meeting Link :<span className={styles.warning}>*</span>
            </label>
            <textarea
            id="_meetingLink"
              className={styles.textareaForMeetingLink}
              value={this.state.meetingLink}
              onChange={(event) => this.handleInputChange(event, "meetingLink")}
              style={{
                border:
                  this.state.meetingLink === "" &&
                  this.state.isWarningMeetingLink
                    ? "2px solid red"
                    : "",
              }}
            />
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
            {
            // (this.state.statusNumber === "" ||
            //   this.state.statusNumber === "1000" ||
            //   this.state.statusNumber === "2000") &&
               (
              <div className={`${styles.peoplePickerAndSpanContainer}`}>
                <div style={{ display: "flex", flexWrap: "wrap" }}>
                  <PeoplePicker
                    key={this.state.committeeMemberskey}
                    placeholder="Add Member..."
                    context={this._peopplePicker}
                    // titleText="People Picker"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users
                    showtooltip={true}
                    defaultSelectedUsers={[""]}
                    disabled={false}
                    ensureUser={true}
                    onChange={this._getPeoplePickerItemsCommitteeMembers}
                    // showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  />
                  {/* <PeoplePicker /> */}
                  <DefaultButton
                    style={{ marginTop: "0px", marginLeft: "6px" }}
                    type="button"
                    className={`${styles.responsiveButton}`}
                    onClick={(e) => {
                      if (this.state.selectedCommitteeMembers.length === 0) {
                        this.setState({
                          isModalOpen: true,
                          modalMessage:
                            "Please select Member then click on Add.",
                        });
                        this._clearCommitteeMembersPeoplePicker();
                        return;
                      }
                      this.handleOnAdd(e, "committeeMembers");
                    }}
                   
                  >
                    + Add
                  </DefaultButton>
                </div>
                <span className={`${styles.spanForPeoplePicker}`}>
                  (Please enter minimum 3 character to search)
                </span>
              </div>
            )}
          </div>
          <div style={{ overflowX: "auto" }}>
            <DetailsList
              items={this.state.committeeMembersData} // Data for the table
              columns={this.columnsCommitteeMembers} // Columns for the table
              layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
              selectionMode={SelectionMode.none} // No selection column
              isHeaderVisible={true} // Show column headers
              onRenderDetailsFooter={(props: IDetailsFooterProps) => {
                if (this.state.committeeMembersData.length === 0) {
                  return (
                    <div
                      style={{
                        textAlign: "center",
                        padding: "20px",
                        color: "gray",
                      }}
                    >
                      No records available
                    </div>
                  );
                }
                return null;
              }}
            />
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
            {
            // (this.state.statusNumber === "" ||
            //   this.state.statusNumber === "1000" ||
            //   this.state.statusNumber === "2000") && 
              (
              <div className={`${styles.peoplePickerAndSpanContainer}`}>
                <div style={{ display: "flex", flexWrap: "wrap" }}>
                  <PeoplePicker
                    key={this.state.committeeGuestMemberskey}
                    placeholder="Add Member..."
                    context={this._peopplePicker}
                    // titleText="People Picker"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users
                    showtooltip={true}
                    defaultSelectedUsers={[""]}
                    disabled={false}
                    ensureUser={true}
                    onChange={this._getPeoplePickerItemsCommitteeGuestMembers}
                    // showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  />
                  {/* <PeoplePicker /> */}
                  <DefaultButton
                    style={{ marginTop: "0px", marginLeft: "6px" }}
                    type="button"
                    className={`${styles.responsiveButton}`}
                    onClick={(e) => {
                      if (
                        this.state.selectedCommitteeGuestMembers.length === 0
                      ) {
                        this.setState({
                          isModalOpen: true,
                          modalMessage:
                            "Please select Guest member then click on Add.",
                        });
                        this._clearCommitteeMembersPeoplePicker();
                        return;
                      }

                      this.handleOnAdd(e, "committeeGuestMembers");
                    }}
                    
                  >
                    + Add
                  </DefaultButton>
                </div>
                <span className={`${styles.spanForPeoplePicker}`}>
                  (Please enter minimum 3 character to search)
                </span>
              </div>
            )}

            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.committeeGuestMembersData} // Data for the table
                columns={this.columnsCommitteeGuestMembers} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
                onRenderDetailsFooter={(props: IDetailsFooterProps) => {
                  if (this.state.committeeGuestMembersData.length === 0) {
                    return (
                      <div
                        style={{
                          textAlign: "center",
                          padding: "20px",
                          color: "gray",
                        }}
                      >
                        No records available
                      </div>
                    );
                  }
                  return null;
                }}
              />
            </div>
          </div>
        </div>
        {/* Committee Note  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          // style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Committee Note Records
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            {
            // (this.state.statusNumber === "" ||
            //   this.state.statusNumber === "1000" ||
            //   this.state.statusNumber === "2000") && 
              (
              <div className={`${styles.peoplePickerAndSpanContainer}`}>
                <div style={{ display: "flex", flexWrap: "wrap" }}>
                  <Dropdown
                    selectedKey={this.state.committeeNoteRecordSelectedValue}
                    onChange={this.handleCommitteeNoteRecordsDropdownChange}
                    options={this.state.committeeNoteRecordDropDownData}
                    placeholder="Add Note Record..."
                    styles={{
                      root: {
                        minWidth: "180px",
                      },
                    }}
                  />
                  {/* <PeoplePicker /> */}
                  <DefaultButton
                    style={{ marginTop: "0px", marginLeft: "6px" }}
                    type="button"
                    className={`${styles.responsiveButton}`}
                    // onClick={(e) => this.handleOnAdd(e, "committeeNoteRecords")}
                    
                    onClick={() => {
                      if (this.state.committeeNoteRecordSelectedValue === "") {
                        this.setState({
                          isModalOpen: true,
                          modalMessage: "Please select Note then click on Add.",
                        });
                        this._clearCommitteeMembersPeoplePicker();
                        return;
                      }
                      this._handleOnAddCommitteeNoteRecords();
                      this.setState({ committeeNoteRecordSelectedValue: "" });
                    }}
                  >
                    + Add
                  </DefaultButton>
                </div>
              </div>
            )}

            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.committeeNoteRecordsData} // Data for the table
                columns={this.columnsCommitteeNoteRecords} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
                onRenderDetailsFooter={(props: IDetailsFooterProps) => {
                  if (this.state.committeeNoteRecordsData.length === 0) {
                    return (
                      <div
                        style={{
                          textAlign: "center",
                          padding: "20px",
                          color: "gray",
                        }}
                      >
                        No records available
                      </div>
                    );
                  }
                  return null;
                }}
              />
            </div>
          </div>
        </div>

        {/* Comments section */}
        {this.state.statusNumber === "7000" && (
          <div
            className={`${styles.generalSectionMainContainer}`}
            style={{ flexGrow: 1, margin: "10 10px" }}
          >
            <h1 className={styles.viewFormHeaderSectionContainer}>Comments</h1>
          </div>
        )}

        {this.state.statusNumber === "7000" && (
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
        )}

        {/* WorkFlow  section */}

        {this._itemId ? (
          <>
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
                    items={this.state.auditTrail} // Data for the table
                    columns={this.columnsCommitteeWorkFlowLog} // Columns for the table
                    layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                    selectionMode={SelectionMode.none} // No selection column
                    isHeaderVisible={true} // Show column headers
                  />
                </div>
              </div>
            </div>
          </>
        ) : (
          ""
        )}

        {/*  Buttons Section */}

        <div className={styles.buttonSectionContainer}>
          {this.state.statusNumber !== "1000" &&
            this.state.statusNumber !== "2000" &&
            this.state.statusNumber !== "3000" &&
            this.state.statusNumber !== "4000" &&
            this.state.statusNumber !== "5000" &&
            this.state.statusNumber !== "6000" &&
            this.state.statusNumber !== "7000" &&
            this.state.statusNumber !== "9000" && (
              <PrimaryButton
                // type="button"
                onClick={() => {
                  if (!this._checkFields()) {
                    this.setState({
                      isModalOpen: true,
                      dialogType: "validation",
                      modalMessage: "Please Enter Reqired Fields",
                    });

                    return;
                  }
                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "createMeeting",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "PageHeaderEdit" }}
              >
                Create Meeting
              </PrimaryButton>
            )}

          {((this.state.statusNumber === "1000" && this.state._IsAuthor)||this.state.statusNumber ==="") &&
        

           (
              <PrimaryButton
                // type="button"

                onClick={async() => {

                  if (this._itemId){
                    const item = await this._getItemDataFromList(this._itemId);
                    
                    const StatusNumber = item?.StatusNumber;
                    const status = item?.MeetingStatus.toLowerCase()

                    if (StatusNumber === '2000' || StatusNumber ==='3000'|| StatusNumber ==='4000'|| StatusNumber ==='5000'|| StatusNumber ==='6000' || StatusNumber ==='7000'|| StatusNumber ==='9000') {
                      if (parseInt(StatusNumber) >= 4000){
                        this.setState({
                     
                          hideParellelActionAlertDialog: true,
                          parellelActionAlertMsg:
                            `This meeting has been MOM published.`,
                        });

                      }else{
                        this.setState({
                     
                          hideParellelActionAlertDialog: true,
                          parellelActionAlertMsg:
                            `This meeting has been ${status.toLowerCase()}.`,
                        });
                      }
                
                      return;
                    }
                    

                  }

                  
                  if (!this._checkFields()) {
                    this.setState({
                      isModalOpen: true,
                      dialogType: "validation",
                      modalMessage: "Please Enter Reqired Fields",
                    });

                    return;
                  }

                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "publishMeeting",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "EntitlementRedemption" }}
              >
                Publish Meeting
              </PrimaryButton>
            )}

          {this.state.statusNumber === "7000" && 
          this.state._IsAuthor &&(
            <PrimaryButton
              // type="button"

              onClick={async() => {

                if (this._itemId){
                  const item = await this._getItemDataFromList(this._itemId);
          

                  const StatusNumber = item?.StatusNumber;
                  const status = item?.MeetingStatus.toLowerCase()

                  if (StatusNumber === '2000' || StatusNumber ==='3000'|| StatusNumber ==='4000'|| StatusNumber ==='5000'|| StatusNumber ==='6000' || StatusNumber ==='9000') {
                    if (parseInt(StatusNumber) >= 4000){
                      this.setState({
                   
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This meeting has been MOM published.`,
                      });

                    }else{
                      this.setState({
                   
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This meeting has been ${status.toLowerCase()}.`,
                      });
                    }
              
                    return;
                  }
                  

                }





                if (!this._checkFields()) {
                  this.setState({
                    isModalOpen: true,
                    dialogType: "validation",
                    modalMessage: "Please Enter Reqired Fields",
                  });

                  return;
                }

                if (this._checkAllNoteRecord()) {
                  this.setState({
                    isModalOpen: true,
                    modalMessage:
                      "MOM cannot be blank, Please add MOM for all the listed Notes.",
                  });

                  return;
                }

                this.setState({
                  isModalOpen: true,
                  dialogType: "confirmation",
                  btnType: "returnBack",
                });
              }
            } 
            
              className={`${styles.responsiveButton} `}
              iconProps={{ iconName: "EntitlementRedemption" }}
            >
              Return Back
            </PrimaryButton>
          )}

          {(
            this.state.statusNumber === "2000") &&
            this.state._IsAuthor && (
            <PrimaryButton
              onClick={async () => {

                if (this._itemId){
                  const item = await this._getItemDataFromList(this._itemId);
                 

                  const StatusNumber = item?.StatusNumber;
                  const status = item?.MeetingStatus.toLowerCase()

                  if ( StatusNumber ==='3000'|| StatusNumber ==='4000'|| StatusNumber ==='5000'|| StatusNumber ==='6000' || StatusNumber ==='7000'|| StatusNumber ==='9000') {
                    if (parseInt(StatusNumber) >= 4000){
                      this.setState({
                   
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This meeting has been MOM published.`,
                      });

                    }else{
                      this.setState({
                   
                        hideParellelActionAlertDialog: true,
                        parellelActionAlertMsg:
                          `This meeting has been ${status.toLowerCase()}.`,
                      });
                    }
              
                    return;
                  }
                  

                }

                if (this.state.statusNumber !== "2000") {
                  if (this._checkAllNoteRecord()) {
                    this.setState({
                      isModalOpen: true,
                      modalMessage:
                        "Please publish the meeting to record meeting over.",
                    });

                    return;
                  }
                }
                this.setState({
                  isModalOpen: true,
                  dialogType: "confirmation",
                  btnType: "meetingOver",
                });
              }}
              className={`${styles.responsiveButton} `}
              iconProps={{ iconName: "EntitlementRedemption" }}
            >
              Meeting Over
            </PrimaryButton>
          )}

          {this.state.statusNumber === "3000" &&
            this.state._IsAuthor && (
              <PrimaryButton
                // type="button"
                onClick={async () => {

                  if (this._itemId){
                    const item = await this._getItemDataFromList(this._itemId);
                 
  
                    const StatusNumber = item?.StatusNumber;
                    const status = item?.MeetingStatus.toLowerCase()
  
                    if (StatusNumber ==='4000'|| StatusNumber ==='5000'|| StatusNumber ==='6000' || StatusNumber ==='7000'|| StatusNumber ==='9000') {
                     
                      if (parseInt(StatusNumber) >= 4000){
                        this.setState({
                     
                          hideParellelActionAlertDialog: true,
                          parellelActionAlertMsg:
                            `This meeting has been MOM published.`,
                        });

                      }else{
                        this.setState({
                     
                          hideParellelActionAlertDialog: true,
                          parellelActionAlertMsg:
                            `This meeting has been ${status.toLowerCase()}.`,
                        });
                      }
                     
                     
                      
                
                      return;
                    }
                    
  
                  }

                  if (this._checkAllNoteRecord()) {
                    this.setState({
                      isModalOpen: true,
                      modalMessage:
                        "MOM cannot be blank, Please add MOM for all the listed Notes.",
                    });

                    return;
                  }

                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "momPublished",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "EntitlementRedemption" }}
              >
                Mom Published
              </PrimaryButton>
            )}

          <DefaultButton
            // type="button"
            onClick={() => {
              const pageURL: string = this._homePageUrl;
              window.location.href = `${pageURL}`;
            }}
            className={`${styles.responsiveButton} `}
            iconProps={{ iconName: "Cancel" }}
          >
            Exit
          </DefaultButton>
        </div>

        {/* Modal for alerts */}
        <Modal
          isOpen={this.state.isModalOpen}
          onDismiss={this._closeModal}
          isBlocking={true}
          containerClassName={this.stylesModal.modal}
        >
          {this._getBodyContentOfDialogBox()}
        </Modal>

        <Modal
          isOpen={this.state.isModalMOMOpen}
          onDismiss={this._closeModal}
          isBlocking={true}
          containerClassName={this.stylesMOMModal.modal}
        >
          <>
            <div className={this.stylesMOMModal.header}>
              <div style={{ display: "flex", alignItems: "center" }}>
                <IconButton iconProps={{ iconName: "Info" }} />
                <h4 className={this.stylesMOMModal.headerTitle}>ADD MOM</h4>
              </div>
              <IconButton
                iconProps={{ iconName: "Cancel" }}
                onClick={() => {
                  this._closeModal();
                  this.setState({ dialogType: "" });
                }}
              />
            </div>
            <div className={this.stylesMOMModal.body}>
              <div
              //  className={` ${styles.richTextContainer}`}
              >

                {this.state.isSmallScreen?<RichText
                key={this.state.isSmallScreen ? "smallScreen" : "largeScreen"}
                  value={this.state.draftResolutionFieldValue}
                  styleOptions={
                     {
                          showBold: true,
                          showItalic: true,
                          showUnderline: true,
                          showList: true,
                          showMore: true,
                        }
                      
                  }
                  onChange={(text: string) =>
                    this._onRichTextChangeForMom(text)
                  }
                />:<RichText
                key={this.state.isSmallScreen ? "smallScreen" : "largeScreen"}
                  value={this.state.draftResolutionFieldValue}
                  styleOptions={
                    {
                      showBold: true,

                      showItalic: true,
                      showUnderline: true,
                      showList: true,
                      showAlign: true,
                      showImage: true,
                      showLink: true,
                      showStyles: true,

                      showMore: true,
                    }
                  }
                  onChange={(text: string) =>
                    this._onRichTextChangeForMom(text)
                  }
                />}
               
              </div>
            </div>
            <div className={this.stylesMOMModal.footer}>
              <PrimaryButton
                iconProps={{ iconName: "Add" }}
                onClick={() => {
                 

                  this.setState((prevState)=>{

                    const updatedData = prevState.committeeNoteRecordsData.map(
                      (each: any) => {
                        if (each.key === this.state.selectedMOMNoteRecord) {

                          const decodedHtml = he.decode(this.state.draftResolutionFieldValue)
                         
                          return {
                            ...each,
                            mom: decodedHtml,
                            meetingMinutesForRichTextComponent:this.state.draftResolutionFieldValue
                          };
                        }
                        return each;
                      }
                    );

                    return {
                      draftResolutionFieldValue: "",
                      committeeNoteRecordsData: updatedData,
                      dialogType: "",
                    }
                  });

                  this._closeModal();
                }}
                text="Add"
              />
            </div>
          </>
        </Modal>

        {/* duplicate validation */}

        <Modal
                          isOpen={this.state.hideParellelActionAlertDialog}
                          onDismiss={() => {
                            
                            window.location.reload();
                            this.setState((prevState)=>({
                              hideParellelActionAlertDialog:
                                !prevState.hideParellelActionAlertDialog,
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

                              }

                              
                              }
                              text="OK"
                            />
                          </div>
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
      </div>
    );
  }
}

