import * as React from 'react';
import styles from './PageComments.module.scss';
import { IPageCommentsProps } from './IPageCommentsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PageCommentService } from "../services/PageCommentService";
import { IPageComment } from "../entities/IPageComment";

import { ActivityItem, IActivityItemProps, Link, mergeStyleSets, TextField } from 'office-ui-fabric-react';
import { PrimaryButton } from 'office-ui-fabric-react';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

export interface IPageCommentsState {
  comments: IPageComment[];
  userComment: string;
  selectedTab: string;
  filePickerResult: IFilePickerResult;
  showRequiredMessage: boolean
}

const classNames = mergeStyleSets({
  commentItem: {
    marginTop: '10px',
  },
  nameText: {
    fontWeight: 'bold',
  },
  attachment: {
    display: 'inline'
  },
  richText: {
    borderStyle: "solid",
    borderWidth: 1,
    position:"relative"
  }
});

export default class PageComments extends React.Component<IPageCommentsProps, IPageCommentsState> {

  private _services: PageCommentService = null;
  constructor(props: IPageCommentsProps) {
    super(props);
    this.state = {
      comments: [],
      userComment: "",
      selectedTab: "0",
      filePickerResult: null,
      showRequiredMessage: false
    };
    this._services = new PageCommentService(this.props.context);
    this.getPageComments = this.getPageComments.bind(this);
    this.postComment = this.postComment.bind(this);
    this.sortComments = this.sortComments.bind(this);
    this.onTextChange = this.onTextChange.bind(this);
  }

  public componentDidMount(): void {
    this.getPageComments();
  }

  public async postComment() {
    if(this.state.userComment.trim().length > 0){
      await this._services.addNewComment(this.state.userComment, this.state.filePickerResult)
      this.getPageComments();
    }
    else{
      this.setState({showRequiredMessage:true});
    }
  }

  public async getPageComments() {
    let pageComments = await this._services.getPageComments();
    this.setState({ 
      comments: pageComments, selectedTab: "0", 
      userComment: "", filePickerResult: null, showRequiredMessage:false });
  }

  public onTextChange = (text: string) => {
    try {
      this.setState({ userComment: text });
      return text;
    }
    catch {
      return "";
    }
  }

  public sortComments(item?: PivotItem, ev?: React.MouseEvent<HTMLElement>): void {
    let { comments } = this.state;
    switch (item.props.headerText) {
      case "Newest":
        comments = comments.sort((a, b) => (a.timeStamp < b.timeStamp) ? 1 : -1)
        break;
      case "Oldest":
        comments = comments.sort((a, b) => (a.timeStamp > b.timeStamp) ? 1 : -1)
        break;
      default:
        break;
    }

    this.setState({ comments: comments, selectedTab: item.props.itemKey });
  }

  public render(): React.ReactElement<IPageCommentsProps> {
    let { comments, userComment, selectedTab, filePickerResult } = this.state;

    const commentActivityItems: (IActivityItemProps & { key: string | number })[] = comments ? comments.map(comment => {
      return {
        key: comment.Id,
        activityDescription: [
          <Link key={comment.Id} className={classNames.nameText}>
            {comment.userName}
          </Link>,
          <span style={{ fontSize: 10, paddingLeft: 5 }}>{comment.createdOn} </span>
        ],
        activityPersonas: [{ imageUrl: "/_layouts/15/userphoto.aspx?size=L&username=" + comment.userEmail }],
        comments: [
          <span dangerouslySetInnerHTML={{ __html: comment.comment }} style={{fontSize:14}}></span>,
          comment.attachmentFilename.length > 0 ? 
          // <a target="_blank" href={comment.attachmentUrl} style={{
          //   textDecoration:"none"
          // }}>{comment.attachmentFilename}</a> 
           <img src={comment.attachmentUrl} width="400px" height="250px" style={{cursor:'pointer'}}
            onClick={() => window.open(comment.attachmentUrl, '_blank')} ></img>
          // <a href={comment.attachmentUrl} target="_blank" style={{
          //   width:400,height:250,backgroundImage:comment.attachmentUrl
          // }}></a>
          : undefined
        ]
      }
    }) : [];
    let numberOfComments = commentActivityItems.length;
    // display: inline-block; width: 50px; height; 50px; background-image: url('path/to/image.jpg');
    return (
      <div className={styles.pageComments}>
        {/* <h2>{numberOfComments > 0 ? numberOfComments : ""} {numberOfComments > 1 ? " Comments" : " Comment"}</h2> */}
        <h2>Add your journal entry in the text box below</h2>
        <div>
          <RichText key={comments.length}
            value={userComment}
            className={classNames.richText}
            onChange={(text) => this.onTextChange(text)}
          />
          {this.state.showRequiredMessage && userComment.trim().length <= 0
            ? <span style={{ color: "red" }}>Enter your comment</span> : undefined}

          {/* <TextField  value={userComment} 
          onChange={(target,value)=>this.setState({userComment:value})} multiline autoAdjustHeight /> */}
        </div>
        <div style={{ textAlign: 'right', paddingTop: 5 }}>
          <FilePicker panelClassName={classNames.attachment}
            accepts={[".jpg", ".jpeg", ".png"]}
            buttonIcon="Attach" buttonLabel={filePickerResult ? filePickerResult.fileName : ""}
            onSave={(filePickerResult: IFilePickerResult) => { this.setState({ filePickerResult }) }}
            context={this.context} buttonClassName={classNames.attachment}
            hideLinkUploadTab={true} hideOneDriveTab={true} hideOrganisationalAssetTab={true}
            hideRecentTab={true} hideSiteFilesTab={true} hideStockImages={true} hideWebSearchTab={true}
          />
          <PrimaryButton text="Post" onClick={this.postComment} />
        </div>
        <div>
          <Pivot onLinkClick={this.sortComments} selectedKey={selectedTab}>
            <PivotItem headerText="Newest" itemKey="0">
            </PivotItem>
            <PivotItem headerText="Oldest" itemKey="1">
            </PivotItem>
          </Pivot>
        </div>
        <div>
          {commentActivityItems.map((item: { key: string | number }) => (
            <ActivityItem {...item} key={item.key} className={classNames.commentItem} />
          ))}
        </div>
      </div>
    );
  }
}
