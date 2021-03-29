import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp/presets/all';
import { IPageComment } from "../entities/IPageComment";
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';

export class PageCommentService {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async addNewComment(comment:string, file:IFilePickerResult){
        try {
            let listName = "PageComments";
            let newListItem = await sp.web.lists.getByTitle(listName).items.add({
                Comment:comment
              });
            if(file){
                let content =  file.downloadFileContent().then(r => {
                    return { name: file.fileName, content: r };
                  })
                await newListItem.item.attachmentFiles.add(file.fileName,(await content).content);
            }
        } catch (err) {
            Promise.reject(err);
        }
    }

    public async getPageComments(): Promise<IPageComment[]> {
        try {
            let listName = "PageComments";
            let selectQuery: any[] = ['Id','Comment','Created','Author/Title','Author/EMail','FieldValuesAsText/Created','AttachmentFiles'];
            let expandQuery:any[] = ['Author','FieldValuesAsText','AttachmentFiles'];
            let listItems= [];
            let items: any;
            items = await sp.web.lists.getByTitle(listName).items
                .select(selectQuery.join())
                .expand(expandQuery.join())
                .orderBy("Created", false)
                .top(2000)
                .get();
            
            let commentItems:IPageComment[] = items.map((item) => {
                return{
                    Id: item.Id,
                    comment:item.Comment,
                    userName:item.Author.Title,
                    timeStamp:item.Created,
                    userEmail:item.Author.EMail,
                    createdOn: item.FieldValuesAsText.Created,
                    attachmentUrl:item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : "",
                    attachmentFilename: item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].FileName : "",
                }
            });
            return commentItems;
        } catch (err) {
            Promise.reject(err);
        }
    }

    public async getMockPageComments(): Promise<IPageComment[]>{
        let items: IPageComment[] = [
            {
                Id:1,
                comment: "Comment 1",
                userName:"Damian, Martin",
                timeStamp:new Date("3/26/2021"),
                userEmail:"",
                createdOn: ""
            },
            {
                Id:2,
                comment: "Comment 2",
                userName:"Wilson, Alex",
                timeStamp:new Date("3/25/2021"),
                userEmail:"",
                createdOn: ""
            },
            {
                Id:3,
                comment: "Comment 3",
                userName:"Webber, Jack",
                timeStamp:new Date("3/24/2021"),
                userEmail:"",
                createdOn: ""
            }
        ]
        return items;
    }
}