export interface IPageComment{
    Id:number,
    comment:string,
    userName:string,
    timeStamp:Date,
    userEmail:string,
    createdOn:string,
    attachmentUrl?:string,
    attachmentFilename?:string
}