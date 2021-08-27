import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";

export default class Service{
    
   public mysitecontext:any;

    public constructor(siteUrl:string,Sitecontext:any){ 
        this.mysitecontext=Sitecontext;
       

        sp.setup({
            sp: {
              baseUrl: siteUrl
              
            },
          });


        
    }
  


    
    public async addItemToSPList(data:any,fileDetails:any):Promise<any>{
        try{
             let listName:string = "OperationsDOC";
             const iar = await sp.web.lists.getByTitle(listName).items.add(data).then(async (item)=>{
                console.log(item);
                console.log(fileDetails);
                const item1: any =  sp.web.lists.getByTitle(listName).items.getById(item.data.ID);
                await item1.attachmentFiles.add(fileDetails.name,fileDetails);
                return item;
             });
              
              
              return iar;
            
        } catch (error) {
            console.log(error);
        }
    }

    public async getUserByLogin(LoginName:string):Promise<any>{
        try{
            const user = await sp.web.siteUsers.getByLoginName(LoginName).get();
            return user;
        }catch(error){
            console.log(error);
        }
    }


    public async uploadFile(fileDetails:any,data:any){

     try
     {

    console.log(this.mysitecontext);
//     sp.web.getFolderByServerRelativeUrl(this.mysitecontext.pageContext.web.serverRelativeUrl + "/OperationsDOC")
//    .files.add(fileDetails.name, fileDetails, true)
//    .then((data) =>{
     
//      console.log(data);
//      alert("File uploaded sucessfully");
//    })
//    .catch((error) =>{
//      alert("Error is uploading");
//    });



const file = await sp.web.getFolderByServerRelativeUrl(this.mysitecontext.pageContext.web.serverRelativeUrl + "/OperationsDOC").files.add(fileDetails.name, fileDetails, true);
const item = await file.file.getItem();

await item.update(data);

//await sp.web.getFileByServerRelativeUrl(this.mysitecontext.pageContext.web.serverRelativeUrl + "/OperationsDOC")



alert("File uploaded sucessfully");
window.location.replace("https://capcoinc.sharepoint.com/sites/TestSPFX");
      


    }

    catch(error){
        console.log(error);
    }

}

}