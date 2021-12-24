import { IInputs } from "../generated/ManifestTypes";
import { DocumentTypes, Entities } from "./Constants";
import { ITemplateItem } from "./Items";

export class ExportWordDocumentRequest {
    EntityTypeCode: number;
    SelectedTemplate: {};
    SelectedRecords: string;
    constructor(typeCode: number, record: string, template: {}) {
        this.EntityTypeCode = typeCode;
        this.SelectedTemplate = template;
        this.SelectedRecords = record;
    }
    getMetadata() {
        var metadata = {
            boundParameter: null,
            parameterTypes: {
                EntityTypeCode: {
                    typeName: "Edm.Int32",
                    structuralProperty: 1
                },
                SelectedTemplate: {
                    typeName: "Microsoft.Dynamics.CRM.crmbaseentity",
                    structuralProperty: 5
                },
                SelectedRecords: {
                    typeName: "Edm.String",
                    structuralProperty: 1
                }
            },
            operationName: "ExportWordDocument",
            operationType: 0
        };
        return metadata
    };
}

export class ExportPdfDocumentRequest {
    EntityTypeCode: number;
    SelectedTemplate: {};
    SelectedRecords: string;
    constructor(typeCode: number, record: string, template: {}) {
        this.EntityTypeCode = typeCode;
        this.SelectedTemplate = template;
        this.SelectedRecords = record;
    }
    getMetadata() {
        var metadata = {
            boundParameter: null,
            parameterTypes: {
                EntityTypeCode: {
                    typeName: "Edm.Int32",
                    structuralProperty: 1
                },
                SelectedTemplate: {
                    typeName: "Microsoft.Dynamics.CRM.crmbaseentity",
                    structuralProperty: 5
                },
                SelectedRecords: {
                    typeName: "Edm.String",
                    structuralProperty: 1
                }
            },
            operationName: "ExportPdfDocument",
            operationType: 0
        };
        return metadata
    };
}

export class UploadDocumentRequest {
    Content: string;
    Entity: {};
    OverwriteExisting: boolean;
    ParentEntityReference: {};
    FolderPath: string;
    constructor(content: string, entity: {}, overwriteExisting: boolean, parentEntityReference: {}, folderPath: string) {
        this.Content = content;
        this.Entity = entity;
        this.OverwriteExisting = overwriteExisting;
        this.ParentEntityReference = parentEntityReference;
        this.FolderPath = folderPath
    }
    getMetadata() {
        var metadata = {
            boundParameter: null,
            parameterTypes: {
                Content: {
                    typeName: "Edm.String",
                    structuralProperty: 1
                },
                Entity: {
                    typeName: "Microsoft.Dynamics.CRM.sharepointdocument",
                    structuralProperty: 5
                },
                OverwriteExisting: {
                    typeName: "Edm.Boolean",
                    structuralProperty: 1
                },
                ParentEntityReference: {
                    typeName: "Microsoft.Dynamics.CRM.crmbaseentity",
                    structuralProperty: 5
                },
                FolderPath: {
                    typeName: "Edm.String",
                    structuralProperty: 1
                }
            },
            operationName: "UploadDocument",
            operationType: 0
        };
        return metadata
    }
}
export class CreateActivityMimeAttachmentRequest {
    etn: string;
    payload: {};
    constructor(entityName: string, entityObject: {}) {
        this.etn = entityName;
        this.payload = entityObject
    }
    getMetadata() {
        var metadata = {
            boundParameter: null,
            operationType: 2,
            operationName: "Create",
            parameterTypes: {}
        };
        return metadata;
    }
}

export const RetrieveWordTemplates = async (context: ComponentFramework.Context<IInputs>, documentTemplate: string, primaryEntityName: string, convertToPdf: boolean) => {
    
    let templates =await context.webAPI.retrieveMultipleRecords(documentTemplate, `?$select=clientdata,description,${documentTemplate}id,documenttype,name&$filter=associatedentitytypecode eq '${primaryEntityName}' and documenttype eq 2 and status eq false &$orderby=name asc`).then(
        (response: ComponentFramework.WebApi.RetrieveMultipleResponse) => {
            let templates:ITemplateItem[]=[];
            for (let i = 0; i < response.entities.length; i++) {
                let template: ITemplateItem = {
                    templateName: response.entities[i].name,
                    documentTemplateId:documentTemplate=== Entities.DocumentTemplates?response.entities[i].documenttemplateid : response.entities[i].personaldocumenttemplateid,
                    key: i,
                    description: response.entities[i].description,
                    fileTypeExtension: convertToPdf ? DocumentTypes.Pdf.Extension : DocumentTypes.Word.Extension,
                    category: documentTemplate === Entities.DocumentTemplates ? "System Templates" : "User Templates"
                }
                templates.push(template);
            }            
            return templates;
        });
        if(templates){
        return templates;
        }
        else{
            return [];
        }
    }