export const ResourceKeys = {
    DOWNLOAD_BUTTON_LABEL: "DOWNLOAD_BUTTON_LABEL",
    EMAIL_BUTTON_LABEL: "EMAIL_BUTTON_LABEL",
    SAVETOSHAREPOINT_BUTTON_LABEL: "SAVETOSHAREPOINT_BUTTON_LABEL",
    DOWNLOAD_DESCRIPTION: "DOWNLOAD_DESCRIPTION",
    EMAIL_DESCRIPTION: "EMAIL_DESCRIPTION",
    SAVETOSHAREPOINT_DESCRIPTION: "SAVETOSHAREPOINT_DESCRIPTION",
    DOWNLOAD_DEFAULTBUTTON_LABEL: "DOWNLOAD_DEFAULTBUTTON_LABEL",
    DOWNLOAD_DEFAULTBUTTON_TEXT: "DOWNLOAD_DEFAULTBUTTON_TEXT",
    DIALOG_TITLE: "DIALOG_TITLE",
    PDF_TOGGLE_LABEL: "PDF_TOGGLE_LABEL",
    SEARCHBOX_PLACEHOLDER: "SEARCHBOX_PLACEHOLDER",
    PROGRESSINDICATOR_LABEL: "PROGRESSINDICATOR_LABEL",
    PROGRESSINDICATOR_TEXT_GENERIC: "PROGRESSINDICATOR_TEXT_GENERIC",
    PROGRESSINDICATOR_TEXT_FETCHING: "PROGRESSINDICATOR_TEXT_FETCHING",
    MESSAGEBAR_ERROR_NORECORDSSELCTED: "MESSAGEBAR_ERROR_NORECORDSSELCTED",
    MESSAGEBAR_SUCCESS_DOWNLOAD: "MESSAGEBAR_SUCCESS_DOWNLOAD",
    MESSAGEBAR_SUCCESS_EMAIL: "MESSAGEBAR_SUCCESS_EMAIL",
    MESSAGEBAR_SUCCESS_SHAREPOINT: "MESSAGEBAR_SUCCESS_SHAREPOINT",
    MESSAGEBAR_ERROR_NOTEMPLATES: "MESSAGEBAR_ERROR_NOTEMPLATES"
} as const


export const Entities = {
    Email: "email",
    DocumentTemplates: "documenttemplate",
    PersonalDocumentTemplates: "personaldocumenttemplate",
    Attachment: "activitymimeattachment"
} as const

export const ToggleValue = {
    YES: "0",
    NO: "1"
} as const

export const DocumentTypes = {
    Word: {
        Extension: "docx",
        MimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    },
    Pdf: {
        Extension: "pdf",
        MimeType: "application/pdf"
    }
} as const
export const WebApiConstants = {
    OdataType: "@odata.type",
    OdataTypes: {
        documentTemplates: "Microsoft.Dynamics.CRM.documenttemplate",
        PersonalDocumentTemplates: "Microsoft.Dynamics.CRM.personaldocumenttemplate",
        sharePointDocument: "Microsoft.Dynamics.CRM.sharepointdocument"
    },
    PrimaryAttributeId: {
        documentTemplates: "documenttemplateid",
        PersonalDocumentTemplates: "personaldocumenttemplateid"
    }
} as const