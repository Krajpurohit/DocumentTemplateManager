import * as React from 'react';
import { useState, useMemo } from "react";
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/stack';
import { Icon } from '@fluentui/react/lib/Icon';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { getFileTypeIconProps, initializeFileTypeIcons } from '@fluentui/react-file-type-icons';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { CommandBar, ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import {
    DetailsList,
    DetailsListLayoutMode,
    IDetailsHeaderProps,
    Selection,
    SelectionMode,
    IColumn,
    ConstrainMode,
    IGroup
} from '@fluentui/react/lib/DetailsList';
import { IRenderFunction } from '@fluentui/react/lib/Utilities';
import { ScrollablePane, ScrollbarVisibility } from '@fluentui/react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from '@fluentui/react/lib/Sticky';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { Dialog, DialogType } from '@fluentui/react/lib/Dialog';
import { ITemplateItem, Columns } from './Items';
import { classNames } from './Styles';
import { IInputs } from '../generated/ManifestTypes';
import { CreateActivityMimeAttachmentRequest, ExportPdfDocumentRequest, ExportWordDocumentRequest, RetrieveWordTemplates, UploadDocumentRequest } from './webApiHelper';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DocumentTypes, Entities, ResourceKeys, ToggleValue, WebApiConstants } from './Constants';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

export interface IDocumentTemplateManagerProps {
    primaryEntityName: string,
    primaryEntityId: string,
    pcfContext: ComponentFramework.Context<IInputs>,
    primaryEntityTypeCode: number,
    isSharePointEnabled: boolean,
    primaryEntitySetName: string,
    hasActivities:boolean,
}

export const DocumentTemplateManagerApp: React.FunctionComponent<IDocumentTemplateManagerProps> = (props: IDocumentTemplateManagerProps) => {
    initializeFileTypeIcons();
    const [selectedItems, setSelectedItems] = useState<ITemplateItem[]>();
    const selection = useMemo(
        () =>
            new Selection({
                onSelectionChanged: () => {
                    setSelectedItems(selection.getSelection() as ITemplateItem[]);
                },
                selectionMode: SelectionMode.multiple,
            }),
        []);
    const [hiddenDialog, SetHiddenDialog] = useState(true);
    const [isInProgress, SetIsInProgress] = useState(false);
    const [templates, SetTemplates] = useState<ITemplateItem[]>([]);
    const [allTemplates, SetAllTemplates] = useState<ITemplateItem[]>([]);
    const [groups, SetGroups] = useState<IGroup[]>([]);
    const [progressIndicatorDescription, SetProgressIndicatorDescription] = useState(props.pcfContext.resources.getString(ResourceKeys.PROGRESSINDICATOR_TEXT_GENERIC))
    const [convertToPdf, SetConvirtToPdf] = useState(false);
    const [displayMessageBar, SetDisplayMessageBar] = useState(false);
    const [isError, SetIsError] = useState(false);
    const [messageBarMessage, SetMessageBarMessage] = useState("");
    const FetchTemplates = async () => {
        SetDisplayMessageBar(false);
        let documentTemplates: ITemplateItem[] = [];
        let Systemtemplates: ITemplateItem[] = [];
        let Usertemplates: ITemplateItem[] = [];
        SetHiddenDialog(false);
        SetIsInProgress(true);
        SetProgressIndicatorDescription(props.pcfContext.resources.getString(ResourceKeys.PROGRESSINDICATOR_TEXT_GENERIC));
        Systemtemplates = await RetrieveWordTemplates(props.pcfContext, Entities.DocumentTemplates, props.primaryEntityName, convertToPdf);
        Usertemplates = await RetrieveWordTemplates(props.pcfContext, Entities.PersonalDocumentTemplates, props.primaryEntityName, convertToPdf);
        documentTemplates = [...Systemtemplates, ...Usertemplates];
        SetIsInProgress(false);
        if (documentTemplates.length === 0) {
            SetIsError(true);
            SetMessageBarMessage(props.pcfContext.resources.getString(ResourceKeys.MESSAGEBAR_ERROR_NOTEMPLATES));
            SetDisplayMessageBar(true);
        }
        SetAllTemplates(documentTemplates);
        SetTemplates(documentTemplates);
        groupTemplates(documentTemplates);
    }
    const groupTemplates = (templates: ITemplateItem[]) => {
        let groups: IGroup[] = [];
        let data: [] = templates.reduce(function (r, a) {
            r[a.category] = r[a.category] || [];
            r[a.category].push(a);
            return (r)
        }, Object.create(null));
        let formattedData = [];
        for (let key in data) {
            formattedData.push([key].concat(Object.values(data[key])));
        }
        let previousStartIndex = 0;
        for (let i = 0; i < formattedData.length; i++) {
            previousStartIndex = i === 0 ? 0 : formattedData[i - 1].length - 1 + previousStartIndex;
            let group: IGroup = {
                key: i.toString(),
                name: formattedData[i][0],
                startIndex: i === 0 ? 0 : previousStartIndex,
                count: formattedData[i].length - 1,
                level: 0
            }
            groups.push(group);
        }
        SetGroups(groups);
    }
    const getItems = (): ICommandBarItemProps[] => {
        let ribbonItems: ICommandBarItemProps[] = [];
        if (props.pcfContext.utils.hasEntityPrivilege(Entities.DocumentTemplates, 2, 1) && props.pcfContext.parameters.allowDownload.raw === ToggleValue.YES) {
            ribbonItems.push({
                key: 'download',
                text: props.pcfContext.resources.getString(ResourceKeys.DOWNLOAD_BUTTON_LABEL),
                iconProps: { iconName: 'Download' },
                onClick: () => {
                    SetProgressIndicatorDescription(props.pcfContext.resources.getString(ResourceKeys.DOWNLOAD_DESCRIPTION))
                    TemplateAction('download')
                }

            })
        }
        if (props.hasActivities && props.pcfContext.utils.hasEntityPrivilege(Entities.Email, 1, 1) && props.pcfContext.parameters.allowEmail.raw === ToggleValue.YES) {
            ribbonItems.push(
                {
                    key: 'email',
                    text: props.pcfContext.resources.getString(ResourceKeys.EMAIL_BUTTON_LABEL),
                    iconProps: { iconName: 'Mail' },
                    onClick: () => {
                        SetProgressIndicatorDescription(props.pcfContext.resources.getString(ResourceKeys.EMAIL_DESCRIPTION))
                        TemplateAction('email')
                    }
                }
            )
        }
        if (props.pcfContext.parameters.allowSaveToSharePoint.raw === ToggleValue.YES && props.isSharePointEnabled) {
            ribbonItems.push(
                {
                    key: 'saveToSharepoint',
                    text: props.pcfContext.resources.getString(ResourceKeys.SAVETOSHAREPOINT_BUTTON_LABEL),
                    iconProps: { iconName: 'SharepointLogo' },
                    onClick: () => {
                        SetProgressIndicatorDescription(props.pcfContext.resources.getString(ResourceKeys.SAVETOSHAREPOINT_DESCRIPTION))
                        TemplateAction('saveToSharePoint')
                    }
                });
        }
        return ribbonItems;
    };
    const TemplateAction = (actionaName: string) => {
        if (selectedItems && selectedItems.length > 0) {
            SetDisplayMessageBar(false);
            SetIsInProgress(true);
            let requests: any = [];
            for (let selectedTemplate of selectedItems) {
                let templateData: any = {};
                if (selectedTemplate.category === "System Templates") {
                    templateData[WebApiConstants.OdataType] = WebApiConstants.OdataTypes.documentTemplates;
                    templateData[WebApiConstants.PrimaryAttributeId.documentTemplates] = selectedTemplate.documentTemplateId
                }
                else {
                    templateData[WebApiConstants.OdataType] = WebApiConstants.OdataTypes.PersonalDocumentTemplates;
                    templateData[WebApiConstants.PrimaryAttributeId.PersonalDocumentTemplates] = selectedTemplate.documentTemplateId
                }
                let request = convertToPdf ? new ExportPdfDocumentRequest(props.primaryEntityTypeCode, `[\"{${props.primaryEntityId}}"\]`, templateData) : new ExportWordDocumentRequest(props.primaryEntityTypeCode, `[\"{${props.primaryEntityId}}"\]`, templateData);
                requests.push(request);
            }
            //@ts-ignore
            props.pcfContext.webAPI.executeMultiple(requests).then(function (responses) {
                if (responses.length > 0) {
                    switch (actionaName) {
                        case 'download':
                            DownloadTemplates(responses);
                            break;
                        case 'email':
                            EmailTemplates(responses);
                            break;
                        case 'saveToSharePoint':
                            SaveToSharePoint(responses);
                            break;
                        default:
                            return;
                    }
                }
            });
        }
        else {
            SetIsInProgress(false);
            SetIsError(true);
            SetMessageBarMessage(props.pcfContext.resources.getString(ResourceKeys.MESSAGEBAR_ERROR_NORECORDSSELCTED));
            SetDisplayMessageBar(true);

        }
    }

    const DownloadTemplates = (responses: [any]) => {
        for (let i = 0; i < responses.length; i++) {
            let file: ComponentFramework.FileObject = {} as ComponentFramework.FileObject;
            let fileOption: ComponentFramework.NavigationApi.OpenFileOptions = {} as ComponentFramework.NavigationApi.OpenFileOptions;
            responses[i].json().then((response: any) => {
                file.fileContent = convertToPdf ? response.PdfFile : response.WordFile;
                file.fileName = GenerateFileName(selectedItems ? selectedItems[i].templateName : "", selectedItems ? selectedItems[i].fileTypeExtension : "");
                file.mimeType = convertToPdf ? DocumentTypes.Pdf.MimeType : DocumentTypes.Word.MimeType;
                fileOption.openMode = 2;
                props.pcfContext.navigation.openFile(file, fileOption);
                SetIsInProgress(false);
                SetIsError(false);
                SetMessageBarMessage(props.pcfContext.resources.getString(ResourceKeys.MESSAGEBAR_SUCCESS_DOWNLOAD));
                SetDisplayMessageBar(true);
            });
        }
    }

    const GenerateFileName = (templateNam: string, ext: string): string => {
        var now: any = new Date
            , datePattern = props.pcfContext.userSettings.dateFormattingInfo.shortDatePattern
            , timePattern = props.pcfContext.userSettings.dateFormattingInfo.longDatePattern
            , dateSeparator = props.pcfContext.userSettings.dateFormattingInfo.dateSeparator
            , timeSeparator = props.pcfContext.userSettings.dateFormattingInfo.timeSeparator
            , fileName = templateNam + " " + now.format(datePattern) + " " + now.format(timePattern);
        fileName = fileName.split(dateSeparator).join("-").split(timeSeparator).join("-");
        return "" + fileName + "." + ext;
    }
    const EmailTemplates = async (responses: [any]) => {
        let emailId = await CreateEmail();
        let requests: any = [];
        for (let i = 0; i < responses.length; i++) {
            responses[i].json().then((response: any) => {
                let attachMentPayload: any = {};
                attachMentPayload["body"] = convertToPdf ? response.PdfFile : response.WordFile;
                attachMentPayload["objectid_activitypointer@odata.bind"] = `activitypointers(${emailId})`;
                attachMentPayload["objecttypecode"] = Entities.Email;
                attachMentPayload["filename"] = `${selectedItems ? selectedItems[i].templateName : ""}.${selectedItems ? selectedItems[i].fileTypeExtension : ""}`;
                let request = new CreateActivityMimeAttachmentRequest(Entities.Attachment, attachMentPayload);
                requests.push(request);
                if (i + 1 === responses.length) {
                    //@ts-ignore
                    props.pcfContext.webAPI.executeMultiple(requests).then((data) => {
                        SetIsInProgress(false);
                        SetIsError(false);
                        SetMessageBarMessage(props.pcfContext.resources.getString(ResourceKeys.MESSAGEBAR_SUCCESS_EMAIL));
                        SetDisplayMessageBar(true);
                        let options: ComponentFramework.NavigationApi.EntityFormOptions = {} as ComponentFramework.NavigationApi.EntityFormOptions;
                        options.entityId = emailId;
                        options.entityName = Entities.Email;
                        options.openInNewWindow = true;
                        props.pcfContext.navigation.openForm(options);

                    });
                }
            });

        }
    }
    const CreateEmail = async () => {
        let email: any = {};
        let customer: any = await FetchCustomer();
        email.email_activity_parties = [];
        email[`regardingobjectid_${props.primaryEntityName}@odata.bind`] = `/${props.primaryEntitySetName}(${props.primaryEntityId})`;
        email["subject"] = props.pcfContext.mode.contextInfo.entityRecordName;
        email["email_activity_parties"] = [{
            "partyid_systemuser@odata.bind": `/systemusers(${props.pcfContext.userSettings.userId.replace('{', "").replace('}', "")})`,
            "participationtypemask": 1   ///From Email
        }];
        let to: any = {};
        if (customer !== null) {
            let customerSetName = await FetchMetadata(customer.enityCustomerIdType);
            to[`partyid_${customer.enityCustomerIdType}@odata.bind`] = `/${customerSetName._entitySetName}(${customer.entityCustomerId})`;
            to["participationtypemask"] = 2;
            email["email_activity_parties"].push(to);
        }
        let id = await props.pcfContext.webAPI.createRecord(Entities.Email, email).then((records: ComponentFramework.LookupValue) => {
            return records.id;
        });
        return id;

    }

    const FetchMetadata = async (logicalName: string) => {
        let metadata = await props.pcfContext.utils.getEntityMetadata(logicalName, []).then((metadata: ComponentFramework.PropertyHelper.EntityMetadata) => { return metadata });
        return metadata;
    }

    const FetchCustomer = async () => {
        let customer = await props.pcfContext.webAPI.retrieveRecord(props.primaryEntityName, props.primaryEntityId, "?$select=_customerid_value").then((entity: ComponentFramework.WebApi.Entity) => {
            let record: any = {};
            //@ts-ignore
            if (!props.pcfContext.utils.isNullOrUndefined(record)) {
                record.entityCustomerId = entity["_customerid_value"];
                record.enityCustomerIdType = entity["_customerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
            }
            else {
                record = null;
            }
            return record;
        },(reason)=>{
            return null;
        });
        return customer;
    }
    const SaveToSharePoint = (responses: [any]) => {
        let requests: any = [];
        for (let i = 0; i < responses.length; i++) {
            responses[i].json().then((response: any) => {
                let entity: any = {};
                entity[WebApiConstants.OdataType] = WebApiConstants.OdataTypes.sharePointDocument;
                entity["locationid"] = "";
                entity["title"] = `${selectedItems ? selectedItems[i].templateName : ""}.${selectedItems ? selectedItems[i].fileTypeExtension : ""}`
                let parentEntityRef: any = {};
                parentEntityRef[WebApiConstants.OdataType] = `Microsoft.Dynamics.CRM.${props.primaryEntityName}`;
                parentEntityRef[`${props.primaryEntityName}id`] = props.primaryEntityId;
                let request = new UploadDocumentRequest(convertToPdf ? response.PdfFile : response.WordFile, entity, true, parentEntityRef, "");
                requests.push(request);
                if (i + 1 === responses.length) {
                    //@ts-ignore
                    props.pcfContext.webAPI.executeMultiple(requests);
                    SetIsInProgress(false);
                    SetIsError(false);
                    SetMessageBarMessage(props.pcfContext.resources.getString(ResourceKeys.MESSAGEBAR_SUCCESS_SHAREPOINT));
                    SetDisplayMessageBar(true);
                }
            });
        }

    }
    const onRenderDetailsHeader = (
        props?: IDetailsHeaderProps,
        defaultRender?: IRenderFunction<IDetailsHeaderProps>
    ): JSX.Element => {
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
                {defaultRender && defaultRender({ ...props! })}
            </Sticky>
        );
    }
    const renderItemColumn = (template: ITemplateItem, index?: number, column?: IColumn) => {
        if (column) {
            const fieldContent = template[column.key as keyof ITemplateItem] as string;

            switch (column.key) {
                case 'fileTypeExtension':
                    return <Icon {...getFileTypeIconProps({ extension: fieldContent })}></Icon>;
                case 'templateName':
                case 'description':
                    return <div>{fieldContent}</div>;
                default:
                    break;
            }
        }
    }
    const onFilterChanged = (ev?: React.ChangeEvent<HTMLInputElement>, text?: string): void => {
        let filteredTemplates: ITemplateItem[] = [];
        if (text) {
            filteredTemplates = allTemplates.filter((item: ITemplateItem) => hasText(item, text));
        }
        else {
            filteredTemplates = allTemplates;
        }
        SetTemplates(filteredTemplates);
        groupTemplates(filteredTemplates);
    };
    const hasText = (item: ITemplateItem, text: string): boolean => {
        return `${item.templateName.toLowerCase()}|${item.description === null ? item.description : item.description.toLowerCase()}`.indexOf(text.toLowerCase()) > -1;
    }
    const OnChangeCovertToPDF = (ev: React.MouseEvent<HTMLElement, MouseEvent>, checked?: boolean) => {
        if (checked) {
            SetTemplates(allTemplates.map(template => { template.fileTypeExtension = DocumentTypes.Pdf.Extension; return template }));
            SetConvirtToPdf(true);
        }
        else {
            SetTemplates(allTemplates.map(template => { template.fileTypeExtension = DocumentTypes.Word.Extension; return template }));
            SetConvirtToPdf(false);
        }
    }
    return (
        <div>
            <DefaultButton
                className={classNames.defaultButton}
                name={props.pcfContext.resources.getString(ResourceKeys.DOWNLOAD_DEFAULTBUTTON_LABEL)}
                iconProps={{ ...getFileTypeIconProps({ extension: DocumentTypes.Word.Extension }) }}
                ariaLabel={props.pcfContext.resources.getString(ResourceKeys.DOWNLOAD_DEFAULTBUTTON_LABEL)}
                onClick={FetchTemplates}
                text={props.pcfContext.resources.getString(ResourceKeys.DOWNLOAD_DEFAULTBUTTON_TEXT)}
                disabled={props.pcfContext.mode.isControlDisabled && props.pcfContext.parameters.enableForInactiveRecords.raw === "1"} />
            <Dialog
                hidden={hiddenDialog}
                onDismiss={() => { SetHiddenDialog(true); }}
                dialogContentProps={{
                    type: DialogType.close,
                    title: props.pcfContext.resources.getString(ResourceKeys.DIALOG_TITLE)
                }}
                modalProps={{ isBlocking: false }}
                minWidth='900px'>
                <CommandBar items={getItems()} styles={{ root: { padding: 0, margin: 0 } }} />
                <div className={classNames.wrapper}>
                    <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                        <Sticky stickyPosition={StickyPositionType.Header}>
                            <Stack horizontal tokens={{ childrenGap: 20, padding: 10 }}>
                                <Stack.Item grow={1} className={classNames.stackItemStyles}>
                                    <Toggle label={props.pcfContext.resources.getString(ResourceKeys.PDF_TOGGLE_LABEL)} inlineLabel checked={convertToPdf} onChange={OnChangeCovertToPDF} />
                                </Stack.Item>
                                <Stack.Item grow={3} className={classNames.stackItemStyles}>
                                    <SearchBox className={classNames.searchBox} placeholder={props.pcfContext.resources.getString(ResourceKeys.SEARCHBOX_PLACEHOLDER)} onChange={onFilterChanged} />
                                </Stack.Item>
                            </Stack>
                            <Stack>
                                {isInProgress && <ProgressIndicator label={props.pcfContext.resources.getString(ResourceKeys.PROGRESSINDICATOR_LABEL)} description={progressIndicatorDescription} />}
                            </Stack>
                            <Stack>
                                {displayMessageBar && <MessageBar messageBarType={isError ? MessageBarType.error : MessageBarType.success} isMultiline={false} >{messageBarMessage}</MessageBar>}
                            </Stack>
                        </Sticky>
                        <MarqueeSelection selection={selection}>
                            <DetailsList
                                groups={groups}
                                items={templates}
                                columns={Columns}
                                setKey="set"
                                layoutMode={DetailsListLayoutMode.fixedColumns}
                                constrainMode={ConstrainMode.unconstrained}
                                onRenderItemColumn={renderItemColumn}
                                onRenderDetailsHeader={onRenderDetailsHeader}
                                selection={selection}
                                selectionPreservedOnEmptyClick={true}
                            />
                        </MarqueeSelection>
                    </ScrollablePane>
                </div>
            </Dialog>
        </div>
    );
}
