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
import { CreateActivityMimeAttachmentRequest, ExportWordDocumentRequest, UploadDocumentRequest } from './webApiHelper';

export interface IDocumentTemplateManagerProps {
    primaryEntityName: string,
    primaryEntityId: string,
    pcfContext: ComponentFramework.Context<IInputs>,
    primaryEntityTypeCode: number,
    isSharePointEnabled:boolean,
    primaryEntitySetName:string
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
    const [progressIndicatorDescription, SetProgressIndicatorDescription] = useState("Please wait...")
    const FetchTemplates = (): void => {
        let documentTemplates: ITemplateItem[] = [];
        SetHiddenDialog(false);
        SetIsInProgress(true);
        SetProgressIndicatorDescription("Fetching document templates from dataverse...");
        props.pcfContext.webAPI.retrieveMultipleRecords('documenttemplate', `?$select=clientdata,description,documenttemplateid,documenttype,name&$filter=associatedentitytypecode eq '${props.primaryEntityName}' and documenttype eq 2`).then(
            (response: ComponentFramework.WebApi.RetrieveMultipleResponse) => {
                for (let i = 0; i < response.entities.length; i++) {
                    let template: ITemplateItem = {
                        templateName: response.entities[i].name,
                        documentTemplateId: response.entities[i].documenttemplateid,
                        key: i,
                        description: response.entities[i].description,
                        fileTypeExtension: "docx",
                        category: "System Templates"
                    }
                    documentTemplates.push(template);
                }
                props.pcfContext.webAPI.retrieveMultipleRecords('personaldocumenttemplate', `?$select=clientdata,description,personaldocumenttemplateid,documenttype,name&$filter=associatedentitytypecode eq '${props.primaryEntityName}' and documenttype eq 2`).then(
                    (response: ComponentFramework.WebApi.RetrieveMultipleResponse) => {
                        for (let i = 0; i < response.entities.length; i++) {
                            let template: ITemplateItem = {
                                templateName: response.entities[i].name,
                                documentTemplateId: response.entities[i].personaldocumenttemplateid,
                                key: i,
                                description: response.entities[i].description,
                                fileTypeExtension: "docx",
                                category: "User Templates"
                            }
                            documentTemplates.push(template);
                        }
                        SetIsInProgress(false);
                        SetAllTemplates(documentTemplates);
                        SetTemplates(documentTemplates);
                        groupTemplates(documentTemplates);
                    });
            });

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
        let ribbonItems:ICommandBarItemProps[]=[];
        if(props.pcfContext.utils.hasEntityPrivilege("documenttemplate",2,1) && props.pcfContext.parameters.allowDownload.raw==="0"){
            ribbonItems.push({
                key: 'download',
                text: 'Download',
                iconProps: { iconName: 'Download' },
                onClick: () => {
                    SetProgressIndicatorDescription("Downloading selected document templates...")
                    TemplateAction('download')
                }
                
            })
        }
        if(props.pcfContext.utils.hasEntityPrivilege("email",1,1)&& props.pcfContext.parameters.allowEmail.raw==="0"){
            ribbonItems.push(
                {
                    key: 'email',
                    text: 'Email',
                    iconProps: { iconName: 'Mail' },
                    onClick: () => {
                        SetProgressIndicatorDescription("Creating email and attaching selected document templates...")
                        TemplateAction('email')
                    }
                }
            )
        }
        if(props.pcfContext.parameters.allowSaveToSharePoint.raw==="0" && props.isSharePointEnabled){
            ribbonItems.push(
            {
                key: 'saveToSharepoint',
                text: 'Save To SharePoint',
                iconProps: { iconName: 'SharepointLogo' },
                onClick: () => {
                    SetProgressIndicatorDescription("Saving selected document templates into SharePoint...")
                    TemplateAction('saveToSharePoint')
                }
            });
        }
        return ribbonItems;
    };
    const TemplateAction = (actionaName: string) => {
        if (selectedItems && selectedItems.length > 0) {
            SetIsInProgress(true);
            let requests: any = [];
            for (let selectedTemplate of selectedItems) {
                let templateData: any = {};
                if (selectedTemplate.category === "System Templates") {
                    templateData["@odata.type"] = "Microsoft.Dynamics.CRM.documenttemplate";
                    templateData["documenttemplateid"] = selectedTemplate.documentTemplateId
                }
                else {
                    templateData["@odata.type"] = "Microsoft.Dynamics.CRM.personaldocumenttemplate";
                    templateData["personaldocumenttemplateid"] = selectedTemplate.documentTemplateId
                }
                let request = new ExportWordDocumentRequest(props.primaryEntityTypeCode, `[\"{${props.primaryEntityId}}"\]`, templateData);
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
        }
    }

    const DownloadTemplates = (responses: [any]) => {
        for (let i = 0; i < responses.length; i++) {
            let file: ComponentFramework.FileObject = {} as ComponentFramework.FileObject;
            let fileOption: ComponentFramework.NavigationApi.OpenFileOptions = {} as ComponentFramework.NavigationApi.OpenFileOptions;
            responses[i].json().then((response: { WordFile: any; }) => {
                file.fileContent = response.WordFile;
                file.fileName = GenerateFileName(selectedItems ? selectedItems[i].templateName : "");
                file.mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                fileOption.openMode = 2;
                props.pcfContext.navigation.openFile(file, fileOption);
                SetIsInProgress(false);
            });
        }
    }

    const GenerateFileName = (templateNam: string): string => {
        var now: any = new Date
            , datePattern = props.pcfContext.userSettings.dateFormattingInfo.shortDatePattern
            , timePattern = props.pcfContext.userSettings.dateFormattingInfo.longDatePattern
            , dateSeparator = props.pcfContext.userSettings.dateFormattingInfo.dateSeparator
            , timeSeparator = props.pcfContext.userSettings.dateFormattingInfo.timeSeparator
            , fileName = templateNam + " " + now.format(datePattern) + " " + now.format(timePattern);
        fileName = fileName.split(dateSeparator).join("-").split(timeSeparator).join("-");
        return "" + fileName + ".docx";
    }
    const EmailTemplates = async (responses: [any]) => {
        let emailId = await CreateEmail();
        let requests: any = [];
        for (let i = 0; i < responses.length; i++) {
            responses[i].json().then((response: { WordFile: any; }) => {
                let attachMentPayload: any = {};
                attachMentPayload["body"] = response.WordFile;
                attachMentPayload["objectid_activitypointer@odata.bind"] = `activitypointers(${emailId})`;
                attachMentPayload["objecttypecode"] = "email";
                attachMentPayload["filename"] = `${selectedItems ? selectedItems[i].templateName : ""}.${selectedItems ? selectedItems[i].fileTypeExtension : ""}`;
                let request = new CreateActivityMimeAttachmentRequest("activitymimeattachment", attachMentPayload);
                requests.push(request);
                if (i + 1 === responses.length) {
                    //@ts-ignore
                    props.pcfContext.webAPI.executeMultiple(requests).then((data) => {
                        SetIsInProgress(false);
                        let options: ComponentFramework.NavigationApi.EntityFormOptions = {} as ComponentFramework.NavigationApi.EntityFormOptions;
                        options.entityId = emailId;
                        options.entityName = "email";
                        options.openInNewWindow = true;
                        props.pcfContext.navigation.openForm(options);

                    });
                }
            });

        }
    }
    const CreateEmail = async () => {
        let email: any = {};
        email[`regardingobjectid_${props.primaryEntityName}@odata.bind`] = `/${props.primaryEntitySetName}(${props.primaryEntityId})`;
        email["subject"] = "Test Subject";
        email["email_activity_parties"] = [{
            "partyid_systemuser@odata.bind": `/systemusers(${props.pcfContext.userSettings.userId.replace('{', "").replace('}', "")})`,
            "participationtypemask": 1   ///From Email
        }]
        let id = await props.pcfContext.webAPI.createRecord("email", email).then((records: ComponentFramework.LookupValue) => {
            return records.id;
        });
        return id;

    }
    const SaveToSharePoint = (responses: [any]) => {
        let requests: any = [];
        for (let i = 0; i < responses.length; i++) {
            responses[i].json().then((response: { WordFile: any; }) => {
                let entity: any = {};
                entity["@odata.type"] = "Microsoft.Dynamics.CRM.sharepointdocument";
                entity["locationid"] = "";
                entity["title"] = `${selectedItems ? selectedItems[i].templateName : ""}.${selectedItems ? selectedItems[i].fileTypeExtension : ""}`
                let parentEntityRef: any = {};
                parentEntityRef["@odata.type"] = `Microsoft.Dynamics.CRM.${props.primaryEntityName}`;
                parentEntityRef[`${props.primaryEntityName}id`] = props.primaryEntityId;
                let request = new UploadDocumentRequest(response.WordFile, entity, true, parentEntityRef, "");
                requests.push(request);
                if (i + 1 === responses.length) {
                    //@ts-ignore
                    props.pcfContext.webAPI.executeMultiple(requests);
                    SetIsInProgress(false);
                }
            });
        }

    }
    const onItemInvoked = (template: ITemplateItem): void => {
        console.log('Item invoked: ' + template.templateName);
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
    return (
        <div>
            <DefaultButton
                className={classNames.defaultButton}
                name='Download Document Templates'
                iconProps={{ ...getFileTypeIconProps({ extension: 'docx' }) }}
                ariaLabel='Download Document Templates'
                onClick={FetchTemplates}
                text='Export to Word'
                disabled={props.pcfContext.mode.isControlDisabled && props.pcfContext.parameters.enableForInactiveRecords.raw==="1"} />
            <Dialog
                hidden={hiddenDialog}
                onDismiss={() => { SetHiddenDialog(true); }}
                dialogContentProps={{
                    type: DialogType.close,
                    title: 'Export Document Templates'
                }}
                modalProps={{ isBlocking: false }}
                minWidth='900px'>
                <CommandBar items={getItems()} className={classNames.commandbar} />
                <div className={classNames.wrapper}>
                    <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                        <Sticky stickyPosition={StickyPositionType.Header}>
                            <Stack horizontal tokens={{ childrenGap: 20, padding: 10 }}>
                                <Stack.Item grow align="stretch">
                                    <SearchBox className={classNames.searchBox} placeholder="Search Templates" onChange={onFilterChanged} />
                                </Stack.Item>
                            </Stack>
                            <Stack>
                                {isInProgress && <ProgressIndicator label="In progress" description={progressIndicatorDescription} />}
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
                                ariaLabelForSelectionColumn="Toggle selection"
                                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                onItemInvoked={onItemInvoked}
                            />
                        </MarqueeSelection>
                    </ScrollablePane>
                </div>
            </Dialog>
        </div>
    );
}
