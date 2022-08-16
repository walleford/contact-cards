import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneDropdown,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CallingCardsWebPartStrings';
import CallingCards from './components/CallingCards';
import { ICallingCardsProps } from './components/ICallingCardsProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { spfi, SPFx } from '@pnp/sp';
import {
    FilePicker,
    IFilePickerProps
} from '@pnp/spfx-controls-react/lib/FilePicker';
import {
    PropertyFieldFilePicker,
    IPropertyFieldFilePickerProps,
    IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import '@pnp/sp/files';
import '@pnp/sp/webs';
import '@pnp/sp/folders';


export interface ICallingCardsWebPartProps {
    description: string;
    CallingCards: any[];
    filePickerResult: IFilePickerResult;
    Layout: string;
}


export default class CallingCardsWebPart extends BaseClientSideWebPart<ICallingCardsWebPartProps> {
    private sp = spfi().using(SPFx(this.context));
    


    public render(): void {
        const element: React.ReactElement<ICallingCardsProps> = React.createElement(
            CallingCards,
            {
                description: this.properties.description,
                spfxContext: this.context,
                CallingCards: this.properties.CallingCards,
                Layout: this.properties.Layout
            }
        );

        ReactDom.render(element, this.domElement);
    }
    
    private options = [
        { key: 'vertical', text: 'Vertical' }, { key: 'horizontal', text: 'Horizontal' }
    ]
    private async uploadFiles(fileContent, fileName, FolderPath) {
        try {
            if (fileContent.size <= 10485760) {
                // small upload
                let result = await this.sp.web.getFolderByServerRelativePath(FolderPath).files.addUsingPath(FolderPath, fileContent.type, { Overwrite: true });
                console.log(result);
            } else {
                // large upload
                let result = await this.sp.web.getFolderByServerRelativePath(FolderPath).files.addChunked(FolderPath, fileContent.type, data => {
                    console.log(`progress`);
                }, true);
            }
            
        }
        catch (err) {
            // (error handling removed for simplicity)
            return Promise.resolve(false);
        }
    }

    protected async onPropertyPaneFieldChanged(
        propertyPath: string,
        oldValue: any,
        newValue: any
    ): Promise<void> {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        this.context.propertyPane.refresh();
        this.render();
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneDropdown('Layout', {
                                    label: 'Vertical or Horizontal Layout?',
                                    options: this.options
                                }),
                                PropertyFieldCollectionData('CallingCards', {
                                    key: 'CallingCards',
                                    label: 'Contact Card Information',
                                    panelHeader: 'Contact Card Panel',
                                    manageBtnLabel: 'Manage Contact Cards',
                                    value: this.properties.CallingCards,
                                    fields: [
                                        {
                                            id: 'Name',
                                            title: 'Contact Name',
                                            type: CustomCollectionFieldType.string,
                                            required: true
                                        },
                                        {
                                            id: 'Position',
                                            title: 'Contact Position',
                                            type: CustomCollectionFieldType.string,
                                            required: true
                                        },
                                        {
                                            id: 'Email',
                                            title: 'Contact Email',
                                            type: CustomCollectionFieldType.string,
                                            required: true
                                        },
                                        {
                                            id: 'PhoneNumber',
                                            title: 'Phone Number 1 (Type Label as well!)',
                                            type: CustomCollectionFieldType.string,
                                            required: false
                                        },
                                        {
                                            id: 'dsn',
                                            title: 'Phone Number 2 (Type Label as well!)',
                                            type: CustomCollectionFieldType.string,
                                            required: false},
                                        {
                                            id: 'duty',
                                            title: 'Phone Number 3 (Type Label as well!)',
                                            type: CustomCollectionFieldType.string,
                                            required: false},
                                        {
                                            id: 'Branch',
                                            title: 'Military Branch (If needed)',
                                            type: CustomCollectionFieldType.string,
                                            required: false},
                                        {
                                            id: 'bioLink',
                                            title: 'Link to Bio',
                                            type: CustomCollectionFieldType.string,
                                            required: false
                                        },
                                        {
                                            id: "filePicker",
                                            title: "Select File",
                                            type: CustomCollectionFieldType.custom,
                                            onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                                                return (
                                                    React.createElement(FilePicker, {
                                                        key: itemId,
                                                        context: this.context,
                                                        buttonLabel: "Select File",
                                                        onChange: (filePickerResult: IFilePickerResult[]) => {
                                                            console.log('changing....', field);
                                                            onUpdate(field.id, filePickerResult[0]);
                                                            
                                                            this.context.propertyPane.refresh();
                                                            this.render();
                                                        },
                                                        onSave:
                                                            async (filePickerResult: IFilePickerResult[]) => {
                                                                for (const filePicked of filePickerResult) {
                                                                    if (filePicked.fileAbsoluteUrl == null) {
                                                                        filePicked.downloadFileContent().then(async r => {
                                                                            let fileresult = await this.sp.web.getFolderByServerRelativePath(`${this.context.pageContext.site.serverRelativeUrl}/SiteAssets/SitePages`).files.addChunked(filePicked.fileName, r);
                                                                            this.properties.CallingCards[0].filePickerResult = filePicked;
                                                                            this.properties.CallingCards[0].filePickerResult.fileAbsoluteUrl = `${this.context.pageContext.site.absoluteUrl}/SiteAssets/SitePages/${fileresult.data.Name}`;
                                                                            this.context.propertyPane.refresh();
                                                                            this.render();
                                                                        });
                                                                    } else {
                                                                        console.log('saving....', filePicked);
                                                                        onUpdate(field.id, filePicked);
                                                                        this.context.propertyPane.refresh();
                                                                        this.render();
                                                                    }
                                                                }
                                                            },
                                                        hideLocalUploadTab: false,
                                                        hideLocalMultipleUploadTab: true,
                                                        hideLinkUploadTab: false,
                                                    })
                                                );
                                            },
                                            required: true
                                        },
                                    ],
                                    disabled: false,
                                }),
                            ]
                        }
                    ]
                }
            ]
        };
    }
}

