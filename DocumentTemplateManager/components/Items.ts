import { IColumn } from '@fluentui/react/lib/DetailsList';

export interface ITemplateItem {
  key: number | string,
  documentTemplateId: string,
  templateName: string,
  description: string,
  fileTypeExtension: string,
  category: string
}
export const Columns: IColumn[] = [
  {
    key: 'fileTypeExtension',
    name: '',
    fieldName: 'fileTypeExtension',
    minWidth: 20,
    maxWidth: 40,
    isResizable: false
  },
  {
    key: 'templateName',
    name: 'Template Name',
    fieldName: 'templateName',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'description',
    name: 'Description',
    fieldName: 'description',
    minWidth: 150,
    maxWidth: 250,
    isResizable: true
  },
]