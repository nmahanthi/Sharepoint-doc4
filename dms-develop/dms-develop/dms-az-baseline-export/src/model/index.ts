export interface IBaselineItem {
    itemZipPath: string;
    itemServerRelativeUrl: string;
    siteUrl: string,
    id: string,
    type: 'docSet' | 'folder' | 'file',
    metadata?: { [key: string]: unknown },
    error?: string
}