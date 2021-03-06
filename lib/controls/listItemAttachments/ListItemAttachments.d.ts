/// <reference types="react" />
import * as React from 'react';
import { IListItemAttachmentsProps } from '.';
import { IListItemAttachmentsState } from '.';
export declare class ListItemAttachments extends React.Component<IListItemAttachmentsProps, IListItemAttachmentsState> {
    private _spservice;
    private previewImages;
    private _utilities;
    constructor(props: IListItemAttachmentsProps);
    /**
     * componentDidMount lifecycle hook
     */
    componentDidMount(): void;
    private loadAttachmentPreview(file);
    /**
     * Load Item Attachments
     */
    private loadAttachments();
    /**
     * Close the dialog
     */
    private _closeDialog;
    /**
     * Attachment uploaded event handler
     */
    private _onAttachmentUpload;
    /**
     * On delete attachment event handler
     *
     * @param file
     */
    private onDeleteAttachment;
    /**
     * Delete the attachment once it was confirmed
     */
    private onConfirmedDeleteAttachment;
    /**
     * Default React render method
     */
    render(): JSX.Element;
}
