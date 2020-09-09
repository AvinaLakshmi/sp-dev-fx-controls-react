import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
import { IControlsTestWebPartProps } from './IControlsTestWebPartProps';
/**
 * Web part to test the React controls
 */
export default class ControlsTestWebPart extends BaseClientSideWebPart<IControlsTestWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
