"use strict";
/// <reference types="sinon" />
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var chai_1 = require("chai");
var enzyme_1 = require("enzyme");
var FileTypeIcon_1 = require("./FileTypeIcon");
var IFileTypeIcon_1 = require("./IFileTypeIcon");
describe('<FileTypeIcon />', function () {
    var fileTypeIcon;
    afterEach(function () {
        fileTypeIcon.unmount();
    });
    /**
     * Font icon tests
     */
    it('Font icon test with application', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.font, application: IFileTypeIcon_1.ApplicationType.SASS }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('i.ms-Icon--FileSass')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('div')).to.have.length(0);
        done();
    });
    it('Font icon test with path', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.font, path: "https://contoso.sharepoint.com/documents/filename.doc" }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('i.ms-Icon--WordDocument')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('div')).to.have.length(0);
        done();
    });
    it('Font icon test with path that contains querystring params', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.font, path: "https://contoso.sharepoint.com/documents/filename.vsd?param1=prop1&param2=prop2" }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('i.ms-Icon--VisioDocument')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('div')).to.have.length(0);
        done();
    });
    it('Font icon test with path of unknown extension', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.font, path: "https://contoso.sharepoint.com/documents/filename.unknown" }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('i.ms-Icon--Page')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('div')).to.have.length(0);
        done();
    });
    it('Font icon test without application or path, should render the generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.font }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('i.ms-Icon--Page')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('div')).to.have.length(0);
        done();
    });
    it('Font icon test with both the application and path, should take the path into account', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.font, application: IFileTypeIcon_1.ApplicationType.SASS, path: "https://contoso.sharepoint.com/documents/filename.doc" }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('i.ms-Icon--WordDocument')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('div')).to.have.length(0);
        done();
    });
    /**
     * Image icon tests
     */
    it('Image icon test with application', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, application: IFileTypeIcon_1.ApplicationType.Word }));
        // Check if "i" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--docx')).to.have.length(1);
        // Check if no "div" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon test with path', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.doc" }));
        // Check if "div" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--docx')).to.have.length(1);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon test with path that contains querystring params', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.vsd?param1=prop1&param2=prop2" }));
        // Check if "div" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--vsdx')).to.have.length(1);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon test with path of unknown extension', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.unknown" }));
        // Check if "img" element exists
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        // Check if it did not create a brand icon element
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(0);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon test without application or path, should render the generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image }));
        // Check if "img" element exists
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        // Check if it did not create a brand icon element
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(0);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon test without application or path, should render the generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, application: IFileTypeIcon_1.ApplicationType.SASS, path: "https://contoso.sharepoint.com/documents/filename.docx" }));
        // Check if "div" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--docx')).to.have.length(1);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size SMALL test', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, application: IFileTypeIcon_1.ApplicationType.Word, size: IFileTypeIcon_1.ImageSize.small }));
        // Check if "div" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(1);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size MEDIUM test', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, application: IFileTypeIcon_1.ApplicationType.Word, size: IFileTypeIcon_1.ImageSize.medium }));
        // Check if "div" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon48')).to.have.length(1);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size LARGE test', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, application: IFileTypeIcon_1.ApplicationType.Word, size: IFileTypeIcon_1.ImageSize.large }));
        // Check if "div" element exists
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon96')).to.have.length(1);
        // Check if no "i" element has been rendered
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size SMALL test for generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, size: IFileTypeIcon_1.ImageSize.small }));
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(0);
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size MEDIUM test for generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, size: IFileTypeIcon_1.ImageSize.medium }));
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon48')).to.have.length(0);
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size LARGE test for generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, size: IFileTypeIcon_1.ImageSize.large }));
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon96')).to.have.length(0);
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size test with unkown size', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, size: 8, application: IFileTypeIcon_1.ApplicationType.Word }));
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(1);
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon size test with unkown size for generic icon', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, size: 8 }));
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(0);
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
    it('Image icon test with unkown application', function (done) {
        fileTypeIcon = enzyme_1.mount(React.createElement(FileTypeIcon_1.FileTypeIcon, { type: IFileTypeIcon_1.IconType.image, application: 90 }));
        chai_1.expect(fileTypeIcon.find('div img')).to.have.length(1);
        chai_1.expect(fileTypeIcon.find('div.ms-BrandIcon--icon16')).to.have.length(0);
        chai_1.expect(fileTypeIcon.find('i')).to.have.length(0);
        done();
    });
});

//# sourceMappingURL=FileTypeIcon.test.js.map
