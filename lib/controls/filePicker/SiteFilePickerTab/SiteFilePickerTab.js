"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var controls_1 = require("../controls");
var Button_1 = require("office-ui-fabric-react/lib/components/Button");
var Breadcrumb_1 = require("office-ui-fabric-react/lib/Breadcrumb");
var Link_1 = require("office-ui-fabric-react/lib/Link");
var SiteFilePickerTab_module_scss_1 = require("./SiteFilePickerTab.module.scss");
var strings = require("ControlStrings");
var SiteFilePickerTab = (function (_super) {
    __extends(SiteFilePickerTab, _super);
    function SiteFilePickerTab(props) {
        var _this = _super.call(this, props) || this;
        _this.renderBreadcrumbItem = function (item) {
            return (React.createElement(Link_1.Link, { href: item.href, onClick: item.onClick, key: item.key, className: "ms-Link ms-Breadcrumb-itemLink " + SiteFilePickerTab_module_scss_1.default.breadcrumbNavItem }, item.text));
        };
        /**
         * Handles breadcrump item click
         */
        _this.onBreadcrumpItemClick = function (node) {
            var breadcrumbItems = _this.state.breadcrumbItems;
            var breadcrumbClickedItemIndx = 0;
            // Site node clicked
            if (node.libraryData == null && node.folderData == null) {
                _this.setState({
                    libraryAbsolutePath: undefined,
                    libraryPath: undefined,
                    folderName: undefined
                });
            }
            else if (node.folderData != null) {
                _this._handleOpenFolder(node.folderData, false);
                // select which node has been clicked
                breadcrumbClickedItemIndx = sp_lodash_subset_1.findIndex(breadcrumbItems, function (item) { return item.folderData && item.folderData.absoluteUrl === node.key; });
            }
            else if (node.libraryData != null) {
                _this._handleOpenLibrary(node.libraryData, false);
                // select which node has been clicked
                breadcrumbClickedItemIndx = sp_lodash_subset_1.findIndex(breadcrumbItems, function (item) { return item.libraryData && item.libraryData.serverRelativeUrl === node.key; });
            }
            // Trim nodes array
            breadcrumbItems = breadcrumbItems.slice(0, breadcrumbClickedItemIndx + 1);
            // Set new current node
            breadcrumbItems[breadcrumbItems.length - 1].isCurrentItem = true;
            _this.setState({
                breadcrumbItems: breadcrumbItems,
                filePickerResult: undefined
            });
        };
        /**
         * Is called when user selects a different file
         */
        _this._handleSelectionChange = function (filePickerResult) {
            if (filePickerResult) {
                filePickerResult.downloadFileContent = function () { return _this.props.fileBrowserService.downloadSPFileContent(filePickerResult.fileAbsoluteUrl, filePickerResult.fileName); };
            }
            // this.props.fileBrowserService
            _this.setState({
                filePickerResult: filePickerResult
            });
        };
        /**
         * Called when user saves
         */
        _this._handleSave = function () {
            _this.props.onSave(_this.state.filePickerResult);
        };
        /**
         * Called when user closes tab
         */
        _this._handleClose = function () {
            _this.props.onClose();
        };
        /**
         * Triggered when user opens a file folder
         */
        _this._handleOpenFolder = function (folder, addBreadcrumbNode) {
            var breadcrumbItems = _this.state.breadcrumbItems;
            if (addBreadcrumbNode) {
                breadcrumbItems.map(function (item) { return item.isCurrentItem = false; });
                var breadcrumbNode_1 = {
                    folderData: folder,
                    isCurrentItem: true,
                    text: folder.name,
                    key: folder.absoluteUrl
                };
                breadcrumbNode_1.onClick = function () { _this.onBreadcrumpItemClick(breadcrumbNode_1); };
                breadcrumbItems.push(breadcrumbNode_1);
            }
            _this.setState({
                filePickerResult: null,
                libraryPath: folder.serverRelativeUrl,
                folderName: folder.name,
                libraryAbsolutePath: folder.absoluteUrl,
                breadcrumbItems: breadcrumbItems
            });
        };
        /**
         * Triggered when user opens a top-level document library
         */
        _this._handleOpenLibrary = function (library, addBreadcrumbNode) {
            var breadcrumbItems = _this.state.breadcrumbItems;
            if (addBreadcrumbNode) {
                breadcrumbItems.map(function (item) { return item.isCurrentItem = false; });
                var breadcrumbNode_2 = {
                    libraryData: library,
                    isCurrentItem: true,
                    text: library.title,
                    key: library.serverRelativeUrl
                };
                breadcrumbNode_2.onClick = function () { _this.onBreadcrumpItemClick(breadcrumbNode_2); };
                breadcrumbItems.push(breadcrumbNode_2);
            }
            _this.setState({
                libraryAbsolutePath: library.absoluteUrl,
                libraryTitle: library.title,
                libraryPath: library.serverRelativeUrl,
                breadcrumbItems: breadcrumbItems
            });
        };
        // Add current site to the breadcrumb or the provided node
        var breadcrumbSiteNode = _this.props.breadcrumbFirstNode ? _this.props.breadcrumbFirstNode : {
            isCurrentItem: true,
            text: props.context.pageContext.web.title,
            key: props.context.pageContext.web.id.toString()
        };
        breadcrumbSiteNode.onClick = function () { _this.onBreadcrumpItemClick(breadcrumbSiteNode); };
        _this.state = {
            filePickerResult: null,
            libraryAbsolutePath: undefined,
            libraryTitle: strings.DocumentLibraries,
            libraryPath: undefined,
            folderName: strings.DocumentLibraries,
            breadcrumbItems: [breadcrumbSiteNode]
        };
        return _this;
    }
    SiteFilePickerTab.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: SiteFilePickerTab_module_scss_1.default.tabContainer },
            React.createElement("div", { className: SiteFilePickerTab_module_scss_1.default.tabHeaderContainer },
                React.createElement(Breadcrumb_1.Breadcrumb, { items: this.state.breadcrumbItems, onRenderItem: this.renderBreadcrumbItem, className: SiteFilePickerTab_module_scss_1.default.breadcrumbNav })),
            React.createElement("div", { className: SiteFilePickerTab_module_scss_1.default.tabFiles },
                this.state.libraryAbsolutePath === undefined &&
                    React.createElement(controls_1.DocumentLibraryBrowser, { fileBrowserService: this.props.fileBrowserService, onOpenLibrary: function (selectedLibrary) { return _this._handleOpenLibrary(selectedLibrary, true); } }),
                this.state.libraryAbsolutePath !== undefined &&
                    React.createElement(controls_1.FileBrowser, { onChange: function (filePickerResult) { return _this._handleSelectionChange(filePickerResult); }, onOpenFolder: function (folder) { return _this._handleOpenFolder(folder, true); }, fileBrowserService: this.props.fileBrowserService, libraryName: this.state.libraryTitle, folderPath: this.state.libraryPath, accepts: this.props.accepts })),
            React.createElement("div", { className: SiteFilePickerTab_module_scss_1.default.actionButtonsContainer },
                React.createElement("div", { className: SiteFilePickerTab_module_scss_1.default.actionButtons },
                    React.createElement(Button_1.PrimaryButton, { disabled: !this.state.filePickerResult, onClick: function () { return _this._handleSave(); }, className: SiteFilePickerTab_module_scss_1.default.actionButton }, strings.OpenButtonLabel),
                    React.createElement(Button_1.DefaultButton, { onClick: function () { return _this._handleClose(); }, className: SiteFilePickerTab_module_scss_1.default.actionButton }, strings.CancelButtonLabel)))));
    };
    return SiteFilePickerTab;
}(React.Component));
exports.default = SiteFilePickerTab;

//# sourceMappingURL=SiteFilePickerTab.js.map
