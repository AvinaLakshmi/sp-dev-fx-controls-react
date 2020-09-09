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
var ReactDom = require("react-dom");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var strings = require("ControlsTestWebPartStrings");
var ControlsTest_1 = require("./components/ControlsTest");
/**
 * Web part to test the React controls
 */
var ControlsTestWebPart = (function (_super) {
    __extends(ControlsTestWebPart, _super);
    function ControlsTestWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ControlsTestWebPart.prototype.render = function () {
        var _this = this;
        var element = React.createElement(ControlsTest_1.default, {
            context: this.context,
            description: this.properties.description,
            title: this.properties.title,
            displayMode: this.displayMode,
            updateProperty: function (value) {
                _this.properties.title = value;
            },
            totalPages: this.properties.totalPages
        });
        ReactDom.render(element, this.domElement);
    };
    Object.defineProperty(ControlsTestWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    ControlsTestWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                sp_webpart_base_1.PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('totalPages', {
                                    label: 'Total pages in pagination'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return ControlsTestWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = ControlsTestWebPart;

//# sourceMappingURL=ControlsTestWebPart.js.map
