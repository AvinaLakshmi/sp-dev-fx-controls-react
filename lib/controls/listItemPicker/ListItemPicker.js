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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var strings = require("ControlStrings");
var React = require("react");
var SPService_1 = require("../../services/SPService");
var TagPicker_1 = require("office-ui-fabric-react/lib/components/pickers/TagPicker/TagPicker");
var Label_1 = require("office-ui-fabric-react/lib/Label");
var telemetry = require("../../common/telemetry");
var ListItemPicker = (function (_super) {
    __extends(ListItemPicker, _super);
    function ListItemPicker(props) {
        var _this = _super.call(this, props) || this;
        /**
         * On Selected Item
         */
        _this.onItemChanged = function (selectedItems) {
            _this.selectedItems = selectedItems;
            _this.props.onSelectedItem(selectedItems);
        };
        /**
         * Filter Change
         */
        _this.onFilterChanged = function (filterText, tagList) { return __awaiter(_this, void 0, void 0, function () {
            var resolvedSugestions, filteredSuggestions, _loop_1, this_1, _i, resolvedSugestions_1, suggestion;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.loadListItems(filterText)];
                    case 1:
                        resolvedSugestions = _a.sent();
                        // Filter out the already retrieved items, so that they cannot be selected again
                        if (this.selectedItems && this.selectedItems.length > 0) {
                            filteredSuggestions = [];
                            _loop_1 = function (suggestion) {
                                var exists = this_1.selectedItems.filter(function (sItem) { return sItem.key === suggestion.key; });
                                if (!exists || exists.length === 0) {
                                    filteredSuggestions.push(suggestion);
                                }
                            };
                            this_1 = this;
                            for (_i = 0, resolvedSugestions_1 = resolvedSugestions; _i < resolvedSugestions_1.length; _i++) {
                                suggestion = resolvedSugestions_1[_i];
                                _loop_1(suggestion);
                            }
                            resolvedSugestions = filteredSuggestions;
                        }
                        if (resolvedSugestions) {
                            this.setState({
                                errorMessage: "",
                                showError: false
                            });
                            return [2 /*return*/, resolvedSugestions];
                        }
                        else {
                            return [2 /*return*/, []];
                        }
                        return [2 /*return*/];
                }
            });
        }); };
        /**
         * Function to load List Items
         */
        _this.loadListItems = function (filterText) { return __awaiter(_this, void 0, void 0, function () {
            var _a, listId, columnInternalName, keyColumnInternalName, webUrl, filter, substringSearch, arrayItems, keyColumn, listItems, _i, listItems_1, item, error_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this.props, listId = _a.listId, columnInternalName = _a.columnInternalName, keyColumnInternalName = _a.keyColumnInternalName, webUrl = _a.webUrl, filter = _a.filter, substringSearch = _a.substringSearch;
                        arrayItems = [];
                        keyColumn = keyColumnInternalName || 'Id';
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._spservice.getListItems(filterText, listId, columnInternalName, keyColumn, webUrl, filter, substringSearch)];
                    case 2:
                        listItems = _b.sent();
                        // Check if the list had items
                        if (listItems.length > 0) {
                            for (_i = 0, listItems_1 = listItems; _i < listItems_1.length; _i++) {
                                item = listItems_1[_i];
                                arrayItems.push({ key: item[keyColumn], name: item[columnInternalName] });
                            }
                        }
                        return [2 /*return*/, arrayItems];
                    case 3:
                        error_1 = _b.sent();
                        console.log("Error get Items", error_1);
                        this.setState({
                            showError: true,
                            errorMessage: error_1.message,
                            noresultsFoundText: error_1.message
                        });
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        }); };
        telemetry.track('ListItemPicker', {});
        // States
        _this.state = {
            noresultsFoundText: !_this.props.noResultsFoundText ? strings.genericNoResultsFoundText : _this.props.noResultsFoundText,
            showError: false,
            errorMessage: "",
            suggestionsHeaderText: !_this.props.suggestionsHeaderText ? strings.ListItemPickerSelectValue : _this.props.suggestionsHeaderText
        };
        // Get SPService Factory
        _this._spservice = new SPService_1.default(_this.props.context);
        _this.selectedItems = [];
        return _this;
    }
    ListItemPicker.prototype.componentDidUpdate = function (prevProps, prevState) {
        if (this.props.listId !== prevProps.listId) {
            this.selectedItems = [];
        }
    };
    /**
     * Render the field
     */
    ListItemPicker.prototype.render = function () {
        var _a = this.props, className = _a.className, disabled = _a.disabled, itemLimit = _a.itemLimit, placeholder = _a.placeholder;
        return (React.createElement("div", null,
            React.createElement(TagPicker_1.TagPicker, { onResolveSuggestions: this.onFilterChanged, 
                //   getTextFromItem={(item: any) => { return item.name; }}
                getTextFromItem: this.getTextFromItem, pickerSuggestionsProps: {
                    suggestionsHeaderText: this.state.suggestionsHeaderText,
                    noResultsFoundText: this.state.noresultsFoundText
                }, defaultSelectedItems: this.props.defaultSelectedItems || [], onChange: this.onItemChanged, className: className, itemLimit: itemLimit, disabled: disabled, inputProps: {
                    placeholder: placeholder
                } }),
            !!this.state.errorMessage &&
                (React.createElement(Label_1.Label, { style: { color: '#FF0000' } },
                    " ",
                    this.state.errorMessage,
                    " "))));
    };
    /**
     * Get text from Item
     */
    ListItemPicker.prototype.getTextFromItem = function (item) {
        return item.name;
    };
    return ListItemPicker;
}(React.Component));
exports.ListItemPicker = ListItemPicker;

//# sourceMappingURL=ListItemPicker.js.map
