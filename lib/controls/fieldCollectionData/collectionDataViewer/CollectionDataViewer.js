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
var FieldCollectionData_module_scss_1 = require("../FieldCollectionData.module.scss");
var collectionDataItem_1 = require("../collectionDataItem");
var Button_1 = require("office-ui-fabric-react/lib/components/Button");
var Icon_1 = require("office-ui-fabric-react/lib/components/Icon");
var strings = require("ControlStrings");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var pagination_1 = require("../../pagination");
var TextField_1 = require("office-ui-fabric-react/lib/TextField");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var CollectionDataViewer = (function (_super) {
    __extends(CollectionDataViewer, _super);
    function CollectionDataViewer(props) {
        var _this = _super.call(this, props) || this;
        _this.SORT_IDX = "sortIdx";
        /**
         * Add a new item to the collection
         */
        _this.addItem = function (item) {
            if (!item.uniqueId) {
                item.uniqueId = sp_core_library_1.Guid.newGuid();
            }
            _this.setState(function (prevState) {
                var crntItems = prevState.crntItems.concat([item]);
                crntItems = _this.updateSortProperty(crntItems);
                var pagesCount = _this.getPagesCount(crntItems);
                var currentPage = pagesCount < 1 ? 1 : pagesCount;
                return {
                    crntItems: crntItems,
                    inCreationItem: null,
                    currentPage: currentPage
                };
            });
        };
        /**
         * Remove an item from the collection
         */
        _this.updateItem = function (idx, item) {
            _this.setState(function (prevState) {
                var crntItems = prevState.crntItems;
                // Update the item in the array
                crntItems[idx] = item;
                return {
                    crntItems: crntItems,
                    currentPage: prevState.currentPage
                };
            });
        };
        /**
         * Remove an item from the collection
         */
        _this.deleteItem = function (idx) {
            _this.setState(function (prevState) {
                var crntItems = prevState.crntItems;
                crntItems.splice(idx, 1);
                // Update the sort propety
                crntItems = _this.updateSortProperty(crntItems);
                // Calculate current page
                var pagesCount = _this.getPagesCount(crntItems);
                var currentPage = _this.state.currentPage > pagesCount ? pagesCount : _this.state.currentPage;
                return {
                    crntItems: sp_lodash_subset_1.sortBy(crntItems, _this.SORT_IDX),
                    currentPage: currentPage
                };
            });
        };
        /**
         * Validate every item
         */
        _this.validateItem = function (idx, isValid) {
            _this.setState(function (prevState) {
                var validation = prevState.validation;
                validation[idx] = isValid;
                return {
                    validation: prevState.validation
                };
            });
        };
        /**
         * Currently in creation
         */
        _this.addInCreation = function (item) {
            _this.setState({
                inCreationItem: item
            });
        };
        /**
         * Add the item and save the form
         */
        _this.addAndSave = function () {
            // Check if the item is not empty
            if (_this.state.inCreationItem) {
                _this.props.fOnSave(_this.state.crntItems.concat([_this.state.inCreationItem]));
            }
            else {
                _this.onSave();
            }
        };
        /**
         * Update the sort order
         */
        _this.updateSortOrder = function (oldIdx, newIdx) {
            _this.setState(function (prevState) {
                var crntItems = prevState.crntItems;
                var newOrderedItems = sp_lodash_subset_1.cloneDeep(crntItems);
                newOrderedItems = _this.moveItemTo(newOrderedItems, oldIdx, newIdx - 1);
                newOrderedItems = _this.updateSortProperty(newOrderedItems);
                newOrderedItems = sp_lodash_subset_1.sortBy(newOrderedItems, _this.SORT_IDX);
                return {
                    crntItems: newOrderedItems
                };
            });
        };
        /**
         * Save the collection data
         */
        _this.onSave = function () {
            _this.props.fOnSave(_this.state.crntItems);
        };
        /**
         * Cancel
         */
        _this.onCancel = function () {
            _this.props.fOnClose();
        };
        _this.getPagesCount = function (items) {
            if (!_this.isPagingEnabled()) {
                return 1;
            }
            return Math.ceil(items.length / _this.props.itemsPerPage);
        };
        _this.getPageItems = function (currentPageIndex, currentItems) {
            if (currentPageIndex === void 0) { currentPageIndex = null; }
            if (currentItems === void 0) { currentItems = null; }
            var _a = _this.state, crntItems = _a.crntItems, currentPage = _a.currentPage;
            var items = !currentItems ? crntItems : currentItems;
            var pageIndex = !currentPageIndex ? currentPage : currentPageIndex;
            if (!_this.isPagingEnabled || !items) {
                return items;
            }
            var startIndex = _this.getFirstElementIndex(pageIndex, items);
            var endIndex = _this.getLastElementIndex(pageIndex, items);
            return items.slice(startIndex, endIndex);
        };
        _this.getFirstElementIndex = function (currentPage, crntItems) {
            var itemsPerPage = _this.props.itemsPerPage;
            if (!crntItems) {
                return null;
            }
            var isPginingEnabled = _this.isPagingEnabled();
            if (!isPginingEnabled || currentPage <= 1) {
                return 0;
            }
            var firstElementIndex = (currentPage - 1) * itemsPerPage;
            return firstElementIndex;
        };
        _this.getLastElementIndex = function (currentPage, crntItems) {
            var itemsPerPage = _this.props.itemsPerPage;
            if (!crntItems) {
                return null;
            }
            var isPginingEnabled = _this.isPagingEnabled();
            if (!isPginingEnabled) {
                return crntItems.length;
            }
            var lastElementIndex = currentPage * itemsPerPage > crntItems.length ? crntItems.length : currentPage * itemsPerPage;
            return lastElementIndex;
        };
        _this.getCollectionDataItem = function (item, idx, allItems) {
            return React.createElement(collectionDataItem_1.CollectionDataItem, { key: item.uniqueId, fields: _this.props.fields, index: idx, item: item, totalItems: allItems.length, sortingEnabled: _this.props.enableSorting, disableItemDeletion: _this.props.disableItemDeletion, fUpdateItem: _this.updateItem, fDeleteItem: _this.deleteItem, fValidation: _this.validateItem, fOnSorting: _this.updateSortOrder });
        };
        _this.isPagingEnabled = function () {
            var isPagingEnabled = !!_this.props.itemsPerPage && _this.props.itemsPerPage > 0;
            return isPagingEnabled;
        };
        _this.executeItemsFiltering = function (items) {
            var executeFiltering = _this.props.executeFiltering;
            var searchFilter = _this.state.searchFilter;
            if (!items || items.length <= 0) {
                return false;
            }
            if (executeFiltering == null) {
                return false;
            }
            if (sp_lodash_subset_1.isEmpty(searchFilter)) {
                return false;
            }
            return true;
        };
        _this.getFilteredItems = function (items) {
            var executeFiltering = _this.props.executeFiltering;
            var searchFilter = _this.state.searchFilter;
            return items.filter(function (item) { return executeFiltering(searchFilter, item); });
        };
        _this.state = {
            crntItems: [],
            inCreationItem: null,
            validation: {},
            currentPage: 1
        };
        return _this;
    }
    /**
     * componentDidMount lifecycle hook
     */
    CollectionDataViewer.prototype.componentDidMount = function () {
        var _this = this;
        var crntItems = this.props.value ? sp_lodash_subset_1.sortBy(sp_lodash_subset_1.cloneDeep(this.props.value), this.SORT_IDX) : [];
        crntItems = crntItems.map(function (item, idx) {
            if (!item[_this.SORT_IDX]) {
                item[_this.SORT_IDX] = idx + 1;
            }
            if (!item.uniqueId) {
                item.uniqueId = sp_core_library_1.Guid.newGuid();
            }
            return item;
        });
        // Update the sort propety
        crntItems = this.updateSortProperty(crntItems);
        this.setState({
            crntItems: sp_lodash_subset_1.sortBy(crntItems, this.SORT_IDX)
        });
    };
    /**
     * Check if all items are valid
     */
    CollectionDataViewer.prototype.allItemsValid = function () {
        var validation = this.state.validation;
        if (validation) {
            var keys = Object.keys(validation);
            for (var _i = 0, keys_1 = keys; _i < keys_1.length; _i++) {
                var key = keys_1[_i];
                if (!validation[key]) {
                    return false;
                }
            }
        }
        return true;
    };
    /**
     * Move an item in the array
     *
     * @param crntItems
     * @param oldIdx
     * @param newIdx
     */
    CollectionDataViewer.prototype.moveItemTo = function (crntItems, oldIdx, newIdx) {
        if (newIdx > -1 && newIdx < crntItems.length) {
            var removedElement = crntItems.splice(oldIdx, 1)[0];
            if (removedElement) {
                crntItems.splice(newIdx, 0, removedElement);
            }
        }
        return crntItems;
    };
    /**
     * Update the sort property
     *
     * @param crntItems
     */
    CollectionDataViewer.prototype.updateSortProperty = function (crntItems) {
        var _this = this;
        // Update the sort order
        return crntItems.map(function (item, itemIdx) {
            item[_this.SORT_IDX] = itemIdx + 1;
            return item;
        });
    };
    /**
     * Default React render
     */
    CollectionDataViewer.prototype.render = function () {
        var _this = this;
        var _a = this.state, currentPage = _a.currentPage, crntItems = _a.crntItems;
        var isPagingEnabled = this.isPagingEnabled();
        if (this.executeItemsFiltering(crntItems)) {
            crntItems = this.getFilteredItems(crntItems);
        }
        var pageCount = this.getPagesCount(crntItems);
        var elements = this.getPageItems(currentPage, crntItems);
        return (React.createElement("div", null,
            this.props.executeFiltering &&
                React.createElement(TextField_1.TextField, { onChanged: function (newValue) { _this.setState({ searchFilter: newValue, currentPage: 1 }); }, placeholder: strings.CollectionDataSearch, className: "FieldCollectionData__panel__search-box" }),
            React.createElement("div", { className: "FieldCollectionData__panel__table " + FieldCollectionData_module_scss_1.default.table + " " + (this.props.tableClassName || "") },
                React.createElement("div", { className: "FieldCollectionData__panel__table-head " + FieldCollectionData_module_scss_1.default.tableRow + " " + FieldCollectionData_module_scss_1.default.tableHead },
                    this.props.enableSorting && (React.createElement("span", { className: "FieldCollectionData__panel__table-cell " + FieldCollectionData_module_scss_1.default.tableCell })),
                    this.props.fields.map(function (f) { return (React.createElement("span", { key: "dataviewer-" + f.id, className: "FieldCollectionData__panel__table-cell " + FieldCollectionData_module_scss_1.default.tableCell },
                        f.title,
                        " ",
                        f.required && React.createElement(Icon_1.Icon, { className: FieldCollectionData_module_scss_1.default.required, iconName: "Asterisk" }))); }),
                    React.createElement("span", { className: "FieldCollectionData__panel__table-cell " + FieldCollectionData_module_scss_1.default.tableCell }),
                    React.createElement("span", { className: "FieldCollectionData__panel__table-cell " + FieldCollectionData_module_scss_1.default.tableCell })),
                (this.state.crntItems && this.state.crntItems.length > 0) &&
                    this.state.crntItems.map(function (item, idx, allItems) {
                        var elementIndex = sp_lodash_subset_1.findIndex(elements, function (x) { return x.uniqueId === item.uniqueId; });
                        if (elementIndex >= 0) {
                            return _this.getCollectionDataItem(item, idx, allItems);
                        }
                        else {
                            return null;
                        }
                    }),
                !this.props.disableItemCreation && (React.createElement(collectionDataItem_1.CollectionDataItem, { fields: this.props.fields, index: null, item: null, sortingEnabled: this.props.enableSorting, totalItems: null, fAddItem: this.addItem, fAddInCreation: this.addInCreation }))),
            isPagingEnabled && this.state.crntItems && this.state.crntItems.length > 0 &&
                React.createElement("div", { className: "FieldCollectionData__panel__pagination" },
                    React.createElement(pagination_1.Pagination, { currentPage: currentPage, totalPages: pageCount, onChange: function (page) { _this.setState({ currentPage: page }); } })),
            (!this.state.crntItems || this.state.crntItems.length === 0) && (React.createElement("p", { className: "FieldCollectionData__panel__no-collection-data " + FieldCollectionData_module_scss_1.default.noCollectionData }, strings.CollectionDataEmptyValue)),
            React.createElement("div", { className: "FieldCollectionData__panel__actions " + FieldCollectionData_module_scss_1.default.panelActions },
                this.state.inCreationItem && React.createElement(Button_1.PrimaryButton, { text: this.props.saveAndAddBtnLabel || strings.CollectionSaveAndAddButtonLabel, onClick: this.addAndSave, disabled: !this.allItemsValid(), className: "FieldCollectionData__panel__action__add" }),
                !this.state.inCreationItem && React.createElement(Button_1.PrimaryButton, { text: this.props.saveBtnLabel || strings.SaveButtonLabel, onClick: this.onSave, disabled: !this.allItemsValid(), className: "FieldCollectionData__panel__action__save" }),
                React.createElement(Button_1.DefaultButton, { text: this.props.cancelBtnLabel || strings.CancelButtonLabel, onClick: this.onCancel, className: "FieldCollectionData__panel__action__cancel" }))));
    };
    return CollectionDataViewer;
}(React.Component));
exports.CollectionDataViewer = CollectionDataViewer;

//# sourceMappingURL=CollectionDataViewer.js.map
