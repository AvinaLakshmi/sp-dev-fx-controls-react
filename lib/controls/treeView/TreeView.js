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
var TreeView_module_scss_1 = require("./TreeView.module.scss");
var uniqBy = require("lodash/uniqBy");
var ITreeViewProps_1 = require("./ITreeViewProps");
var TreeItem_1 = require("./TreeItem");
var telemetry = require("../../common/telemetry");
/**
 * Renders the controls for TreeItem component
 */
var TreeView = (function (_super) {
    __extends(TreeView, _super);
    /**
     * Constructor method
     * @param props properties interface
     */
    function TreeView(props) {
        var _this = _super.call(this, props) || this;
        _this.nodesToExpand = [];
        _this.pathTo = function (array, target) {
            var result;
            array.some(function (_a) {
                var key = _a.key, _b = _a.children, children = _b === void 0 ? [] : _b;
                if (key === target) {
                    _this.nodesToExpand.push(key);
                    result = key;
                    return true;
                }
                var temp = _this.pathTo(children, target);
                if (temp) {
                    _this.nodesToExpand.push(key);
                    result = key + '.' + temp;
                    return true;
                }
            });
            return result;
        };
        telemetry.track('TreeView');
        _this.state = {
            loaded: true,
            defaultExpanded: props.defaultExpanded,
            activeItems: []
        };
        // Bind control events
        _this.handleTreeExpandCollapse = _this.handleTreeExpandCollapse.bind(_this);
        _this.handleOnSelect = _this.handleOnSelect.bind(_this);
        if (props.expandToSelected) {
            props.defaultSelectedKeys.forEach(function (element) {
                _this.pathTo(props.items, element);
            });
        }
        return _this;
    }
    TreeView.prototype.getSelectedItems = function (treeItems, selectedKeys, selectedChildren) {
        var _this = this;
        var selectedItems = [];
        treeItems.forEach(function (item) {
            if (selectedKeys.indexOf(item.key) !== -1 && item.selectable !== false && !item.disabled) {
                selectedItems.push(item);
                if (selectedChildren) {
                    _this.selectAllChildren(item, selectedItems);
                }
            }
            else {
                if (item.children) {
                    selectedItems.push.apply(selectedItems, _this.getSelectedItems(item.children, selectedKeys, selectedChildren));
                }
            }
        });
        return selectedItems;
    };
    /**
     * Fires When expand / collapse item in TreeView
     * @argument item The expanded / collapsed item
     * @argument isExpanded The status of item  (expanded / collapsed)
     */
    TreeView.prototype.handleTreeExpandCollapse = function (item, isExpanded) {
        if (typeof this.props.onExpandCollapse === "function") {
            this.props.onExpandCollapse(item, isExpanded);
        }
    };
    /**
     * Selects all child nodes when parent node is selected.
     * @param item current tree item
     */
    TreeView.prototype.selectAllChildren = function (item, selectedItems) {
        var _this = this;
        if (item.children) {
            item.children.forEach(function (element) {
                if (!element.disabled && element.selectable !== false) {
                    selectedItems.push(element);
                }
                if (element.children) {
                    _this.selectAllChildren(element, selectedItems);
                }
            });
        }
    };
    /**
     * Unselects all child nodes of selected parent.
     */
    TreeView.prototype.unSelectChildren = function (item, unselectArray) {
        var _this = this;
        var tempItem = item;
        if (tempItem.children) {
            tempItem.children.forEach(function (element) {
                unselectArray.push(element.key);
                if (element.children) {
                    _this.unSelectChildren(element, unselectArray);
                }
            });
        }
    };
    /**
     * Fires When Tree Item is selected in TreeView
     * @argument item The selected item
     * @argument isSelected The status of item selection
     */
    TreeView.prototype.handleOnSelect = function (item, isSelected) {
        var selectedItems = this.state.activeItems;
        if (isSelected) {
            if (this.props.selectionMode == ITreeViewProps_1.TreeViewSelectionMode.Multiple) {
                // Add the checked term
                selectedItems.push(item);
                if (this.props.selectChildrenIfParentSelected) {
                    this.selectAllChildren(item, selectedItems);
                }
                selectedItems = uniqBy(selectedItems, 'key');
                // Filter out the duplicate terms
                this.setState({
                    activeItems: selectedItems
                });
            }
            else {
                // Only store the current selected item
                this.setState({
                    activeItems: [item]
                });
                selectedItems = [item];
            }
        }
        else {
            // Remove the item from the list of active nodes
            var unselectArray = [];
            unselectArray.push(item.key);
            if (this.props.selectChildrenIfParentSelected) {
                this.unSelectChildren(item, unselectArray);
            }
            unselectArray.forEach(function (element) {
                selectedItems = selectedItems.filter(function (i) { return i.key != element; });
            });
            this.setState({
                activeItems: selectedItems
            });
        }
        if (typeof this.props.onSelect === "function") {
            this.props.onSelect(selectedItems);
        }
    };
    TreeView.prototype.componentDidMount = function () {
        var _a = this.props, items = _a.items, defaultSelectedKeys = _a.defaultSelectedKeys, selectChildrenIfParentSelected = _a.selectChildrenIfParentSelected;
        if (defaultSelectedKeys) {
            var selectedItems = this.getSelectedItems(items, defaultSelectedKeys, selectChildrenIfParentSelected);
            this.setState({
                activeItems: selectedItems
            });
        }
    };
    /**
     * Default React render method
     */
    TreeView.prototype.render = function () {
        var _this = this;
        var _a = this.props, items = _a.items, selectionMode = _a.selectionMode, onRenderItem = _a.onRenderItem, showCheckboxes = _a.showCheckboxes, treeItemActionsDisplayMode = _a.treeItemActionsDisplayMode, defaultExpanded = _a.defaultExpanded;
        return (React.createElement("div", { className: TreeView_module_scss_1.default.treeView }, items.map(function (treeNodeItem, index) { return (React.createElement(TreeItem_1.default, { treeItem: treeNodeItem, leftOffset: 20, isFirstRender: true, defaultExpanded: defaultExpanded, selectionMode: selectionMode, activeItems: _this.state.activeItems, parentCallbackExpandCollapse: _this.handleTreeExpandCollapse, parentCallbackOnSelect: _this.handleOnSelect, onRenderItem: onRenderItem, showCheckboxes: showCheckboxes, treeItemActionsDisplayMode: treeItemActionsDisplayMode, nodesToExpand: _this.nodesToExpand })); })));
    };
    return TreeView;
}(React.Component));
exports.TreeView = TreeView;

//# sourceMappingURL=TreeView.js.map
