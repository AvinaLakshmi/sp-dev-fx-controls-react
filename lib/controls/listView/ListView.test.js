"use strict";
/// <reference types="sinon" />
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
var chai_1 = require("chai");
var enzyme_1 = require("enzyme");
var ListView_1 = require("./ListView");
var Wrapper = (function (_super) {
    __extends(Wrapper, _super);
    function Wrapper() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Wrapper.prototype.render = function () {
        return React.Children.only(this.props.children);
    };
    return Wrapper;
}(React.Component));
/* tslint:disable */
describe('<ListView />', function () {
    var listView;
    var dummyItems = [{
            id: 0,
            title: 'Dummy item 1',
            path: 'https://contoso.sharepoint.com/documents/word.docx',
            prop: {
                sub: 1
            }
        }, {
            id: 1,
            title: 'Dummy item 2',
            path: 'https://contoso.sharepoint.com/documents/powerpoint.pptx',
            prop: {
                sub: 2
            }
        }];
    var dummyViewFields = [
        {
            name: "id",
            displayName: "ID",
            sorting: true
        },
        {
            name: "title",
            displayName: "Document title",
            linkPropertyName: "path",
            sorting: false
        },
        {
            name: "path",
            render: function (item) {
                return React.createElement("a", { href: item["ServerRelativeUrl"] }, "Link");
            }
        }
    ];
    afterEach(function () {
        listView.unmount();
    });
    beforeEach(function () { });
    it('Test view with an empty array of items', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: [] })).update();
        chai_1.expect(listView.state('items')).to.have.length(0);
        chai_1.expect(listView.state('columns')).to.be.null;
        done();
    });
    it('Test view with two items', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: dummyItems })).update();
        chai_1.expect(listView.state('items')).to.have.length(2);
        chai_1.expect(listView.state('columns')).to.be.null;
        done();
    });
    it('Test view by updating the items porperty', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: [] })).update();
        chai_1.expect(listView.state('items')).to.have.length(0);
        // Update the items property, now two items should be available
        listView.setProps({ items: dummyItems }).update();
        chai_1.expect(listView.state('items')).to.have.length(2);
        done();
    });
    it('Test view with iconFieldName', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: dummyItems, iconFieldName: "path" })).update();
        chai_1.expect(listView.state('items')).to.have.length(2);
        chai_1.expect(listView.state('columns')).to.have.length(1);
        chai_1.expect(listView.state('columns')[0]).to.have.property('key', 'fileType');
        done();
    });
    it('Test view with iconFieldName render method', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: dummyItems, iconFieldName: "path" })).update();
        var iconField = enzyme_1.mount(React.createElement(Wrapper, {}, listView.state('columns')[0]["onRender"](dummyItems[0])));
        // Check if Word icon gets rendered
        chai_1.expect(iconField.find('.ms-BrandIcon--docx')).to.have.length(1);
        done();
    });
    it('Test view with _flattenItems function', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: dummyItems })).update();
        chai_1.expect(listView.state('items')).to.have.length(2);
        chai_1.expect(listView.state('items')[0]).to.have.property('prop.sub', 1);
        chai_1.expect(listView.state('items')[1]).to.have.property('prop.sub', 2);
        done();
    });
    it('Test view with viewFields', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: dummyItems, viewFields: dummyViewFields })).update();
        chai_1.expect(listView.state('columns')).to.have.length(3);
        chai_1.expect(listView.state('columns')[0]).to.have.property('key', 'id');
        chai_1.expect(listView.state('columns')[0]).to.have.property('name', 'ID');
        chai_1.expect(listView.state('columns')[0]).to.have.property('onRender', undefined);
        chai_1.expect(listView.state('columns')[1]).to.have.property('key', 'title');
        chai_1.expect(listView.state('columns')[1]).to.have.property('name', 'Document title');
        // onRender binded function because it is linked to path
        chai_1.expect(listView.state('columns')[1]).to.have.property('onRender');
        chai_1.expect(listView.state('columns')[2]).to.have.property('key', 'path');
        chai_1.expect(listView.state('columns')[2]).to.have.property('name', 'path');
        // onRender binded function because a custom render function got specified
        chai_1.expect(listView.state('columns')[2]).to.have.property('onRender');
        done();
    });
    it('Test out the sorting method', function (done) {
        listView = enzyme_1.mount(React.createElement(ListView_1.ListView, { items: dummyItems, viewFields: dummyViewFields })).update();
        // Check if the ID column is not yet sorted
        chai_1.expect(listView.state('columns')[0]).not.to.have.property('isSorted');
        // Test the column click event, this will trigger to sort the items
        listView.instance()["_columnClick"](null, listView.state('columns')[0]);
        // Check if the column is set to sorted
        chai_1.expect(listView.state('columns')[0]).to.have.property('isSorted', true);
        chai_1.expect(listView.state('columns')[0]).to.have.property('isSortedDescending', false);
        // Check if the first item has ID 0 (ascending order)
        chai_1.expect(listView.state('items')[0]).to.have.property('id', 0);
        // Do another sorting test, this should sort the items in descending order
        listView.instance()["_columnClick"](null, listView.state('columns')[0]);
        chai_1.expect(listView.state('columns')[0]).to.have.property('isSorted', true);
        chai_1.expect(listView.state('columns')[0]).to.have.property('isSortedDescending', true);
        // Check if the first item has ID 0 (ascending order)
        chai_1.expect(listView.state('items')[0]).to.have.property('id', 1);
        // Just to be sure, we should go back to ascending order
        listView.instance()["_columnClick"](null, listView.state('columns')[0]);
        // Check if the first item has ID 0 (ascending order)
        chai_1.expect(listView.state('items')[0]).to.have.property('id', 0);
        done();
    });
});
/* tslint:enable */

//# sourceMappingURL=ListView.test.js.map
