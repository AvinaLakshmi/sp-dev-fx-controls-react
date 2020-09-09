"use strict";
/// <reference types="sinon" />
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var chai_1 = require("chai");
var enzyme_1 = require("enzyme");
var Styling_1 = require("office-ui-fabric-react/lib/Styling");
var PlaceholderComponent_1 = require("./PlaceholderComponent");
var PlaceholderComponent_module_scss_1 = require("./PlaceholderComponent.module.scss");
describe('<Placeholder />', function () {
    var placeholder;
    var dummyDescription = "Dummy description";
    var dummyIcon = "Add";
    var dummyText = "Dummy icon text";
    var dummyClass = "dummyClass";
    var dummyLabel = "Dummy label";
    var dummyOnConfigure = sinon.spy(function (evt) { });
    before(function () {
        // Suppress icon warnings.
        Styling_1.setIconOptions({
            disableWarnings: true
        });
    });
    afterEach(function () {
        placeholder.unmount();
    });
    it('Test that placeholder renders', function (done) {
        placeholder = enzyme_1.mount(React.createElement(PlaceholderComponent_1.Placeholder, { description: dummyDescription, iconName: dummyIcon, iconText: dummyText }));
        done();
    });
    it('Test placeholder without button', function (done) {
        placeholder = enzyme_1.mount(React.createElement(PlaceholderComponent_1.Placeholder, { description: dummyDescription, iconName: dummyIcon, iconText: dummyText }));
        chai_1.expect(placeholder.find('i[data-icon-name="Add"]')).to.have.length(1);
        chai_1.expect(placeholder.find("." + PlaceholderComponent_module_scss_1.default.placeholderText)).to.have.length(1);
        chai_1.expect(placeholder.find("." + PlaceholderComponent_module_scss_1.default.placeholderText).text()).to.be.equal(dummyText);
        chai_1.expect(placeholder.find("." + PlaceholderComponent_module_scss_1.default.placeholderDescriptionText)).to.have.length(1);
        chai_1.expect(placeholder.find("." + PlaceholderComponent_module_scss_1.default.placeholderDescriptionText).text()).to.be.equal(dummyDescription);
        chai_1.expect(placeholder.find('button')).to.have.length(0);
        done();
    });
    it('Test placeholder with custom classname', function (done) {
        placeholder = enzyme_1.mount(React.createElement(PlaceholderComponent_1.Placeholder, { description: dummyDescription, iconName: dummyIcon, iconText: dummyText, contentClassName: dummyClass }));
        chai_1.expect(placeholder.find("div." + PlaceholderComponent_module_scss_1.default.placeholder + "." + dummyClass)).to.have.length(1);
        done();
    });
    it('Test placeholder with null values', function (done) {
        placeholder = enzyme_1.mount(React.createElement(PlaceholderComponent_1.Placeholder, { description: null, iconName: null, iconText: null }));
        chai_1.expect(placeholder.find('i[data-icon-name="Add"]')).to.have.length(0);
        chai_1.expect(placeholder.find("div ." + PlaceholderComponent_module_scss_1.default.placeholderText)).to.have.length(1);
        chai_1.expect(placeholder.find("div ." + PlaceholderComponent_module_scss_1.default.placeholderText).text()).to.be.equal('');
        chai_1.expect(placeholder.find("div ." + PlaceholderComponent_module_scss_1.default.placeholderDescriptionText)).to.have.length(1);
        chai_1.expect(placeholder.find("div ." + PlaceholderComponent_module_scss_1.default.placeholderDescriptionText).text()).to.be.equal('');
        chai_1.expect(placeholder.find('button')).to.have.length(0);
        done();
    });
    it('Test placeholder with button', function (done) {
        placeholder = enzyme_1.mount(React.createElement(PlaceholderComponent_1.Placeholder, { description: dummyDescription, iconName: dummyIcon, iconText: dummyText, buttonLabel: dummyLabel }));
        chai_1.expect(placeholder.find("button")).to.have.length(1);
        chai_1.expect(placeholder.find("button").text()).to.contain(dummyLabel);
        done();
    });
    it('Test button onConfigure trigger', function (done) {
        placeholder = enzyme_1.mount(React.createElement(PlaceholderComponent_1.Placeholder, { description: dummyDescription, iconName: dummyIcon, iconText: dummyText, buttonLabel: dummyLabel, onConfigure: dummyOnConfigure }));
        placeholder.find('button').simulate('click');
        /* tslint:disable */
        chai_1.expect(dummyOnConfigure.called).to.be.true;
        /* tslint:enable */
        done();
    });
});

//# sourceMappingURL=PlaceholderComponent.test.js.map
