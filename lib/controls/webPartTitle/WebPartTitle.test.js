"use strict";
/// <reference types="sinon" />
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var chai_1 = require("chai");
var enzyme_1 = require("enzyme");
var WebPartTitle_1 = require("./WebPartTitle");
var WebPartTitle_module_scss_1 = require("./WebPartTitle.module.scss");
var sp_core_library_1 = require("@microsoft/sp-core-library");
describe('<WebPartTitle />', function () {
    var webparttitle;
    var dummyTitle = "Dummy Title";
    var dummyClass = "DummyClass";
    var dummyMoreLink = "See all";
    var dummyColor = "#ffffff";
    var dummyPlaceholder = "News";
    var dummyUpdateFnc = sinon.spy(function (value) { return value; });
    afterEach(function () {
        webparttitle.unmount();
    });
    it('Check the read mode of the component', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Read, title: dummyTitle, updateProperty: function () { } }));
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " span").text()).to.be.equal(dummyTitle);
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " textarea")).to.have.length(0);
        done();
    });
    it('Check the edit mode of the component', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Edit, title: dummyTitle, updateProperty: function () { } }));
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " span")).to.have.length(0);
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " textarea")).to.have.length(1);
        done();
    });
    it('Check class change', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Read, title: dummyTitle, updateProperty: function () { }, className: dummyClass }));
        chai_1.expect(webparttitle.find("div." + dummyClass + " span").text()).to.be.equal(dummyTitle);
        done();
    });
    it('Check if the change is processed correctly', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Edit, title: dummyTitle, updateProperty: dummyUpdateFnc }));
        // Specify a new web part title
        var textArea = webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " textarea");
        textArea.simulate('change', { target: { value: "New web part title" } });
        // Check if function is called
        /* tslint:disable */
        chai_1.expect(dummyUpdateFnc.called).to.be.true;
        /* tslint:enable */
        // Check if the returned value is correct
        chai_1.expect(dummyUpdateFnc.args[0]).contains("New web part title");
        done();
    });
    it('Check more link is shown if function specified', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Read, title: dummyTitle, updateProperty: function () { }, moreLink: function () { return React.createElement("a", { href: "#" }, dummyMoreLink); } }));
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " span").text()).to.be.equal(dummyTitle);
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " textarea")).to.have.length(0);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink)).to.have.length(1);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink + " a")).to.have.length(1);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink + " a").text()).to.be.equal(dummyMoreLink);
        done();
    });
    it('Check more link is shown if children specified', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Read, title: dummyTitle, updateProperty: function () { }, moreLink: React.createElement("a", { href: "#" }, dummyMoreLink) }));
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " span").text()).to.be.equal(dummyTitle);
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " textarea")).to.have.length(0);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink)).to.have.length(1);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink + " a")).to.have.length(1);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink + " a").text()).to.be.equal(dummyMoreLink);
        done();
    });
    it('Check more link is not shown otherwise', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Read, title: dummyTitle, updateProperty: function () { } }));
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " span").text()).to.be.equal(dummyTitle);
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle + " textarea")).to.have.length(0);
        chai_1.expect(webparttitle.find("span." + WebPartTitle_module_scss_1.default.moreLink)).to.have.length(0);
        done();
    });
    it('Check theme\'s color is used if specified', function (done) {
        webparttitle = enzyme_1.mount(React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: sp_core_library_1.DisplayMode.Edit, title: dummyTitle, updateProperty: dummyUpdateFnc, themeVariant: { semanticColors: { bodyText: dummyColor } } }));
        chai_1.expect(webparttitle.find("div." + WebPartTitle_module_scss_1.default.webPartTitle).prop('style')).property("color").to.equal(dummyColor);
        done();
    });
});

//# sourceMappingURL=WebPartTitle.test.js.map
