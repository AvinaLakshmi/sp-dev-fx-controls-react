"use strict";
/// <reference types="sinon" />
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var chai_1 = require("chai");
var enzyme_1 = require("enzyme");
var ChartControl_1 = require("./ChartControl");
var ChartControl_module_scss_1 = require("./ChartControl.module.scss");
describe('<ChartControl />', function () {
    var chartControl;
    var dummyTitle = "Dummy Title";
    var dummyClass = "DummyClass";
    var dummySummary = "Dummy Summary";
    var dummyData = {
        labels: ['Human', 'Chimp', 'Dolphin', 'Cat'],
        datasets: [
            {
                label: 'Millions',
                data: [
                    11000, 6200, 5800, 300
                ]
            }
        ]
    };
    afterEach(function () {
        chartControl.unmount();
    });
    it('Check that an accessible table gets created by default', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable)).to.have.length(1);
        done();
    });
    it('Check that the accessible table accepts a custom classname', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData, accessibility: { className: "customClass" } }));
        chai_1.expect(chartControl.find("div.customClass")).to.have.length(1);
        done();
    });
    it('Check that the accessible table doesn\'t get rendered if disabled', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData, accessibility: { enable: false } }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable)).to.have.length(0);
        done();
    });
    it('Check that an accessible table gets created with the caption matching the title', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", accessibility: { enable: true }, options: { title: { text: dummyTitle } }, data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable + " table caption").text()).to.be.equal(dummyTitle);
        done();
    });
    it('Check that an accessible table gets created with a caption', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", accessibility: { enable: true, caption: dummyTitle }, data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable + " table caption").text()).to.be.equal(dummyTitle);
        done();
    });
    it('Check that the accessible table has a number of rows matching the number of data elements', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable + " table tbody tr")).to.have.length(4);
        done();
    });
    it('Check that the accessible table has only one header row', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable + " table thead tr")).to.have.length(1);
        done();
    });
    it('Check that custom class gets rendered', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", className: dummyClass, data: dummyData }));
        chai_1.expect(chartControl.find("div." + dummyClass)).to.have.length(1);
        done();
    });
    it('Check that a canvas gets rendered', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.chartComponent + " canvas")).to.have.length(1);
        done();
    });
    it('Check that it doesn\'t crash if data is omitted', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar" }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.chartComponent + " canvas")).to.have.length(1);
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.accessibleTable)).to.have.length(0);
        done();
    });
    it('Check that it applies a themed background by default', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.chartComponent + "." + ChartControl_module_scss_1.default.themed)).to.have.length(1);
        done();
    });
    it('Check that it disables themed background when useTheme is set to false', function (done) {
        chartControl = enzyme_1.mount(React.createElement(ChartControl_1.ChartControl, { type: "bar", data: dummyData, useTheme: false }));
        chai_1.expect(chartControl.find("div." + ChartControl_module_scss_1.default.chartComponent + "." + ChartControl_module_scss_1.default.themed)).to.have.length(0);
        done();
    });
});

//# sourceMappingURL=ChartControl.test.js.map
