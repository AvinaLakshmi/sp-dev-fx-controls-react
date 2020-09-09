"use strict";
/// <reference types="sinon" />
Object.defineProperty(exports, "__esModule", { value: true });
var chai_1 = require("chai");
var _1 = require(".");
var ChartControl_types_1 = require("./ChartControl.types");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
describe('PaletteGenerator', function () {
    it('Should repeat palette if array is longer than number of available colors in repeating palette', function (done) {
        var palette = _1.PaletteGenerator.GetPalette(ChartControl_types_1.ChartPalette.OfficeColorful2, 60);
        chai_1.expect(palette).to.have.length(60);
        done();
    });
    it('Should stretch palette if array is longer than number of available colors in non-repeating palette', function (done) {
        var palette = _1.PaletteGenerator.GetPalette(ChartControl_types_1.ChartPalette.OfficeMonochromatic1, 60);
        chai_1.expect(palette).to.have.length(60);
        // Array shouldn't repeat, so the unique array should be the same length as the returned array
        chai_1.expect(sp_lodash_subset_1.uniq(palette)).to.have.length(60);
        done();
    });
    it('Should return the right alpha palette length', function (done) {
        var alphaPalette = _1.PaletteGenerator.alpha(_1.PaletteGenerator.GetPalette(ChartControl_types_1.ChartPalette.OfficeColorful1, 60), 0.5);
        chai_1.expect(alphaPalette).to.have.length(60);
        done();
    });
    it('Should repeat a shorter array of colors', function (done) {
        var desiredPattern = ["#0000ff", "#00ff00", "#ff0000"];
        var palette = _1.PaletteGenerator.generateRepeatingPattern(desiredPattern, 60);
        chai_1.expect(palette).to.have.length(60);
        // Array should repeat the same exact values.
        chai_1.expect(sp_lodash_subset_1.uniq(palette)).to.have.length(3);
        done();
    });
    it('Should not repeat colors in a gradient', function (done) {
        var gradientExtremes = ["#0000ff", "#ff0000"];
        var palette = _1.PaletteGenerator.generateNonRepeatingGradient(gradientExtremes, 60);
        chai_1.expect(palette).to.have.length(60);
        // Array should not repeat
        chai_1.expect(sp_lodash_subset_1.uniq(palette)).to.have.length(60);
        done();
    });
    it('Should return an array of alpha value for a given array of colors', function (done) {
        var arrayColors = ["#00ff00", "#ff0000", "#0000ff"];
        var alphaColors = _1.PaletteGenerator.alpha(arrayColors, 0.5);
        chai_1.expect(alphaColors).to.be.an('array');
        chai_1.expect(alphaColors).to.have.length(3);
        done();
    });
    it('Should return a single alpha value for a single color', function (done) {
        var singleColor = "#00ff00";
        var alphaColor = _1.PaletteGenerator.alpha(singleColor, 0.5);
        chai_1.expect(alphaColor).to.be.a('string');
        done();
    });
    it('Should return a single alpha value for any valid type of color value', function (done) {
        chai_1.expect(_1.PaletteGenerator.alpha('red', 0.5)).to.be.a('string');
        chai_1.expect(_1.PaletteGenerator.alpha('#ff0000', 0.5)).to.be.a('string');
        chai_1.expect(_1.PaletteGenerator.alpha('rgb(255,0,0)', 0.5)).to.be.a('string');
        chai_1.expect(_1.PaletteGenerator.alpha('hsl(120,100%,50%)', 0.5)).to.be.a('string');
        done();
    });
});

//# sourceMappingURL=PaletteGenerator.test.js.map
