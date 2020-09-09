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
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
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
var React = require("react");
var ControlsTest_module_scss_1 = require("./ControlsTest.module.scss");
var FileTypeIcon_1 = require("../../../FileTypeIcon");
var Dropdown_1 = require("office-ui-fabric-react/lib/components/Dropdown");
var Button_1 = require("office-ui-fabric-react/lib/components/Button");
var Dialog_1 = require("office-ui-fabric-react/lib/components/Dialog");
var Placeholder_1 = require("../../../Placeholder");
var ListView_1 = require("../../../ListView");
var sp_http_1 = require("@microsoft/sp-http");
var SiteBreadcrumb_1 = require("../../../SiteBreadcrumb");
var WebPartTitle_1 = require("../../../WebPartTitle");
var TaxonomyPicker_1 = require("../../../TaxonomyPicker");
var ListPicker_1 = require("../../../ListPicker");
var IFrameDialog_1 = require("../../../IFrameDialog");
var IFramePanel_1 = require("../../../IFramePanel");
var Panel_1 = require("office-ui-fabric-react/lib/Panel");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var SecurityTrimmedControl_1 = require("../../../SecurityTrimmedControl");
var sp_page_context_1 = require("@microsoft/sp-page-context");
var PeoplePicker_1 = require("../../../PeoplePicker");
var DateValues_1 = require("office-ui-fabric-react/lib/utilities/dateValues/DateValues");
var DateTimePicker_1 = require("../../../DateTimePicker");
var ListItemPicker_1 = require("../../../ListItemPicker");
var Map_1 = require("../../../Map");
var ChartControl_1 = require("../../../ChartControl");
var Progress_1 = require("../../../Progress");
var taxonomyPicker_1 = require("../../../controls/taxonomyPicker");
var termActions_1 = require("../../../controls/taxonomyPicker/termActions");
var RichText_1 = require("../../../RichText");
var Link_1 = require("office-ui-fabric-react/lib/components/Link");
var carousel_1 = require("../../../controls/carousel");
var TimeDisplayControlType_1 = require("../../../controls/dateTimePicker/TimeDisplayControlType");
var GridLayout_1 = require("../../../GridLayout");
var ComboBoxListItemPicker_1 = require("../../../controls/listItemPicker/ComboBoxListItemPicker");
var treeView_1 = require("../../../controls/treeView");
var iconPicker_1 = require("../../../controls/iconPicker");
// Used to render document cards
var DocumentCard_1 = require("office-ui-fabric-react/lib/DocumentCard");
var Image_1 = require("office-ui-fabric-react/lib/Image");
var FilePicker_1 = require("../../../FilePicker");
var FolderPicker_1 = require("../../../controls/folderPicker/FolderPicker");
var FolderExplorer_1 = require("../../../FolderExplorer");
var pagination_1 = require("../../../controls/pagination");
var FieldCollectionData_1 = require("../../../FieldCollectionData");
var __1 = require("../../..");
/**
 * The sample data below was randomly generated (except for the title). It is used by the grid layout
 */
var sampleGridData = [{
        thumbnail: "https://pixabay.com/get/57e9dd474952a414f1dc8460825668204022dfe05555754d742e7bd6/hot-air-balloons-1984308_640.jpg",
        title: "Adventures in SPFx",
        name: "Perry Losselyong",
        profileImageSrc: "https://robohash.org/blanditiisadlabore.png?size=50x50&set=set1",
        location: "SharePoint",
        activity: "3/13/2019"
    }, {
        thumbnail: "https://pixabay.com/get/55e8d5474a52ad14f1dc8460825668204022dfe05555754d742d79d0/autumn-3804001_640.jpg",
        title: "The Wild, Untold Story of SharePoint!",
        name: "Ebonee Gallyhaock",
        profileImageSrc: "https://robohash.org/delectusetcorporis.bmp?size=50x50&set=set1",
        location: "SharePoint",
        activity: "6/29/2019"
    }, {
        thumbnail: "https://pixabay.com/get/57e8dd454c50ac14f1dc8460825668204022dfe05555754d742c72d7/log-cabin-1886620_640.jpg",
        title: "Low Code Solutions: PowerApps",
        name: "Seward Keith",
        profileImageSrc: "https://robohash.org/asperioresautquasi.jpg?size=50x50&set=set1",
        location: "PowerApps",
        activity: "12/31/2018"
    }, {
        thumbnail: "https://pixabay.com/get/55e3d445495aa514f1dc8460825668204022dfe05555754d742b7dd5/portrait-3316389_640.jpg",
        title: "Not Your Grandpa's SharePoint",
        name: "Sharona Selkirk",
        profileImageSrc: "https://robohash.org/velnammolestiae.png?size=50x50&set=set1",
        location: "SharePoint",
        activity: "11/20/2018"
    }, {
        thumbnail: "https://pixabay.com/get/57e6dd474352ae14f1dc8460825668204022dfe05555754d742a7ed1/faucet-1684902_640.jpg",
        title: "Get with the Flow",
        name: "Boyce Batstone",
        profileImageSrc: "https://robohash.org/nulladistinctiomollitia.jpg?size=50x50&set=set1",
        location: "Flow",
        activity: "5/26/2019"
    }];
var sampleItems = [
    {
        Langue: { Nom: 'Français' },
        Question: 'Charger des fichiers et dossiers',
        Reponse: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    },
    {
        Langue: { Nom: 'Français' },
        Question: 'Enregistrer un fichier',
        Reponse: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    },
    {
        Langue: { Nom: 'Français' },
        Question: 'Troisième exemple',
        Reponse: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    },
    {
        Langue: { Nom: 'Français' },
        Question: 'Quatrième exemple',
        Reponse: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    },
    {
        Langue: { Nom: 'Français' },
        Question: 'Cinquième exemple',
        Reponse: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    },
    {
        Langue: { Nom: 'Français' },
        Question: 'Sixième exemple',
        Reponse: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
    }
];
/**
 * Component that can be used to test out the React controls from this project
 */
var ControlsTest = (function (_super) {
    __extends(ControlsTest, _super);
    function ControlsTest(props) {
        var _this = _super.call(this, props) || this;
        _this.taxService = null;
        _this.richTextValue = null;
        /**
         * Static array for carousel control example.
         */
        _this.carouselElements = [
            React.createElement("div", { id: "1", key: "1" }, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis a mattis libero, nec consectetur neque. Suspendisse potenti. Fusce ultrices faucibus consequat. Suspendisse ex diam, ullamcorper sit amet justo ac, accumsan congue neque. Vestibulum aliquam mauris non justo convallis, id molestie purus sodales. Maecenas scelerisque aliquet turpis, ac efficitur ex iaculis et. Vivamus finibus mi eget urna tempor, sed porta justo tempus. Vestibulum et lectus magna. Integer ante felis, ullamcorper venenatis lectus ac, vulputate pharetra magna. Morbi eget nisl tempus, viverra diam ac, mollis tortor. Nam odio ex, viverra bibendum mauris vehicula, consequat suscipit ligula. Nunc sed ultrices augue, eu tincidunt diam."),
            React.createElement("div", { id: "2", key: "2" }, "Quisque metus lectus, facilisis id consectetur ac, hendrerit eget quam. Interdum et malesuada fames ac ante ipsum primis in faucibus. Ut faucibus posuere felis vel efficitur. Maecenas et massa in sem tincidunt finibus. Duis sit amet bibendum nisi. Vestibulum pretium pretium libero, vel tincidunt sem vestibulum sed. Interdum et malesuada fames ac ante ipsum primis in faucibus. Proin quam lorem, venenatis id bibendum id, tempus eu nibh. Sed tristique semper ligula, vitae gravida diam gravida vitae. Donec eget posuere mauris, pharetra semper lectus."),
            React.createElement("div", { id: "3", key: "3" }, "Pellentesque tempor et leo at tincidunt. Vivamus et leo sed eros vehicula mollis vitae in dui. Duis posuere sodales enim ut ultricies. Cras in venenatis nulla. Ut sed neque dignissim, sollicitudin tellus convallis, placerat leo. Aliquam vestibulum, leo pharetra sollicitudin pretium, ipsum nisl tincidunt orci, in molestie ipsum dui et mi. Praesent aliquam accumsan risus sed bibendum. Cras consectetur elementum turpis, a mollis velit gravida sit amet. Praesent non augue cursus, varius justo at, molestie lorem. Nulla cursus tellus quis odio congue elementum. Vivamus sit amet quam nec lectus hendrerit blandit. Duis ac condimentum sem. Morbi hendrerit elementum purus, non facilisis arcu bibendum vitae. Vivamus commodo tristique euismod."),
            React.createElement("div", { id: "4", key: "4" }, "Proin semper egestas porta. Nullam risus nisl, auctor ac hendrerit in, dapibus quis ex. Quisque vitae nisi quam. Etiam vel sapien ut libero ornare rhoncus nec vestibulum dolor. Curabitur lacinia aliquam arcu. Proin ultrices risus velit, in vehicula tellus vehicula at. Sed ultrices et felis fringilla ultricies."),
            React.createElement("div", { id: "5", key: "5" }, "Donec orci lorem, imperdiet eu nisi sit amet, condimentum scelerisque tortor. Etiam nec lacinia dui. Duis non turpis neque. Sed pellentesque a erat et accumsan. Pellentesque elit odio, elementum nec placerat nec, ornare in tortor. Suspendisse gravida magna maximus mollis facilisis. Duis odio libero, finibus ac suscipit sed, aliquam et diam. Aenean posuere lacus ex. Donec dapibus, sem ac luctus ultrices, justo libero tempor eros, vitae lacinia ex ante non dolor. Curabitur condimentum, ligula id pharetra dictum, libero libero ullamcorper nunc, eu blandit sem arcu ut felis. Nullam lacinia dapibus auctor.")
        ];
        _this.skypeCheckIcon = { iconName: 'SkypeCheck' };
        _this.treeitems = [
            {
                key: "R1",
                label: "Root",
                subLabel: "This is a sub label for node",
                iconProps: _this.skypeCheckIcon,
                actions: [{
                        title: "Get item",
                        iconProps: {
                            iconName: 'Warning',
                            style: {
                                color: 'salmon',
                            },
                        },
                        id: "GetItem",
                        actionCallback: function (treeItem) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                console.log(treeItem);
                                return [2 /*return*/];
                            });
                        }); }
                    }],
                children: [
                    {
                        key: "1",
                        label: "Parent 1",
                        selectable: false,
                        children: [
                            {
                                key: "3",
                                label: "Child 1",
                                subLabel: "This is a sub label for node",
                                actions: [{
                                        title: "Share",
                                        iconProps: {
                                            iconName: 'Share'
                                        },
                                        id: "GetItem",
                                        actionCallback: function (treeItem) { return __awaiter(_this, void 0, void 0, function () {
                                            return __generator(this, function (_a) {
                                                console.log(treeItem);
                                                return [2 /*return*/];
                                            });
                                        }); }
                                    }],
                                children: [
                                    {
                                        key: "gc1",
                                        label: "Grand Child 1",
                                        actions: [{
                                                title: "Get Grand Child item",
                                                iconProps: {
                                                    iconName: 'Mail'
                                                },
                                                id: "GetItem",
                                                actionCallback: function (treeItem) { return __awaiter(_this, void 0, void 0, function () {
                                                    return __generator(this, function (_a) {
                                                        console.log(treeItem);
                                                        return [2 /*return*/];
                                                    });
                                                }); }
                                            }]
                                    }
                                ]
                            },
                            {
                                key: "4",
                                label: "Child 2",
                                iconProps: _this.skypeCheckIcon
                            }
                        ]
                    },
                    {
                        key: "2",
                        label: "Parent 2"
                    },
                    {
                        key: "5",
                        label: "Parent 3",
                        disabled: true
                    },
                    {
                        key: "6",
                        label: "Parent 4",
                        selectable: true
                    }
                ]
            },
            {
                key: "R2",
                label: "Root 2",
                children: [
                    {
                        key: "8",
                        label: "Parent 5"
                    },
                    {
                        key: "9",
                        label: "Parent 6"
                    },
                    {
                        key: "10",
                        label: "Parent 7"
                    },
                    {
                        key: "11",
                        label: "Parent 8"
                    }
                ]
            },
            {
                key: "R3",
                label: "Root 3",
                children: [
                    {
                        key: "12",
                        label: "Parent 9"
                    },
                    {
                        key: "13",
                        label: "Parent 10",
                        children: [
                            {
                                key: "gc3",
                                label: "Child of Parent 10"
                            },
                        ]
                    },
                    {
                        key: "14",
                        label: "Parent 11"
                    },
                    {
                        key: "15",
                        label: "Parent 12"
                    }
                ]
            }
        ];
        /**
         *
         *Method that retrieves the selected terms from the taxonomy picker and sets state
         * @private
         * @param {IPickerTerms} terms
         * @memberof ControlsTest
         */
        _this.onServicePickerChange = function (terms) {
            _this.setState({
                initialValues: terms
            });
            // console.log("serviceTerms", terms);
        };
        /**
         * Method that retrieves the selected terms from the taxonomy picker
         * @param terms
         */
        _this._onTaxPickerChange = function (terms) {
            _this.setState({
                initialValues: terms
            });
            console.log("Terms:", terms);
        };
        /**
         * Method that retrieves the selected date/time from the DateTime picker
         * @param dateTimeValue
         */
        _this._onDateTimePickerChange = function (dateTimeValue) {
            _this.setState({ dateTimeValue: dateTimeValue });
            console.log("Selected Date/Time:", dateTimeValue.toLocaleString());
        };
        /**
         * Selected lists change event
         * @param lists
         */
        _this.onListPickerChange = function (lists) {
            console.log("Lists:", lists);
            _this.setState({
                selectedList: typeof lists === "string" ? lists : lists.pop()
            });
        };
        /**
         * Deletes second item from the list
         */
        _this.deleteItem = function () {
            var items = _this.state.items;
            if (items.length >= 2) {
                items.splice(1, 1);
                _this.setState({
                    items: items
                });
            }
        };
        /**
         * Triggers element change for the carousel example.
         */
        _this.triggerNextElement = function (index) {
            var canMovePrev = index > 0;
            var canMoveNext = index < _this.carouselElements.length - 1;
            var nextElement = _this.carouselElements[index];
            setTimeout(function () {
                _this.setState({
                    canMovePrev: canMovePrev,
                    canMoveNext: canMoveNext,
                    currentCarouselElement: nextElement
                });
            }, 500);
        };
        _this._onFilePickerSave = function (filePickerResult) { return __awaiter(_this, void 0, void 0, function () {
            var fileResultContent;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ filePickerResult: filePickerResult });
                        if (!filePickerResult) return [3 /*break*/, 2];
                        return [4 /*yield*/, filePickerResult.downloadFileContent()];
                    case 1:
                        fileResultContent = _a.sent();
                        console.log(fileResultContent);
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        }); };
        _this._onFolderSelect = function (folder) {
            console.log('selected folder', folder);
        };
        _this._onRenderGridItem = function (item, _finalSize, isCompact) {
            var previewProps = {
                previewImages: [
                    {
                        previewImageSrc: item.thumbnail,
                        imageFit: Image_1.ImageFit.cover,
                        height: 130
                    }
                ]
            };
            return React.createElement("div", { "data-is-focusable": true, role: "listitem", "aria-label": item.title },
                React.createElement(DocumentCard_1.DocumentCard, { type: isCompact ? DocumentCard_1.DocumentCardType.compact : DocumentCard_1.DocumentCardType.normal, onClick: function (ev) { return alert("You clicked on a grid item"); } },
                    React.createElement(DocumentCard_1.DocumentCardPreview, __assign({}, previewProps)),
                    !isCompact && React.createElement(DocumentCard_1.DocumentCardLocation, { location: item.location }),
                    React.createElement("div", null,
                        React.createElement(DocumentCard_1.DocumentCardTitle, { title: item.title, shouldTruncate: true }),
                        React.createElement(DocumentCard_1.DocumentCardActivity, { activity: item.activity, people: [{ name: item.name, profileImageSrc: item.profileImageSrc }] }))));
        };
        _this.getRandomCollectionFieldData = function () {
            var result = [];
            for (var i = 1; i < 16; i++) {
                result.push({ "Field1": "String" + i, "Field2": i, "Field3": "https://pnp.github.io/", "Field4": true });
            }
            return result;
        };
        _this.state = {
            imgSize: FileTypeIcon_1.ImageSize.small,
            items: [],
            iFrameDialogOpened: false,
            iFramePanelOpened: false,
            initialValues: [],
            authorEmails: [],
            selectedList: null,
            progressActions: _this._initProgressActions(),
            dateTimeValue: new Date(),
            richTextValue: null,
            canMovePrev: false,
            canMoveNext: true,
            currentCarouselElement: _this.carouselElements[0]
        };
        _this._onIconSizeChange = _this._onIconSizeChange.bind(_this);
        _this._onConfigure = _this._onConfigure.bind(_this);
        _this._startProgress = _this._startProgress.bind(_this);
        return _this;
    }
    /**
     * React componentDidMount lifecycle hook
     */
    ControlsTest.prototype.componentDidMount = function () {
        var _this = this;
        var restApi = this.props.context.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl('Shared%20Documents')/files?$expand=ListItemAllFields";
        this.props.context.spHttpClient.get(restApi, sp_http_1.SPHttpClient.configurations.v1)
            .then(function (resp) { return resp.json(); })
            .then(function (items) {
            _this.setState({
                items: items.value ? items.value : []
            });
        });
        // // Get Authors in the SharePoint Document library -- For People Picker Testing
        // const restAuthorApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Documents')/Items?$select=Id, Author/EMail&$expand=Author/EMail`;
        // this.props.context.spHttpClient.get(restAuthorApi, SPHttpClient.configurations.v1)
        // .then(resp => { return resp.json(); })
        // .then(items => {
        //   let emails : string[] = items.value ? items.value.map((item, key)=> { return item.Author.EMail}) : [];
        //   console.log(emails);
        //   this.setState({
        //     authorEmails: emails
        //   });
        // });
    };
    /**
     * Event handler when changing the icon size in the dropdown
     * @param element
     */
    ControlsTest.prototype._onIconSizeChange = function (element) {
        this.setState({
            imgSize: parseInt(element.key.toString())
        });
    };
    /**
     * Open the property pane
     */
    ControlsTest.prototype._onConfigure = function () {
        this.props.context.propertyPane.open();
    };
    /**
     * Method that retrieves the selected items in the list view
     * @param items
     */
    ControlsTest.prototype._getSelection = function (items) {
        console.log('Items:', items);
    };
    /**
     * Method that retrieves the selected items from People  Picker
     * @param items
     */
    ControlsTest.prototype._getPeoplePickerItems = function (items) {
        console.log('Items:', items);
    };
    /**
     * Selected item from the list data picker
     */
    ControlsTest.prototype.listItemPickerDataSelected = function (item) {
        console.log(item);
    };
    ControlsTest.prototype._startProgress = function () {
        var _this = this;
        var currentIndex = 0;
        var intervalId = setInterval(function () {
            var actions = _this.state.progressActions;
            if (currentIndex >= actions.length) {
                clearInterval(intervalId);
            }
            else {
                var action = actions[currentIndex];
                if (currentIndex == 1) {
                    action.hasError = true;
                    action.errorMessage = 'some error message';
                }
            }
            _this.setState({
                currentProgressActionIndex: currentIndex,
                progressActions: actions
            });
            currentIndex++;
        }, 5000);
    };
    ControlsTest.prototype._initProgressActions = function () {
        return [{
                title: 'First Step',
                subActionsTitles: [
                    'Sub action 1',
                    'Sub action 2'
                ]
            }, {
                title: 'Second step'
            }, {
                title: 'Third Step',
                subActionsTitles: [
                    'Sub action 1',
                    'Sub action 2',
                    'Sub action 3'
                ]
            }, {
                title: 'Fourth Step'
            }];
    };
    /**
     * Renders the component
     */
    ControlsTest.prototype.render = function () {
        var _this = this;
        // Size options for the icon size dropdown
        var sizeOptions = [
            {
                key: FileTypeIcon_1.ImageSize.small,
                text: FileTypeIcon_1.ImageSize[FileTypeIcon_1.ImageSize.small],
                selected: FileTypeIcon_1.ImageSize.small === this.state.imgSize
            },
            {
                key: FileTypeIcon_1.ImageSize.medium,
                text: FileTypeIcon_1.ImageSize[FileTypeIcon_1.ImageSize.medium],
                selected: FileTypeIcon_1.ImageSize.medium === this.state.imgSize
            },
            {
                key: FileTypeIcon_1.ImageSize.large,
                text: FileTypeIcon_1.ImageSize[FileTypeIcon_1.ImageSize.large],
                selected: FileTypeIcon_1.ImageSize.large === this.state.imgSize
            }
        ];
        // Specify the fields that need to be viewed in the listview
        var viewFields = [
            {
                name: 'ListItemAllFields.Id',
                displayName: 'ID',
                maxWidth: 40,
                sorting: true,
                isResizable: true
            },
            {
                name: 'ListItemAllFields.Underscore_Field',
                displayName: "Underscore_Field",
                sorting: true,
                isResizable: true
            },
            {
                name: 'Name',
                linkPropertyName: 'ServerRelativeUrl',
                sorting: true,
                isResizable: true
            },
            {
                name: 'ServerRelativeUrl',
                displayName: 'Path',
                render: function (item) {
                    return React.createElement("a", { href: item['ServerRelativeUrl'] }, "Link");
                },
                isResizable: true
            },
            {
                name: 'Title',
                isResizable: true
            }
        ];
        // Specify the fields on which you want to group your items
        // Grouping is takes the field order into account from the array
        // const groupByFields: IGrouping[] = [{ name: "ListItemAllFields.City", order: GroupOrder.ascending }, { name: "ListItemAllFields.Country.Label", order: GroupOrder.descending }];
        var groupByFields = [{ name: "ListItemAllFields.Department.Label", order: ListView_1.GroupOrder.ascending }];
        var iframeUrl = '/temp/workbench.html';
        if (sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.SharePoint) {
            iframeUrl = '/_layouts/15/sharepoint.aspx';
        }
        else if (sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.ClassicSharePoint) {
            iframeUrl = this.context.pageContext.web.serverRelativeUrl;
        }
        var additionalBreadcrumbItems = [{
                text: 'Places', key: 'Places', onClick: function () {
                    console.log('additional breadcrumb item');
                },
            }];
        return (React.createElement("div", { className: ControlsTest_module_scss_1.default.controlsTest },
            React.createElement(WebPartTitle_1.WebPartTitle, { displayMode: this.props.displayMode, title: this.props.title, updateProperty: this.props.updateProperty, moreLink: React.createElement(Link_1.Link, { href: "https://pnp.github.io/sp-dev-fx-controls-react/" }, "See all") }),
            sampleItems.map(function (item, index) { return (React.createElement(__1.Accordion, { title: item.Question, defaultCollapsed: false, className: "itemCell", key: index },
                React.createElement("div", { className: "itemContent" },
                    React.createElement("div", { className: "itemResponse" }, item.Reponse),
                    React.createElement("div", { className: "itemIndex" }, "Langue :  " + item.Langue.Nom)))); }),
            React.createElement("div", { className: "ms-font-m" },
                "Services tester:",
                React.createElement(TaxonomyPicker_1.TaxonomyPicker, { allowMultipleSelections: true, 
                    //termsetNameOrID="61837936-29c5-46de-982c-d1adb6664b32" // id to termset that has a custom sort
                    termsetNameOrID: "8ea5ac06-fd7c-4269-8d0d-02f541df8eb9", initialValues: [{
                            key: "c05250ff-80e7-41e6-bfb3-db2db62d63d3",
                            name: "Business",
                            path: "Business",
                            termSet: "8ea5ac06-fd7c-4269-8d0d-02f541df8eb9",
                            termSetName: "Trip Types"
                        }, {
                            key: "a05250ff-80e7-41e6-bfb3-db2db62d63d3",
                            name: "BBusiness",
                            path: "BBusiness",
                            termSet: "8ea5ac06-fd7c-4269-8d0d-02f541df8eb9",
                            termSetName: "Trip Types"
                        }], validateOnLoad: true, panelTitle: "Select Sorted Term", label: "Service Picker with custom actions", context: this.props.context, onChange: this.onServicePickerChange, isTermSetSelectable: false, termActions: {
                        actions: [{
                                title: "Get term labels",
                                iconName: "LocaleLanguage",
                                id: "test",
                                invokeActionOnRender: true,
                                hidden: true,
                                actionCallback: function (taxService, term) { return __awaiter(_this, void 0, void 0, function () {
                                    var updateAction;
                                    return __generator(this, function (_a) {
                                        updateAction = {
                                            updateActionType: TaxonomyPicker_1.UpdateType.updateTermLabel,
                                            value: term.Name + " (updated)"
                                        };
                                        return [2 /*return*/, updateAction];
                                    });
                                }); },
                                applyToTerm: function (term) { return (term && term.Name && term.Name.toLowerCase() === "about us"); }
                            },
                        ],
                        termActionsDisplayMode: termActions_1.TermActionsDisplayMode.buttons,
                        termActionsDisplayStyle: taxonomyPicker_1.TermActionsDisplayStyle.textAndIcon
                    } }),
                React.createElement(TaxonomyPicker_1.TaxonomyPicker, { allowMultipleSelections: true, termsetNameOrID: "8ea5ac06-fd7c-4269-8d0d-02f541df8eb9" // id to termset that has a default sort
                    , panelTitle: "Select Default Sorted Term", label: "Service Picker", context: this.props.context, onChange: this.onServicePickerChange, isTermSetSelectable: false, placeholder: "Select service", required: true, errorMessage: 'this field is required', onGetErrorMessage: function (value) { return 'comment errorMessage to see this one'; } }),
                React.createElement(TaxonomyPicker_1.TaxonomyPicker, { initialValues: this.state.initialValues, allowMultipleSelections: true, termsetNameOrID: "41dec50a-3e09-4b3f-842a-7224cffc74c0", anchorId: "436a6154-9691-4925-baa5-4c9bb9212cbf", 
                    // disabledTermIds={["943fd9f0-3d7c-415c-9192-93c0e54573fb", "0e415292-cce5-44ac-87c7-ef99dd1f01f4"]}
                    // disabledTermIds={["943fd9f0-3d7c-415c-9192-93c0e54573fb", "73d18756-20af-41de-808c-2a1e21851e44", "0e415292-cce5-44ac-87c7-ef99dd1f01f4"]}
                    // disabledTermIds={["cd6f6d3c-672d-4244-9320-c1e64cc0626f", "0e415292-cce5-44ac-87c7-ef99dd1f01f4"]}
                    // disableChildrenOfDisabledParents={true}
                    panelTitle: "Select Term", label: "Taxonomy Picker", context: this.props.context, onChange: this._onTaxPickerChange, isTermSetSelectable: false, hideDeprecatedTags: true, hideTagsNotAvailableForTagging: true, termActions: {
                        actions: [{
                                title: "Get term labels",
                                iconName: "LocaleLanguage",
                                id: "test",
                                invokeActionOnRender: true,
                                hidden: true,
                                actionCallback: function (taxService, term) { return __awaiter(_this, void 0, void 0, function () {
                                    return __generator(this, function (_a) {
                                        console.log(term.Name, term.TermsCount);
                                        return [2 /*return*/, {
                                                updateActionType: TaxonomyPicker_1.UpdateType.updateTermLabel,
                                                value: term.Name + " (updated)"
                                            }];
                                    });
                                }); },
                                applyToTerm: function (term) { return (term && term.Name && term.Name === "internal"); }
                            },
                            {
                                title: "Hide term",
                                id: "hideTerm",
                                invokeActionOnRender: true,
                                hidden: true,
                                actionCallback: function (taxService, term) { return __awaiter(_this, void 0, void 0, function () {
                                    return __generator(this, function (_a) {
                                        return [2 /*return*/, {
                                                updateActionType: TaxonomyPicker_1.UpdateType.hideTerm,
                                                value: true
                                            }];
                                    });
                                }); },
                                applyToTerm: function (term) { return (term && term.Name && (term.Name.toLowerCase() === "help desk" || term.Name.toLowerCase() === "multi-column valo site page")); }
                            },
                            {
                                title: "Disable term",
                                id: "disableTerm",
                                invokeActionOnRender: true,
                                hidden: true,
                                actionCallback: function (taxService, term) { return __awaiter(_this, void 0, void 0, function () {
                                    return __generator(this, function (_a) {
                                        return [2 /*return*/, {
                                                updateActionType: TaxonomyPicker_1.UpdateType.disableTerm,
                                                value: true
                                            }];
                                    });
                                }); },
                                applyToTerm: function (term) { return (term && term.Name && term.Name.toLowerCase() === "secured"); }
                            },
                            {
                                title: "Disable or hide term",
                                id: "disableOrHideTerm",
                                invokeActionOnRender: true,
                                hidden: true,
                                actionCallback: function (taxService, term) { return __awaiter(_this, void 0, void 0, function () {
                                    return __generator(this, function (_a) {
                                        if (term.TermsCount > 0) {
                                            return [2 /*return*/, {
                                                    updateActionType: TaxonomyPicker_1.UpdateType.disableTerm,
                                                    value: true
                                                }];
                                        }
                                        return [2 /*return*/, {
                                                updateActionType: TaxonomyPicker_1.UpdateType.hideTerm,
                                                value: true
                                            }];
                                    });
                                }); },
                                applyToTerm: function (term) { return true; }
                            }],
                        termActionsDisplayMode: termActions_1.TermActionsDisplayMode.buttons,
                        termActionsDisplayStyle: taxonomyPicker_1.TermActionsDisplayStyle.textAndIcon
                    } }),
                React.createElement(Button_1.DefaultButton, { text: "Add", onClick: function () {
                        _this.setState({
                            initialValues: [{
                                    key: "ab703558-2546-4b23-b8b8-2bcb2c0086f5",
                                    name: "HR",
                                    path: "HR",
                                    termSet: "b3e9b754-2593-4ae6-abc2-35345402e186"
                                }]
                        });
                    } })),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (unspecified = date and time)", isMonthPickerVisible: false, showSeconds: false, onChange: function (value) { return console.log("DateTimePicker value:", value); }, placeholder: "Pick a date" }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker 12-hour clock", showSeconds: true, onChange: function (value) { return console.log("DateTimePicker value:", value); } }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker 24-hour clock", showSeconds: true, timeConvention: DateTimePicker_1.TimeConvention.Hours24, onChange: function (value) { return console.log("DateTimePicker value:", value); } }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker no seconds", value: new Date(), onChange: function (value) { return console.log("DateTimePicker value:", value); } }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (unspecified = date and time)", timeConvention: DateTimePicker_1.TimeConvention.Hours24, value: new Date(), onChange: function (value) { return console.log("DateTimePicker value:", value); } }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker dropdown", showSeconds: true, timeDisplayControlType: TimeDisplayControlType_1.TimeDisplayControlType.Dropdown, value: new Date(), onChange: function (value) { return console.log("DateTimePicker value:", value); } }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker date only", showLabels: false, dateConvention: DateTimePicker_1.DateConvention.Date, value: new Date(), onChange: function (value) { return console.log("DateTimePicker value:", value); }, minDate: new Date("05/01/2019"), maxDate: new Date("05/01/2020") }),
            React.createElement(RichText_1.RichText, { isEditMode: this.props.displayMode === sp_core_library_1.DisplayMode.Edit, onChange: function (value) { _this.setState({ richTextValue: value }); return value; } }),
            React.createElement(Placeholder_1.Placeholder, { iconName: 'Edit', iconText: 'Configure your web part', description: 'Please configure the web part.', buttonLabel: 'Configure', hideButton: this.props.displayMode === sp_core_library_1.DisplayMode.Read, onConfigure: this._onConfigure }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (Group not found)", webAbsoluteUrl: this.props.context.pageContext.site.absoluteUrl, groupName: "Team Site Visitors 123", ensureUser: true, principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.DistributionList], defaultSelectedUsers: ["admin@tenant.onmicrosoft.com", "test@tenant.onmicrosoft.com"], selectedItems: this._getPeoplePickerItems }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (search for group)", groupName: "Team Site Visitors", principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.DistributionList], defaultSelectedUsers: ["admin@tenant.onmicrosoft.com", "test@tenant.onmicrosoft.com"], selectedItems: this._getPeoplePickerItems }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (pre-set global users)", principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.DistributionList], defaultSelectedUsers: ["admin@tenant.onmicrosoft.com", "test@tenant.onmicrosoft.com"], selectedItems: this._getPeoplePickerItems, personSelectionLimit: 2, ensureUser: true }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (pre-set local users)", webAbsoluteUrl: this.props.context.pageContext.site.absoluteUrl, principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.DistributionList], defaultSelectedUsers: ["admin@tenant.onmicrosoft.com", "test@tenant.onmicrosoft.com"], selectedItems: this._getPeoplePickerItems }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (tenant scoped)", personSelectionLimit: 5, 
                // groupName={"Team Site Owners"}
                showtooltip: true, isRequired: true, 
                //defaultSelectedUsers={["tenantUser@domain.onmicrosoft.com", "test@user.com"]}
                //defaultSelectedUsers={this.state.authorEmails}
                selectedItems: this._getPeoplePickerItems, showHiddenInUI: false, principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.DistributionList], suggestionsLimit: 2, resolveDelay: 200, placeholder: 'Select a SharePoint principal (User or Group)' }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (local scoped)", webAbsoluteUrl: this.props.context.pageContext.site.absoluteUrl, personSelectionLimit: 5, 
                // groupName={"Team Site Owners"}
                showtooltip: true, isRequired: true, 
                //defaultSelectedUsers={["tenantUser@domain.onmicrosoft.com", "test@user.com"]}
                //defaultSelectedUsers={this.state.authorEmails}
                selectedItems: this._getPeoplePickerItems, showHiddenInUI: false, principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.DistributionList], suggestionsLimit: 2, resolveDelay: 200 }),
            React.createElement(PeoplePicker_1.PeoplePicker, { context: this.props.context, titleText: "People Picker (disabled)", disabled: true, showtooltip: true, defaultSelectedUsers: ['aleksei.dovzhyk@sharepointalist.com'] }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (unspecified = date and time)" }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (unspecified = date and time, no seconds)" }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (date and time - default time = 12h)", dateConvention: DateTimePicker_1.DateConvention.DateTime, showSeconds: true }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (date and time - 12h)", dateConvention: DateTimePicker_1.DateConvention.DateTime, timeConvention: DateTimePicker_1.TimeConvention.Hours12, showSeconds: false }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (date and time - 24h)", dateConvention: DateTimePicker_1.DateConvention.DateTime, timeConvention: DateTimePicker_1.TimeConvention.Hours24, firstDayOfWeek: DateValues_1.DayOfWeek.Monday, showSeconds: true }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (Controlled)", formatDate: function (d) { return d.getFullYear() + " - " + (d.getMonth() + 1) + " - " + d.getDate(); }, dateConvention: DateTimePicker_1.DateConvention.DateTime, timeConvention: DateTimePicker_1.TimeConvention.Hours24, firstDayOfWeek: DateValues_1.DayOfWeek.Monday, value: this.state.dateTimeValue, onChange: this._onDateTimePickerChange, isMonthPickerVisible: false, showMonthPickerAsOverlay: true, showWeekNumbers: true, showSeconds: true }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (date only)", dateConvention: DateTimePicker_1.DateConvention.Date }),
            React.createElement(DateTimePicker_1.DateTimePicker, { label: "DateTime Picker (disabled)", disabled: true }),
            React.createElement(ListView_1.ListView, { items: this.state.items, viewFields: viewFields, iconFieldName: 'ServerRelativeUrl', groupByFields: groupByFields, compact: true, selectionMode: ListView_1.SelectionMode.single, selection: this._getSelection, showFilter: true }),
            React.createElement(ChartControl_1.ChartControl, { type: ChartControl_1.ChartType.Bar, data: {
                    labels: ["Red", "Blue", "Yellow", "Green", "Purple", "Orange"],
                    datasets: [{
                            label: '# of Votes',
                            data: [12, 19, 3, 5, 2, 3],
                            backgroundColor: [
                                'rgba(255, 99, 132, 0.2)',
                                'rgba(54, 162, 235, 0.2)',
                                'rgba(255, 206, 86, 0.2)',
                                'rgba(75, 192, 192, 0.2)',
                                'rgba(153, 102, 255, 0.2)',
                                'rgba(255, 159, 64, 0.2)'
                            ],
                            borderColor: [
                                'rgba(255,99,132,1)',
                                'rgba(54, 162, 235, 1)',
                                'rgba(255, 206, 86, 1)',
                                'rgba(75, 192, 192, 1)',
                                'rgba(153, 102, 255, 1)',
                                'rgba(255, 159, 64, 1)'
                            ],
                            borderWidth: 1
                        }]
                }, options: {
                    scales: {
                        yAxes: [{
                                ticks: {
                                    beginAtZero: true
                                }
                            }]
                    }
                } }),
            React.createElement(Map_1.Map, { titleText: "New map control", coordinates: { latitude: 51.507351, longitude: -0.127758 }, enableSearch: true, mapType: Map_1.MapType.normal, onUpdateCoordinates: function (coordinates) { return console.log("Updated location:", coordinates); } }),
            React.createElement("div", { className: ControlsTest_module_scss_1.default.container },
                React.createElement("div", { className: "ms-Grid-row ms-bgColor-neutralLight ms-fontColor-neutralDark " + ControlsTest_module_scss_1.default.row },
                    React.createElement("div", { className: "ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1" },
                        React.createElement("span", { className: "ms-font-xl" }, "Controls testing"),
                        React.createElement(SecurityTrimmedControl_1.SecurityTrimmedControl, { context: this.props.context, level: SecurityTrimmedControl_1.PermissionLevel.currentWeb, permissions: [sp_page_context_1.SPPermission.viewListItems], className: "TestingClass", noPermissionsControl: React.createElement("p", null, "You do not have permissions.") },
                            React.createElement("p", null, "You have permissions to view list items.")),
                        React.createElement("p", { className: "ms-font-l" }, "File type icon control"),
                        React.createElement("div", { className: "ms-font-m" },
                            "Font icons:\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.font, path: "https://contoso.sharepoint.com/documents/filename.docx" }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.font, path: "https://contoso.sharepoint.com/documents/filename.unknown" }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.font, path: "https://contoso.sharepoint.com/documents/filename.doc" }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.font, application: FileTypeIcon_1.ApplicationType.HTML }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.font, application: FileTypeIcon_1.ApplicationType.Mail }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.font, application: FileTypeIcon_1.ApplicationType.SASS })),
                        React.createElement("div", { className: "ms-font-m" },
                            "Image icons:\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.docx" }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.unknown" }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.pptx?querystring='prop1'&prop2='test'" }),
                            " \u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, application: FileTypeIcon_1.ApplicationType.Word }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, application: FileTypeIcon_1.ApplicationType.PDF }),
                            "\u00A0",
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, path: "https://contoso.sharepoint.com/documents/filename.pdf" })),
                        React.createElement("div", { className: "ms-font-m" },
                            "Icon size tester:",
                            React.createElement(Dropdown_1.Dropdown, { options: sizeOptions, onChanged: this._onIconSizeChange }),
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, size: this.state.imgSize, application: FileTypeIcon_1.ApplicationType.Excel }),
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, size: this.state.imgSize, application: FileTypeIcon_1.ApplicationType.PDF }),
                            React.createElement(FileTypeIcon_1.FileTypeIcon, { type: FileTypeIcon_1.IconType.image, size: this.state.imgSize })),
                        React.createElement("div", { className: "ms-font-m" },
                            "List picker tester:",
                            React.createElement(ListPicker_1.ListPicker, { context: this.props.context, label: "Select your list(s)", placeholder: "Select your list(s)", baseTemplate: 100, includeHidden: false, multiSelect: true, 
                                // filter="Title eq 'Test List'"
                                onSelectionChanged: this.onListPickerChange })),
                        React.createElement("div", { className: "ms-font-m" },
                            "List Item picker list data tester:",
                            React.createElement(ListItemPicker_1.ListItemPicker, { listId: '76a8231b-35b6-4703-b1f4-5d03d3dfb1ca', columnInternalName: "Title", keyColumnInternalName: "Id", filter: "Title eq 'SPFx'", itemLimit: 5, context: this.props.context, placeholder: 'Select list items', onSelectedItem: this.listItemPickerDataSelected })),
                        React.createElement("div", null, "Icon Picker"),
                        React.createElement("div", null,
                            React.createElement(iconPicker_1.IconPicker, { renderOption: "panel", onSave: function (value) { console.log(value); }, currentIcon: 'Warning', buttonLabel: "Icon Picker" })),
                        React.createElement("div", { className: "ms-font-m" },
                            "ComboBoxListItemPicker:",
                            React.createElement(ComboBoxListItemPicker_1.ComboBoxListItemPicker, { listId: '0ffa51d7-4ad1-4f04-8cfe-98209905d6da', columnInternalName: 'Title', keyColumnInternalName: 'Id', multiSelect: true, onSelectedItem: function (data) {
                                    console.log("Item(s):", data);
                                }, webUrl: this.props.context.pageContext.web.absoluteUrl, spHttpClient: this.props.context.spHttpClient })),
                        React.createElement("div", { className: "ms-font-m" },
                            "iframe dialog tester:",
                            React.createElement(Button_1.PrimaryButton, { text: "Open iframe Dialog", onClick: function () { _this.setState({ iFrameDialogOpened: true }); } }),
                            React.createElement(IFrameDialog_1.IFrameDialog, { url: iframeUrl, iframeOnLoad: function (iframe) { console.log('iframe loaded'); }, hidden: !this.state.iFrameDialogOpened, onDismiss: function () { _this.setState({ iFrameDialogOpened: false }); }, modalProps: {
                                    isBlocking: true
                                }, dialogContentProps: {
                                    type: Dialog_1.DialogType.close,
                                    showCloseButton: true
                                }, width: '570px', height: '315px' })),
                        React.createElement("div", { className: "ms-font-m" },
                            "iframe Panel tester:",
                            React.createElement(Button_1.PrimaryButton, { text: "Open iframe Panel", onClick: function () { _this.setState({ iFramePanelOpened: true }); } }),
                            React.createElement(IFramePanel_1.IFramePanel, { url: iframeUrl, type: Panel_1.PanelType.medium, 
                                //  height="300px"
                                headerText: "iframe panel title", closeButtonAriaLabel: "Close", isOpen: this.state.iFramePanelOpened, onDismiss: function () { _this.setState({ iFramePanelOpened: false }); }, iframeOnLoad: function (iframe) { console.log('iframe loaded'); } })),
                        React.createElement("div", null,
                            React.createElement(FolderPicker_1.default, { context: this.props.context, rootFolder: {
                                    Name: 'Documents',
                                    ServerRelativeUrl: (this.props.context.pageContext.web.serverRelativeUrl === '/' ? '' : this.props.context.pageContext.web.serverRelativeUrl) + "/Shared Documents"
                                }, onSelect: this._onFolderSelect, label: 'Folder Picker', required: true, canCreateFolders: true }))))),
            React.createElement("div", null,
                React.createElement("h3", null, "Carousel with fixed elements:"),
                React.createElement(carousel_1.Carousel, { buttonsLocation: carousel_1.CarouselButtonsLocation.top, buttonsDisplay: carousel_1.CarouselButtonsDisplay.block, contentContainerStyles: ControlsTest_module_scss_1.default.carouselContent, containerButtonsStyles: ControlsTest_module_scss_1.default.carouselButtonsContainer, isInfinite: true, element: this.carouselElements, onMoveNextClicked: function (index) { console.log("Next button clicked: " + index); }, onMovePrevClicked: function (index) { console.log("Prev button clicked: " + index); } })),
            React.createElement("div", null,
                React.createElement("h3", null, "Carousel with CarouselImage elements:"),
                React.createElement(carousel_1.Carousel, { buttonsLocation: carousel_1.CarouselButtonsLocation.center, buttonsDisplay: carousel_1.CarouselButtonsDisplay.buttonsOnly, contentContainerStyles: ControlsTest_module_scss_1.default.carouselImageContent, 
                    //containerButtonsStyles={styles.carouselButtonsContainer}
                    isInfinite: true, indicatorShape: carousel_1.CarouselIndicatorShape.circle, pauseOnHover: true, element: [
                        {
                            imageSrc: 'https://images.unsplash.com/photo-1588614959060-4d144f28b207?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=3078&q=80',
                            title: 'Colosseum',
                            description: 'This is Colosseum',
                            url: 'https://en.wikipedia.org/wiki/Colosseum',
                            showDetailsOnHover: true,
                            imageFit: Image_1.ImageFit.cover
                        },
                        {
                            imageSrc: 'https://www.telegraph.co.uk/content/dam/science/2018/06/20/stonehenge-2326750_1920_trans%2B%2BZgEkZX3M936N5BQK4Va8RWtT0gK_6EfZT336f62EI5U.jpg',
                            title: 'Stonehenge',
                            description: 'This is Stonehendle',
                            url: 'https://en.wikipedia.org/wiki/Stonehenge',
                            showDetailsOnHover: true,
                            imageFit: Image_1.ImageFit.cover
                        },
                        {
                            imageSrc: 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/af/All_Gizah_Pyramids.jpg/2560px-All_Gizah_Pyramids.jpg',
                            title: 'Pyramids of Giza',
                            description: 'This are Pyramids of Giza (Egypt)',
                            url: 'https://en.wikipedia.org/wiki/Egyptian_pyramids',
                            showDetailsOnHover: true,
                            imageFit: Image_1.ImageFit.cover
                        }
                    ], onMoveNextClicked: function (index) { console.log("Next button clicked: " + index); }, onMovePrevClicked: function (index) { console.log("Prev button clicked: " + index); } })),
            React.createElement("div", null,
                React.createElement("h3", null, "Carousel with triggerPageElement:"),
                React.createElement(carousel_1.Carousel, { buttonsLocation: carousel_1.CarouselButtonsLocation.bottom, buttonsDisplay: carousel_1.CarouselButtonsDisplay.buttonsOnly, contentContainerStyles: ControlsTest_module_scss_1.default.carouselContent, canMoveNext: this.state.canMoveNext, canMovePrev: this.state.canMovePrev, triggerPageEvent: this.triggerNextElement, element: this.state.currentCarouselElement })),
            React.createElement("div", { className: ControlsTest_module_scss_1.default.siteBreadcrumb },
                React.createElement(SiteBreadcrumb_1.SiteBreadcrumb, { context: this.props.context })),
            React.createElement("div", null,
                React.createElement(FilePicker_1.FilePicker, { bingAPIKey: "<BING API KEY>", 
                    //accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
                    buttonLabel: "Upload image", buttonIcon: "FileImage", onSave: this._onFilePickerSave, onChanged: function (filePickerResult) { _this.setState({ filePickerResult: filePickerResult }); }, context: this.props.context, hideRecentTab: true }),
                this.state.filePickerResult &&
                    React.createElement("div", null,
                        "FileName: ",
                        this.state.filePickerResult.fileName)),
            React.createElement("p", null,
                React.createElement("a", { href: "javascript:;", onClick: this.deleteItem }, "Deletes second item")),
            React.createElement("div", null,
                React.createElement(Progress_1.Progress, { title: 'Progress Test', showOverallProgress: true, showIndeterminateOverallProgress: false, hideNotStartedActions: false, actions: this.state.progressActions, currentActionIndex: this.state.currentProgressActionIndex, longRunningText: 'This operation takes longer than expected', longRunningTextDisplayDelay: 7000, height: '350px', inProgressIconName: 'ChromeBackMirrored' }),
                React.createElement(Button_1.PrimaryButton, { text: 'Start Progress', onClick: this._startProgress })),
            React.createElement("div", { className: "ms-font-l" }, "Grid Layout"),
            React.createElement(GridLayout_1.GridLayout, { ariaLabel: "List of content, use right and left arrow keys to navigate, arrow down to access details.", items: sampleGridData, onRenderGridItem: function (item, finalSize, isCompact) { return _this._onRenderGridItem(item, finalSize, isCompact); } }),
            React.createElement(iconPicker_1.IconPicker, { buttonLabel: 'Icon', onChange: function (iconName) { console.log(iconName); }, onSave: function (iconName) { console.log(iconName); } }),
            React.createElement("div", null,
                React.createElement(FolderExplorer_1.FolderExplorer, { context: this.props.context, rootFolder: {
                        Name: 'Documents',
                        ServerRelativeUrl: (this.props.context.pageContext.web.serverRelativeUrl === '/' ? '' : this.props.context.pageContext.web.serverRelativeUrl) + "/Shared Documents"
                    }, defaultFolder: {
                        Name: 'Documents',
                        ServerRelativeUrl: (this.props.context.pageContext.web.serverRelativeUrl === '/' ? '' : this.props.context.pageContext.web.serverRelativeUrl) + "/Shared Documents"
                    }, onSelect: this._onFolderSelect, canCreateFolders: true })),
            React.createElement("div", null,
                React.createElement("h3", null, "Tree View"),
                React.createElement(treeView_1.TreeView, { items: this.treeitems, defaultExpanded: false, selectionMode: treeView_1.TreeViewSelectionMode.Multiple, showCheckboxes: true, treeItemActionsDisplayMode: treeView_1.TreeItemActionsDisplayMode.ContextualMenu, defaultSelectedKeys: ['gc1', 'gc3'], onExpandCollapse: this.onExpandCollapseTree, onSelect: this.onItemSelected })),
            React.createElement("div", null,
                React.createElement(pagination_1.Pagination, { currentPage: 3, onChange: function (page) { return (_this._getPage(page)); }, totalPages: this.props.totalPages || 13 })),
            React.createElement("div", null,
                React.createElement(FieldCollectionData_1.FieldCollectionData, { key: "FieldCollectionData", label: "Fields Collection", itemsPerPage: 3, manageBtnLabel: "Manage", onChanged: function (value) { console.log(value); }, panelHeader: "Manage values", enableSorting: true, fields: [
                        { id: "Field1", title: "String field", type: FieldCollectionData_1.CustomCollectionFieldType.string, required: true },
                        { id: "Field2", title: "Number field", type: FieldCollectionData_1.CustomCollectionFieldType.number },
                        { id: "Field3", title: "URL field", type: FieldCollectionData_1.CustomCollectionFieldType.url },
                        { id: "Field4", title: "Boolean field", type: FieldCollectionData_1.CustomCollectionFieldType.boolean },
                    ], value: this.getRandomCollectionFieldData() }))));
    };
    ControlsTest.prototype.onExpandCollapseTree = function (item, isExpanded) {
        console.log((isExpanded ? "item expanded: " : "item collapsed: ") + item);
    };
    ControlsTest.prototype.onItemSelected = function (items) {
        console.log("items selected: " + items.length);
    };
    ControlsTest.prototype.renderCustomTreeItem = function (item) {
        return (React.createElement("span", null,
            item.iconProps &&
                React.createElement("i", { className: "ms-Icon ms-Icon--" + item.iconProps.iconName, style: { paddingRight: '4px' } }),
            item.label));
    };
    ControlsTest.prototype._getPage = function (page) {
        console.log('Page:', page);
    };
    return ControlsTest;
}(React.Component));
exports.default = ControlsTest;

//# sourceMappingURL=ControlsTest.js.map
