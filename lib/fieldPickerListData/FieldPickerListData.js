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
import * as React from "react";
import { TagPicker } from "office-ui-fabric-react/lib/components/pickers/TagPicker/TagPicker";
import SPservice from "../services/SPservice";
import { Label } from "office-ui-fabric-react/lib/Label";
var FieldPickerListData = (function (_super) {
    __extends(FieldPickerListData, _super);
    function FieldPickerListData(props) {
        var _this = _super.call(this, props) || this;
        // States
        _this.state = {
            noresultsFoundText: typeof _this.props.noresultsFoundText === undefined
                ? "No Items Found"
                : _this.props.noresultsFoundText,
            showError: false,
            errorMessage: "",
            suggestionsHeaderText: typeof _this.props.sugestedHeaderText === undefined
                ? "Select Value"
                : _this.props.sugestedHeaderText
        };
        // Get SPService Factory
        _this._spservice = new SPservice(_this.props.context);
        // handlers
        _this.onFilterChanged = _this.onFilterChanged.bind(_this);
        _this.getTextFromItem = _this.getTextFromItem.bind(_this);
        _this.onItemChanged = _this.onItemChanged.bind(_this);
        // Teste Parameters
        _this._value = _this.props.value !== undefined ? _this.props.value : [];
        return _this;
    }
    // Render Field
    FieldPickerListData.prototype.render = function () {
        var _a = this.props, className = _a.className, disabled = _a.disabled, itemLimit = _a.itemLimit;
        return (React.createElement("div", null,
            React.createElement(TagPicker, { onResolveSuggestions: this.onFilterChanged, 
                //   getTextFromItem={(item: any) => { return item.name; }}
                getTextFromItem: this.getTextFromItem, pickerSuggestionsProps: {
                    suggestionsHeaderText: this.state.suggestionsHeaderText,
                    noResultsFoundText: this.state.noresultsFoundText
                }, defaultSelectedItems: this._value, onChange: this.onItemChanged, className: className, itemLimit: itemLimit, disabled: disabled }),
            React.createElement(Label, { color: "red" },
                " ",
                this.state.errorMessage,
                " ")));
    };
    // Get text from Item
    FieldPickerListData.prototype.getTextFromItem = function (item) {
        return item.name;
    };
    /*
    On Selected Item
  */
    FieldPickerListData.prototype.onItemChanged = function (selectedItems) {
        var item = selectedItems[0];
        console.log("selected items nr: " + selectedItems.length);
        this.props.onSelectedItem(selectedItems);
    };
    // Filter Change
    FieldPickerListData.prototype.onFilterChanged = function (filterText, tagList) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.loadListItems(filterText)
                .then(function (resolvedSugestions) {
                _this.setState({
                    errorMessage: "",
                    showError: false
                });
                resolve(resolvedSugestions);
            })
                .catch(function (reason) {
                console.log("Error get Items " + reason);
                _this.setState({
                    showError: true,
                    errorMessage: reason.message,
                    noresultsFoundText: reason.message
                });
                resolve([]);
            });
        });
    };
    /*
  Function to load List Items
  */
    FieldPickerListData.prototype.loadListItems = function (filterText) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, listId, columnInternalName, webUrl, arrayItems, listItems, error_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this.props, listId = _a.listId, columnInternalName = _a.columnInternalName, webUrl = _a.webUrl;
                        arrayItems = [];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._spservice.getListItems(filterText, listId, columnInternalName, webUrl)];
                    case 2:
                        listItems = _b.sent();
                        // has Items ?
                        if (listItems.length > 0) {
                            listItems.map(function (item, i) {
                                arrayItems.push({ key: item.Id, name: item[columnInternalName] });
                            });
                        }
                        return [2 /*return*/, Promise.resolve(arrayItems)];
                    case 3:
                        error_1 = _b.sent();
                        return [2 /*return*/, Promise.reject(error_1)];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return FieldPickerListData;
}(React.Component));
export { FieldPickerListData };
//# sourceMappingURL=FieldPickerListData.js.map