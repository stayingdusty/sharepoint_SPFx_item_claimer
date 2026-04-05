define("fae2eec7-5401-4ea3-a4a8-9958dd98721f_1.0.0", ["@microsoft/sp-core-library","@microsoft/sp-dialog","@microsoft/sp-listview-extensibility","@microsoft/sp-http","ClaimQueueItemCommandSetStrings"], (__WEBPACK_EXTERNAL_MODULE__878__, __WEBPACK_EXTERNAL_MODULE__529__, __WEBPACK_EXTERNAL_MODULE__249__, __WEBPACK_EXTERNAL_MODULE__272__, __WEBPACK_EXTERNAL_MODULE__320__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 878:
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__878__;

/***/ }),

/***/ 529:
/*!***************************************!*\
  !*** external "@microsoft/sp-dialog" ***!
  \***************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__529__;

/***/ }),

/***/ 272:
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__272__;

/***/ }),

/***/ 249:
/*!*******************************************************!*\
  !*** external "@microsoft/sp-listview-extensibility" ***!
  \*******************************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__249__;

/***/ }),

/***/ 320:
/*!**************************************************!*\
  !*** external "ClaimQueueItemCommandSetStrings" ***!
  \**************************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__320__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!*******************************************************************!*\
  !*** ./lib/extensions/claimQueueItem/ClaimQueueItemCommandSet.js ***!
  \*******************************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-core-library */ 878);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-dialog */ 529);
/* harmony import */ var _microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_listview_extensibility__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-listview-extensibility */ 249);
/* harmony import */ var _microsoft_sp_listview_extensibility__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_listview_extensibility__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-http */ 272);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ClaimQueueItemCommandSetStrings */ 320);
/* harmony import */ var ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__);
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
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





var LOG_SOURCE = 'ClaimQueueItemCommandSet';
var ClaimQueueItemCommandSet = /** @class */ (function (_super) {
    __extends(ClaimQueueItemCommandSet, _super);
    function ClaimQueueItemCommandSet() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ClaimQueueItemCommandSet.prototype.onInit = function () {
        _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__.Log.info(LOG_SOURCE, 'Initialized ClaimQueueItemCommandSet');
        return Promise.resolve();
    };
    ClaimQueueItemCommandSet.prototype.onListViewUpdated = function (event) {
        var _a;
        var claimCommand = this.tryGetCommand('CLAIM_ITEM');
        if (claimCommand) {
            claimCommand.visible = ((_a = event.selectedRows) === null || _a === void 0 ? void 0 : _a.length) === 1;
        }
    };
    ClaimQueueItemCommandSet.prototype.onExecute = function (event) {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = event.itemId;
                        switch (_a) {
                            case 'CLAIM_ITEM': return [3 /*break*/, 1];
                        }
                        return [3 /*break*/, 3];
                    case 1: return [4 /*yield*/, this._claimSelectedItem(event.selectedRows[0])];
                    case 2:
                        _b.sent();
                        return [3 /*break*/, 4];
                    case 3: throw new Error("Unknown command: ".concat(event.itemId));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ClaimQueueItemCommandSet.prototype._claimSelectedItem = function (selectedRow) {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var itemId, claimFieldInternalName, currentUserId, itemResponse, item, assignedUserLabel, etag, updateResponse, error_1, message;
            var _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        itemId = this._getSelectedItemId(selectedRow);
                        claimFieldInternalName = this.properties.claimFieldInternalName || 'Assigned_To';
                        currentUserId = Number(((_a = this.context.pageContext.legacyPageContext) === null || _a === void 0 ? void 0 : _a.userId) || 0);
                        if (!currentUserId) {
                            throw new Error('Could not resolve the current SharePoint user ID.');
                        }
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 11, , 13]);
                        return [4 /*yield*/, this.context.spHttpClient.get(this._getReadItemUrl(itemId, claimFieldInternalName), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_3__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata'
                                }
                            })];
                    case 2:
                        itemResponse = _c.sent();
                        if (!itemResponse.ok) {
                            throw new Error("Could not load the current item. Status ".concat(itemResponse.status, "."));
                        }
                        return [4 /*yield*/, itemResponse.json()];
                    case 3:
                        item = (_c.sent());
                        assignedUserLabel = this._getAssignedUserLabel(item, claimFieldInternalName);
                        if (!this._hasAssignee(item, claimFieldInternalName)) return [3 /*break*/, 5];
                        return [4 /*yield*/, _microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1__.Dialog.alert(assignedUserLabel
                                ? "Already taken by ".concat(assignedUserLabel, ".")
                                : ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__.AlreadyTakenMessage)];
                    case 4:
                        _c.sent();
                        return [2 /*return*/];
                    case 5:
                        etag = itemResponse.headers.get('ETag') || item['@odata.etag'] || '*';
                        return [4 /*yield*/, this.context.spHttpClient.post(this._getUpdateItemUrl(itemId), _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_3__.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata',
                                    'IF-MATCH': etag,
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: JSON.stringify((_b = {},
                                    _b["".concat(claimFieldInternalName, "Id")] = currentUserId,
                                    _b))
                            })];
                    case 6:
                        updateResponse = _c.sent();
                        if (!updateResponse.ok) return [3 /*break*/, 8];
                        return [4 /*yield*/, _microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1__.Dialog.alert(ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__.SuccessMessage)];
                    case 7:
                        _c.sent();
                        window.location.reload();
                        return [2 /*return*/];
                    case 8:
                        if (!(updateResponse.status === 412)) return [3 /*break*/, 10];
                        return [4 /*yield*/, _microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1__.Dialog.alert(ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__.AlreadyTakenMessage)];
                    case 9:
                        _c.sent();
                        return [2 /*return*/];
                    case 10: throw new Error("Claim update failed. Status ".concat(updateResponse.status, "."));
                    case 11:
                        error_1 = _c.sent();
                        message = error_1 instanceof Error ? error_1.message : ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__.UnexpectedErrorMessage;
                        return [4 /*yield*/, _microsoft_sp_dialog__WEBPACK_IMPORTED_MODULE_1__.Dialog.alert("".concat(ClaimQueueItemCommandSetStrings__WEBPACK_IMPORTED_MODULE_4__.UnexpectedErrorMessage, "\n\n").concat(message))];
                    case 12:
                        _c.sent();
                        return [3 /*break*/, 13];
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    ClaimQueueItemCommandSet.prototype._getSelectedItemId = function (selectedRow) {
        var rawValue = selectedRow.getValueByName('ID') || selectedRow.getValueByName('Id');
        var itemId = Number(rawValue);
        if (!itemId) {
            throw new Error('A SharePoint list item must be selected before it can be claimed.');
        }
        return itemId;
    };
    ClaimQueueItemCommandSet.prototype._getReadItemUrl = function (itemId, claimFieldInternalName) {
        var listId = this._getListId();
        var selectClause = [
            'Id',
            'Title',
            "".concat(claimFieldInternalName, "Id"),
            "".concat(claimFieldInternalName, "/Id"),
            "".concat(claimFieldInternalName, "/Title"),
            "".concat(claimFieldInternalName, "/EMail")
        ].join(',');
        return "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists(guid'").concat(listId, "')/items(").concat(itemId, ")?$select=").concat(selectClause, "&$expand=").concat(claimFieldInternalName);
    };
    ClaimQueueItemCommandSet.prototype._getUpdateItemUrl = function (itemId) {
        var listId = this._getListId();
        return "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/lists(guid'").concat(listId, "')/items(").concat(itemId, ")");
    };
    ClaimQueueItemCommandSet.prototype._getListId = function () {
        var _a, _b;
        var listId = (_b = (_a = this.context.pageContext.list) === null || _a === void 0 ? void 0 : _a.id) === null || _b === void 0 ? void 0 : _b.toString();
        if (!listId) {
            throw new Error('This command can only run from a SharePoint list view.');
        }
        return listId;
    };
    ClaimQueueItemCommandSet.prototype._hasAssignee = function (item, claimFieldInternalName) {
        var assigneeId = item["".concat(claimFieldInternalName, "Id")];
        if (typeof assigneeId === 'number') {
            return assigneeId > 0;
        }
        if (typeof assigneeId === 'string') {
            return assigneeId.trim().length > 0 && assigneeId !== '0';
        }
        return Boolean(assigneeId);
    };
    ClaimQueueItemCommandSet.prototype._getAssignedUserLabel = function (item, claimFieldInternalName) {
        var assignee = item[claimFieldInternalName];
        return (assignee === null || assignee === void 0 ? void 0 : assignee.Title) || (assignee === null || assignee === void 0 ? void 0 : assignee.EMail);
    };
    return ClaimQueueItemCommandSet;
}(_microsoft_sp_listview_extensibility__WEBPACK_IMPORTED_MODULE_2__.BaseListViewCommandSet));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (ClaimQueueItemCommandSet);

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});;
//# sourceMappingURL=claim-queue-item-command-set.js.map