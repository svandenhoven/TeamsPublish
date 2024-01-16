"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
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
exports.__esModule = true;
var identity_1 = require("@azure/identity");
var azureTokenCredentials_1 = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
var microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
require('isomorphic-fetch');
var fs = require("fs");
var axios_1 = require("axios");
// Get the command line arguments
var args = process.argv.slice(2);
var command = args[0];
var appId = args[1];
var tenantId = args[2];
var clientId = args[3];
var userName = args[4];
var password = args[5];
// Check if there are any arguments
if (args.length < 6) {
    console.log("No arguments provided. node common.js <command> <appId> <tenantId> <clientId> <userName> <password>");
}
else {
    // Print the provided arguments
    console.log("Arguments provided");
}
// User to get access to App Catalog
// Requirments:
//      - User must be a Teams Service Administrator
//      - User must have a Teams license
var credential = new identity_1.UsernamePasswordCredential(tenantId, clientId, userName, password);
var authProvider = new azureTokenCredentials_1.TokenCredentialAuthenticationProvider(credential, {
    scopes: ['User.Read', 'AppCatalog.ReadWrite.All']
});
function getToken() {
    return __awaiter(this, void 0, void 0, function () {
        var response;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, credential.getToken(['User.Read', 'AppCatalog.ReadWrite.All'])];
                case 1:
                    response = _a.sent();
                    return [2 /*return*/, response.token];
            }
        });
    });
}
var graphClient = microsoft_graph_client_1.Client.initWithMiddleware({ authProvider: authProvider });
function getApps() {
    return __awaiter(this, void 0, void 0, function () {
        var teamsApps;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, graphClient.api('/appCatalogs/teamsApps')
                        .filter('distributionMethod eq \'organization\'')
                        .get()];
                case 1:
                    teamsApps = _a.sent();
                    console.log(teamsApps);
                    return [2 /*return*/];
            }
        });
    });
}
function PostData(data, url) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                    var config, _a;
                    var _b, _c;
                    return __generator(this, function (_d) {
                        switch (_d.label) {
                            case 0:
                                _b = {
                                    method: 'post',
                                    url: url
                                };
                                _c = {};
                                _a = 'Authorization';
                                return [4 /*yield*/, getToken()];
                            case 1:
                                config = (_b.headers = (_c[_a] = _d.sent(),
                                    _c['Content-Type'] = 'application/zip',
                                    _c),
                                    _b.data = data,
                                    _b);
                                (0, axios_1["default"])(config);
                                return [2 /*return*/];
                        }
                    });
                }); })];
        });
    });
}
function publishApp() {
    return __awaiter(this, void 0, void 0, function () {
        var teamsApp;
        var _this = this;
        return __generator(this, function (_a) {
            teamsApp = fs.readFile('./package/appPackage.local.zip', function (err, data) { return __awaiter(_this, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            if (err)
                                throw err;
                            return [4 /*yield*/, PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?requiresReview=false')];
                        case 1:
                            _a.sent();
                            console.log('App published');
                            return [2 /*return*/];
                    }
                });
            }); });
            return [2 /*return*/];
        });
    });
}
function updateApp(appId) {
    return __awaiter(this, void 0, void 0, function () {
        var teamsApp;
        var _this = this;
        return __generator(this, function (_a) {
            teamsApp = fs.readFile('./package/appPackage.local.zip', function (err, data) { return __awaiter(_this, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            if (err)
                                throw err;
                            return [4 /*yield*/, PostData(data, 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/' + appId + '/appDefinitions')];
                        case 1:
                            _a.sent();
                            console.log('App updated');
                            return [2 /*return*/];
                    }
                });
            }); });
            return [2 /*return*/];
        });
    });
}
if (command === 'update')
    updateApp(appId);
if (command === 'publish')
    publishApp();
if (command === 'list')
    getApps();
