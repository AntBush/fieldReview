(window["webpackJsonp"] = window["webpackJsonp"] || []).push([["main"],{

/***/ "./$$_lazy_route_resource lazy recursive":
/*!******************************************************!*\
  !*** ./$$_lazy_route_resource lazy namespace object ***!
  \******************************************************/
/*! no static exports found */
/***/ (function(module, exports) {

function webpackEmptyAsyncContext(req) {
	// Here Promise.resolve().then() is used instead of new Promise() to prevent
	// uncaught exception popping up in devtools
	return Promise.resolve().then(function() {
		var e = new Error("Cannot find module '" + req + "'");
		e.code = 'MODULE_NOT_FOUND';
		throw e;
	});
}
webpackEmptyAsyncContext.keys = function() { return []; };
webpackEmptyAsyncContext.resolve = webpackEmptyAsyncContext;
module.exports = webpackEmptyAsyncContext;
webpackEmptyAsyncContext.id = "./$$_lazy_route_resource lazy recursive";

/***/ }),

/***/ "./src/app/app-routing.module.ts":
/*!***************************************!*\
  !*** ./src/app/app-routing.module.ts ***!
  \***************************************/
/*! exports provided: AppRoutingModule */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppRoutingModule", function() { return AppRoutingModule; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "./node_modules/@angular/core/__ivy_ngcc__/fesm2015/core.js");
/* harmony import */ var _angular_router__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @angular/router */ "./node_modules/@angular/router/__ivy_ngcc__/fesm2015/router.js");




const routes = [];
class AppRoutingModule {
}
AppRoutingModule.ɵmod = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineNgModule"]({ type: AppRoutingModule });
AppRoutingModule.ɵinj = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineInjector"]({ factory: function AppRoutingModule_Factory(t) { return new (t || AppRoutingModule)(); }, imports: [[_angular_router__WEBPACK_IMPORTED_MODULE_1__["RouterModule"].forRoot(routes)], _angular_router__WEBPACK_IMPORTED_MODULE_1__["RouterModule"]] });
(function () { (typeof ngJitMode === "undefined" || ngJitMode) && _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵsetNgModuleScope"](AppRoutingModule, { imports: [_angular_router__WEBPACK_IMPORTED_MODULE_1__["RouterModule"]], exports: [_angular_router__WEBPACK_IMPORTED_MODULE_1__["RouterModule"]] }); })();
/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](AppRoutingModule, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["NgModule"],
        args: [{
                imports: [_angular_router__WEBPACK_IMPORTED_MODULE_1__["RouterModule"].forRoot(routes)],
                exports: [_angular_router__WEBPACK_IMPORTED_MODULE_1__["RouterModule"]]
            }]
    }], null, null); })();


/***/ }),

/***/ "./src/app/app.component.ts":
/*!**********************************!*\
  !*** ./src/app/app.component.ts ***!
  \**********************************/
/*! exports provided: AppComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppComponent", function() { return AppComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "./node_modules/@angular/core/__ivy_ngcc__/fesm2015/core.js");
/* harmony import */ var _docx_form_docx_form_component__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./docx-form/docx-form.component */ "./src/app/docx-form/docx-form.component.ts");
/* harmony import */ var _angular_router__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @angular/router */ "./node_modules/@angular/router/__ivy_ngcc__/fesm2015/router.js");




class AppComponent {
    constructor() {
        this.title = 'docxGenerateUI';
    }
}
AppComponent.ɵfac = function AppComponent_Factory(t) { return new (t || AppComponent)(); };
AppComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: AppComponent, selectors: [["app-root"]], decls: 3, vars: 0, template: function AppComponent_Template(rf, ctx) { if (rf & 1) {
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "div");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](1, "app-docx-form");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](2, "router-outlet");
    } }, directives: [_docx_form_docx_form_component__WEBPACK_IMPORTED_MODULE_1__["DocxFormComponent"], _angular_router__WEBPACK_IMPORTED_MODULE_2__["RouterOutlet"]], styles: ["\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6IiIsImZpbGUiOiJzcmMvYXBwL2FwcC5jb21wb25lbnQuY3NzIn0= */"] });
/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](AppComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-root',
                templateUrl: './app.component.html',
                styleUrls: ['./app.component.css']
            }]
    }], null, null); })();


/***/ }),

/***/ "./src/app/app.module.ts":
/*!*******************************!*\
  !*** ./src/app/app.module.ts ***!
  \*******************************/
/*! exports provided: AppModule */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppModule", function() { return AppModule; });
/* harmony import */ var _angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/platform-browser */ "./node_modules/@angular/platform-browser/__ivy_ngcc__/fesm2015/platform-browser.js");
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @angular/core */ "./node_modules/@angular/core/__ivy_ngcc__/fesm2015/core.js");
/* harmony import */ var _app_routing_module__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./app-routing.module */ "./src/app/app-routing.module.ts");
/* harmony import */ var _app_component__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./app.component */ "./src/app/app.component.ts");
/* harmony import */ var _docx_form_docx_form_component__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./docx-form/docx-form.component */ "./src/app/docx-form/docx-form.component.ts");
/* harmony import */ var _angular_common_http__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @angular/common/http */ "./node_modules/@angular/common/__ivy_ngcc__/fesm2015/http.js");
/* harmony import */ var _angular_forms__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @angular/forms */ "./node_modules/@angular/forms/__ivy_ngcc__/fesm2015/forms.js");








class AppModule {
}
AppModule.ɵmod = _angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵɵdefineNgModule"]({ type: AppModule, bootstrap: [_app_component__WEBPACK_IMPORTED_MODULE_3__["AppComponent"]] });
AppModule.ɵinj = _angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵɵdefineInjector"]({ factory: function AppModule_Factory(t) { return new (t || AppModule)(); }, providers: [], imports: [[
            _angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__["BrowserModule"],
            _app_routing_module__WEBPACK_IMPORTED_MODULE_2__["AppRoutingModule"],
            _angular_forms__WEBPACK_IMPORTED_MODULE_6__["FormsModule"],
            _angular_forms__WEBPACK_IMPORTED_MODULE_6__["ReactiveFormsModule"],
            _angular_common_http__WEBPACK_IMPORTED_MODULE_5__["HttpClientModule"]
        ]] });
(function () { (typeof ngJitMode === "undefined" || ngJitMode) && _angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵɵsetNgModuleScope"](AppModule, { declarations: [_app_component__WEBPACK_IMPORTED_MODULE_3__["AppComponent"],
        _docx_form_docx_form_component__WEBPACK_IMPORTED_MODULE_4__["DocxFormComponent"]], imports: [_angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__["BrowserModule"],
        _app_routing_module__WEBPACK_IMPORTED_MODULE_2__["AppRoutingModule"],
        _angular_forms__WEBPACK_IMPORTED_MODULE_6__["FormsModule"],
        _angular_forms__WEBPACK_IMPORTED_MODULE_6__["ReactiveFormsModule"],
        _angular_common_http__WEBPACK_IMPORTED_MODULE_5__["HttpClientModule"]] }); })();
/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_1__["ɵsetClassMetadata"](AppModule, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_1__["NgModule"],
        args: [{
                declarations: [
                    _app_component__WEBPACK_IMPORTED_MODULE_3__["AppComponent"],
                    _docx_form_docx_form_component__WEBPACK_IMPORTED_MODULE_4__["DocxFormComponent"]
                ],
                imports: [
                    _angular_platform_browser__WEBPACK_IMPORTED_MODULE_0__["BrowserModule"],
                    _app_routing_module__WEBPACK_IMPORTED_MODULE_2__["AppRoutingModule"],
                    _angular_forms__WEBPACK_IMPORTED_MODULE_6__["FormsModule"],
                    _angular_forms__WEBPACK_IMPORTED_MODULE_6__["ReactiveFormsModule"],
                    _angular_common_http__WEBPACK_IMPORTED_MODULE_5__["HttpClientModule"]
                ],
                providers: [],
                bootstrap: [_app_component__WEBPACK_IMPORTED_MODULE_3__["AppComponent"]]
            }]
    }], null, null); })();


/***/ }),

/***/ "./src/app/docx-form/docx-form.component.ts":
/*!**************************************************!*\
  !*** ./src/app/docx-form/docx-form.component.ts ***!
  \**************************************************/
/*! exports provided: DocxFormComponent */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DocxFormComponent", function() { return DocxFormComponent; });
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "./node_modules/@angular/core/__ivy_ngcc__/fesm2015/core.js");
/* harmony import */ var _angular_forms__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @angular/forms */ "./node_modules/@angular/forms/__ivy_ngcc__/fesm2015/forms.js");
/* harmony import */ var _angular_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @angular/common */ "./node_modules/@angular/common/__ivy_ngcc__/fesm2015/common.js");




function DocxFormComponent_div_47_Template(rf, ctx) { if (rf & 1) {
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "div");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](1, "label");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](2, " Date Visited: ");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](3, "input", 28);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} if (rf & 2) {
    const i_r3 = ctx.index;
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](3);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("formControlName", i_r3);
} }
function DocxFormComponent_div_51_Template(rf, ctx) { if (rf & 1) {
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "div");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](1, "label");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](2, " Inspection Notes: ");
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](3, "input", 29);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
} if (rf & 2) {
    const i_r5 = ctx.index;
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](3);
    _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("formControlName", i_r5);
} }
class DocxFormComponent {
    constructor(fb) {
        this.fb = fb;
        this.docxForm = this.fb.group({
            fieldReviewReports: [''],
            commonElementNo: [''],
            fileNo: [''],
            date: [''],
            projectAddress: [''],
            location: [''],
            referenceDwgs: [''],
            projectName: [''],
            builder: [''],
            weatherCondition: [''],
            inspectionCategory: [''],
            datesVisited: this.fb.array([
                this.fb.control('')
            ]),
            inspectionNotes: this.fb.array([
                this.fb.control('')
            ])
        });
    }
    ngOnInit() {
    }
    logIt() {
        console.log(this.datesVisited.value);
        console.log(this.docxForm.value);
    }
    get datesVisited() {
        return this.docxForm.get('datesVisited');
    }
    addDateVisited() {
        console.log("Why no work");
        this.datesVisited.push(this.fb.control(''));
    }
    get inspectionNotes() {
        return this.docxForm.get('inspectionNotes');
    }
    addInspectionNotes() {
        this.inspectionNotes.push(this.fb.control(''));
    }
}
DocxFormComponent.ɵfac = function DocxFormComponent_Factory(t) { return new (t || DocxFormComponent)(_angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdirectiveInject"](_angular_forms__WEBPACK_IMPORTED_MODULE_1__["FormBuilder"])); };
DocxFormComponent.ɵcmp = _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵdefineComponent"]({ type: DocxFormComponent, selectors: [["app-docx-form"]], decls: 57, vars: 3, consts: [[1, "docxFormLayout", 3, "formGroup"], [1, "fieldReviewReports"], ["type", "text", "id", "fieldReviewReports", "formControlName", "fieldReviewReports"], [1, "commonElementNo"], ["type", "text", "id", "commonElementNo", "formControlName", "commonElementNo"], [1, "fileNo"], ["type", "text", "id", "fileNumber", "formControlName", "fileNo"], [1, "date"], ["type", "date", "id", "dateGeneration", "formControlName", "date"], [1, "projectAddress"], ["type", "text", "id", "projectAddress", "formControlName", "projectAddress"], [1, "location"], ["type", "text", "id", "location", "formControlName", "location"], [1, "referenceDwgs"], ["type", "text", "id", "referenceDwgs", "formControlName", "referenceDwgs"], [1, "projectName"], ["type", "text", "id", "projectName", "formControlName", "projectName"], [1, "builder"], ["type", "text", "id", "builder", "formControlName", "builder"], [1, "weatherCondition"], ["type", "text", "id", "weatherCondition", "formControlName", "weatherCondition"], [1, "inspectionCategory"], ["type", "text", "id", "inspectionCategory", "formControlName", "inspectionCategory"], ["formArrayName", "datesVisited", 1, "datesVisited"], [4, "ngFor", "ngForOf"], [3, "click"], ["formArrayName", "inspectionNotes", 1, "inspectionNotes"], [1, "empty5", 3, "click"], ["type", "date", 3, "formControlName"], ["type", "text", 3, "formControlName"]], template: function DocxFormComponent_Template(rf, ctx) { if (rf & 1) {
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](0, "div");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](1, "form", 0);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](2, "div", 1);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](3, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](4, " File Review Reports: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](5, "input", 2);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](6, "div", 3);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](7, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](8, " Common Element Number: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](9, "input", 4);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](10, "div", 5);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](11, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](12, " File No.: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](13, "input", 6);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](14, "div", 7);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](15, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](16, " Date of Document: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](17, "input", 8);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](18, "div", 9);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](19, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](20, " Project Address: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](21, "input", 10);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](22, "div", 11);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](23, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](24, " Location: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](25, "input", 12);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](26, "div", 13);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](27, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](28, " Reference Dwgs: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](29, "input", 14);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](30, "div", 15);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](31, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](32, " Project Name: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](33, "input", 16);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](34, "div", 17);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](35, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](36, " Builder: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](37, "input", 18);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](38, "div", 19);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](39, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](40, " Weather Condition: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](41, "input", 20);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](42, "div", 21);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](43, "label");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](44, " Inspection Category: ");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](45, "input", 22);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](46, "div", 23);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](47, DocxFormComponent_div_47_Template, 4, 1, "div", 24);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](48, "button", 25);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function DocxFormComponent_Template_button_click_48_listener() { return ctx.addDateVisited(); });
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](49, "Add Date Visited");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](50, "div", 26);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtemplate"](51, DocxFormComponent_div_51_Template, 4, 1, "div", 24);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](52, "button", 25);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function DocxFormComponent_Template_button_click_52_listener() { return ctx.addInspectionNotes(); });
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](53, "Add Inspection Note");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementStart"](54, "button", 27);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵlistener"]("click", function DocxFormComponent_Template_button_click_54_listener() { return ctx.logIt(); });
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵtext"](55, "SubmitIGuess");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelement"](56, "p");
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵelementEnd"]();
    } if (rf & 2) {
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](1);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("formGroup", ctx.docxForm);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](46);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngForOf", ctx.datesVisited.controls);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵadvance"](4);
        _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵɵproperty"]("ngForOf", ctx.inspectionNotes.controls);
    } }, directives: [_angular_forms__WEBPACK_IMPORTED_MODULE_1__["ɵangular_packages_forms_forms_y"], _angular_forms__WEBPACK_IMPORTED_MODULE_1__["NgControlStatusGroup"], _angular_forms__WEBPACK_IMPORTED_MODULE_1__["FormGroupDirective"], _angular_forms__WEBPACK_IMPORTED_MODULE_1__["DefaultValueAccessor"], _angular_forms__WEBPACK_IMPORTED_MODULE_1__["NgControlStatus"], _angular_forms__WEBPACK_IMPORTED_MODULE_1__["FormControlName"], _angular_forms__WEBPACK_IMPORTED_MODULE_1__["FormArrayName"], _angular_common__WEBPACK_IMPORTED_MODULE_2__["NgForOf"]], styles: ["[_nghost-%COMP%]   .docxFormLayout[_ngcontent-%COMP%]{\n    display: grid;\n    grid-template-columns: auto auto;\n    grid-template-rows: auto;\n    grid-template-areas: \n    \"fieldReviewReports empty1\"\n    \"commonElementNo empty2\"\n    \"fileNo projectName\"\n    \"date builder\"\n    \"projectAddress weatherCondition\"\n    \"location inspectionCategory\"\n    \"referenceDwgs empty3\"\n    \"datesVisited empty4\"\n    \"inspectionNotes empty5\";\n}\n\n[_nghost-%COMP%]   .datesVisited[_ngcontent-%COMP%] {\n    grid-area: datesVisited;\n}\n\n[_nghost-%COMP%]   .inspectionNotes[_ngcontent-%COMP%] {\n    grid-area: inspectionNotes;\n}\n\n[_nghost-%COMP%]   .referenceDwgs[_ngcontent-%COMP%] {\n    grid-area: referenceDwgs;\n}\n\n[_nghost-%COMP%]   .inspectionCategory[_ngcontent-%COMP%] {\n    grid-area: inspectionCategory;\n}\n\n[_nghost-%COMP%]   .location[_ngcontent-%COMP%] {\n    grid-area: location;\n}\n\n[_nghost-%COMP%]   .projectAddress[_ngcontent-%COMP%] {\n    grid-area: projectAddress;\n}\n\n[_nghost-%COMP%]   .weatherCondition[_ngcontent-%COMP%] {\n    grid-area: weatherCondition;\n}\n\n[_nghost-%COMP%]   .fieldReviewReports[_ngcontent-%COMP%] {\n    grid-area: fieldReviewReports;\n}\n\n[_nghost-%COMP%]   .commonElementNo[_ngcontent-%COMP%] {\n    grid-area: commonElementNo;\n}\n\n[_nghost-%COMP%]   .fileNo[_ngcontent-%COMP%]{\n    grid-area: fileNo;\n}\n\n[_nghost-%COMP%]   .projectName[_ngcontent-%COMP%] {\n    grid-area: projectName;\n}\n\n[_nghost-%COMP%]   .date[_ngcontent-%COMP%] {\n    grid-area: date;\n}\n\n[_nghost-%COMP%]   .bulder[_ngcontent-%COMP%] {\n    grid-area: builder;\n}\n\n[_nghost-%COMP%]   .empty5[_ngcontent-%COMP%] {\n    grid-area: empty5;\n}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbInNyYy9hcHAvZG9jeC1mb3JtL2RvY3gtZm9ybS5jb21wb25lbnQuY3NzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiJBQUFBO0lBQ0ksYUFBYTtJQUNiLGdDQUFnQztJQUNoQyx3QkFBd0I7SUFDeEI7Ozs7Ozs7Ozs0QkFTd0I7QUFDNUI7O0FBRUE7SUFDSSx1QkFBdUI7QUFDM0I7O0FBRUE7SUFDSSwwQkFBMEI7QUFDOUI7O0FBRUE7SUFDSSx3QkFBd0I7QUFDNUI7O0FBQ0E7SUFDSSw2QkFBNkI7QUFDakM7O0FBRUE7SUFDSSxtQkFBbUI7QUFDdkI7O0FBQ0E7SUFDSSx5QkFBeUI7QUFDN0I7O0FBRUE7SUFDSSwyQkFBMkI7QUFDL0I7O0FBRUE7SUFDSSw2QkFBNkI7QUFDakM7O0FBRUE7SUFDSSwwQkFBMEI7QUFDOUI7O0FBQ0E7SUFDSSxpQkFBaUI7QUFDckI7O0FBRUE7SUFDSSxzQkFBc0I7QUFDMUI7O0FBRUE7SUFDSSxlQUFlO0FBQ25COztBQUVBO0lBQ0ksa0JBQWtCO0FBQ3RCOztBQUVBO0lBQ0ksaUJBQWlCO0FBQ3JCIiwiZmlsZSI6InNyYy9hcHAvZG9jeC1mb3JtL2RvY3gtZm9ybS5jb21wb25lbnQuY3NzIiwic291cmNlc0NvbnRlbnQiOlsiOmhvc3QgLmRvY3hGb3JtTGF5b3V0e1xuICAgIGRpc3BsYXk6IGdyaWQ7XG4gICAgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiBhdXRvIGF1dG87XG4gICAgZ3JpZC10ZW1wbGF0ZS1yb3dzOiBhdXRvO1xuICAgIGdyaWQtdGVtcGxhdGUtYXJlYXM6IFxuICAgIFwiZmllbGRSZXZpZXdSZXBvcnRzIGVtcHR5MVwiXG4gICAgXCJjb21tb25FbGVtZW50Tm8gZW1wdHkyXCJcbiAgICBcImZpbGVObyBwcm9qZWN0TmFtZVwiXG4gICAgXCJkYXRlIGJ1aWxkZXJcIlxuICAgIFwicHJvamVjdEFkZHJlc3Mgd2VhdGhlckNvbmRpdGlvblwiXG4gICAgXCJsb2NhdGlvbiBpbnNwZWN0aW9uQ2F0ZWdvcnlcIlxuICAgIFwicmVmZXJlbmNlRHdncyBlbXB0eTNcIlxuICAgIFwiZGF0ZXNWaXNpdGVkIGVtcHR5NFwiXG4gICAgXCJpbnNwZWN0aW9uTm90ZXMgZW1wdHk1XCI7XG59XG5cbjpob3N0IC5kYXRlc1Zpc2l0ZWQge1xuICAgIGdyaWQtYXJlYTogZGF0ZXNWaXNpdGVkO1xufVxuXG46aG9zdCAuaW5zcGVjdGlvbk5vdGVzIHtcbiAgICBncmlkLWFyZWE6IGluc3BlY3Rpb25Ob3Rlcztcbn1cblxuOmhvc3QgLnJlZmVyZW5jZUR3Z3Mge1xuICAgIGdyaWQtYXJlYTogcmVmZXJlbmNlRHdncztcbn1cbjpob3N0IC5pbnNwZWN0aW9uQ2F0ZWdvcnkge1xuICAgIGdyaWQtYXJlYTogaW5zcGVjdGlvbkNhdGVnb3J5O1xufVxuXG46aG9zdCAubG9jYXRpb24ge1xuICAgIGdyaWQtYXJlYTogbG9jYXRpb247XG59XG46aG9zdCAucHJvamVjdEFkZHJlc3Mge1xuICAgIGdyaWQtYXJlYTogcHJvamVjdEFkZHJlc3M7XG59XG5cbjpob3N0IC53ZWF0aGVyQ29uZGl0aW9uIHtcbiAgICBncmlkLWFyZWE6IHdlYXRoZXJDb25kaXRpb247XG59XG5cbjpob3N0IC5maWVsZFJldmlld1JlcG9ydHMge1xuICAgIGdyaWQtYXJlYTogZmllbGRSZXZpZXdSZXBvcnRzO1xufVxuXG46aG9zdCAuY29tbW9uRWxlbWVudE5vIHtcbiAgICBncmlkLWFyZWE6IGNvbW1vbkVsZW1lbnRObztcbn1cbjpob3N0IC5maWxlTm97XG4gICAgZ3JpZC1hcmVhOiBmaWxlTm87XG59XG5cbjpob3N0IC5wcm9qZWN0TmFtZSB7XG4gICAgZ3JpZC1hcmVhOiBwcm9qZWN0TmFtZTtcbn1cblxuOmhvc3QgLmRhdGUge1xuICAgIGdyaWQtYXJlYTogZGF0ZTtcbn1cblxuOmhvc3QgLmJ1bGRlciB7XG4gICAgZ3JpZC1hcmVhOiBidWlsZGVyO1xufVxuXG46aG9zdCAuZW1wdHk1IHtcbiAgICBncmlkLWFyZWE6IGVtcHR5NTtcbn0iXX0= */"] });
/*@__PURE__*/ (function () { _angular_core__WEBPACK_IMPORTED_MODULE_0__["ɵsetClassMetadata"](DocxFormComponent, [{
        type: _angular_core__WEBPACK_IMPORTED_MODULE_0__["Component"],
        args: [{
                selector: 'app-docx-form',
                templateUrl: './docx-form.component.html',
                styleUrls: ['./docx-form.component.css']
            }]
    }], function () { return [{ type: _angular_forms__WEBPACK_IMPORTED_MODULE_1__["FormBuilder"] }]; }, null); })();


/***/ }),

/***/ "./src/environments/environment.ts":
/*!*****************************************!*\
  !*** ./src/environments/environment.ts ***!
  \*****************************************/
/*! exports provided: environment */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "environment", function() { return environment; });
// This file can be replaced during build by using the `fileReplacements` array.
// `ng build --prod` replaces `environment.ts` with `environment.prod.ts`.
// The list of file replacements can be found in `angular.json`.
const environment = {
    production: false
};
/*
 * For easier debugging in development mode, you can import the following file
 * to ignore zone related error stack frames such as `zone.run`, `zoneDelegate.invokeTask`.
 *
 * This import should be commented out in production mode because it will have a negative impact
 * on performance if an error is thrown.
 */
// import 'zone.js/dist/zone-error';  // Included with Angular CLI.


/***/ }),

/***/ "./src/main.ts":
/*!*********************!*\
  !*** ./src/main.ts ***!
  \*********************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _angular_core__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @angular/core */ "./node_modules/@angular/core/__ivy_ngcc__/fesm2015/core.js");
/* harmony import */ var _environments_environment__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./environments/environment */ "./src/environments/environment.ts");
/* harmony import */ var _app_app_module__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./app/app.module */ "./src/app/app.module.ts");
/* harmony import */ var _angular_platform_browser__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @angular/platform-browser */ "./node_modules/@angular/platform-browser/__ivy_ngcc__/fesm2015/platform-browser.js");




if (_environments_environment__WEBPACK_IMPORTED_MODULE_1__["environment"].production) {
    Object(_angular_core__WEBPACK_IMPORTED_MODULE_0__["enableProdMode"])();
}
_angular_platform_browser__WEBPACK_IMPORTED_MODULE_3__["platformBrowser"]().bootstrapModule(_app_app_module__WEBPACK_IMPORTED_MODULE_2__["AppModule"])
    .catch(err => console.error(err));


/***/ }),

/***/ 0:
/*!***************************!*\
  !*** multi ./src/main.ts ***!
  \***************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__(/*! /mnt/c/Users/Tony/Documents/Code/fieldReview/src/main/ts/docxGenerateUI/src/main.ts */"./src/main.ts");


/***/ })

},[[0,"runtime","vendor"]]]);
//# sourceMappingURL=main.js.map