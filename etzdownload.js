/*
v0.1a
This Casper.js module handles the automation of report downloads from http://www.timesheetz.net/, the ETZ timesheets site

https://github.com/phhu/etzdownload
*/

var require = patchRequire(require);
var casper; 
var icasper;
var xpath = require('casper').selectXPath;
var fs = require('fs');
var utils = require("utils");
var moment = require('moment');
var Q = require("Q");

var loginCred;
var outputDir;
var captureDir;
var captureEnabled = false;
var URLBASE = "https://www.timesheetz.net";
var currentAgencyName;

exports.setLogin = function(login){
    loginCred = login;
};
exports.setOutputDir = function(dir){
    outputDir = dir;
};
exports.setCasper = function(c){
    icasper = c;
    casper = icasper.create({
        //onError: function(err){console.log("ERROR: " + err);}
        pageSettings: {    webSecurityEnabled: false   }
        ,waitTimeout: 5 * 60 * 1000
        ,stepTimeout: 5 * 60 * 1000
        ,timeout: 5 * 60 * 1000
        ,resourceTimeout: 5 * 60 * 1000 //240s
        //,exitOnError: false
        //, onError: function(msg, backtrace) {
        //    this.capture('error.png');
        //    this.echo("Error: " + msg);
        //    //throw new ErrorFunc("fatal","error","filename",backtrace,msg);
        //}
        //,verbose: true
        //,logLevel: "debug"    
    });
	return casper;
}
exports.getCasper = function(){
	return casper;
};	
exports.setCapture = function(c,dir){
    captureEnabled = !!c;
    captureDir = dir;
}



/* 
This function can be used to include a password file, which would contain something like this:
exports.cred = {
    "email":"rbear@mybpos.net"
    ,"password":"xxxxxxxxx"
};
*/
exports.checkForPwFile = function(filename,path){
    path = path || fs.workingDirectory ;
    filename = filename || "cred.js";
    var modulePath = path + "/" + filename;
    //console.log ("test " + modulePath);
    var cred;
    try {
        cred = require(modulePath).cred;
    } catch(e){
        console.log ("Password file not found (looked at " + modulePath + ")");
    }
    //console.log ("test" + cred);
    if (cred){
        console.log("Password file found at " + modulePath);
        if (cred.email && cred.password){
            console.log("Setting default login to " + cred.email + "\n");
            loginCred = cred;
        } else {
            console.log("Email / password not found in file, please check");
        }
    }
};    

//create an initial promise
var reports = Q.fcall(function () {
    return true;
});

exports.getReport = function(spec){
    reports = reports.then(function(){return getEtzReport(spec);});
};
exports.exit = function(){
    reports = reports.then(function(){ 
        casper.echo('All reports done'); 
        casper.exit(); 
    });
};

    
var getEtzReport = function (spec,doExit){
    
    var deferred = Q.defer();
    
    
    /*
    casper.echo(fs.absolute("."));
    casper.echo(fs.absolute(".") + '/' + spec.filename);
    casper.echo(fs.isWritable(fs.absolute(".") + '/' + spec.filename));
    if (spec.filename && !(fs.isWritable(fs.absolute(".") + '/' + spec.filename))){
        casper.echo ("File " + spec.filename + " not writable! exiting");
        casper.exit();
    }
    */
    spec.login = spec.login || loginCred || {'email':casper.cli.options['email'],'password':casper.cli.options['password'] };
    spec.outputDir = spec.outputDir || outputDir || casper.cli.options['outputdir'];
    login(spec);
    changeAgency(spec);
    runReport(spec);
    
    casper.run(function() {
        //utils.dump(casper.cli.args);
        //this.echo(moment().format('dddd'));
        this.echo ('Download done\n');
        if (doExit===true){this.exit();}
        //this.exit();
        deferred.resolve(true);
    });
    return deferred.promise;
};
exports.getEtzReport = getEtzReport;


//custom per report 
var reportSpecs = {
    "Multi-Batch Pay-Bill Report" : {
        linkXpath: '//td[span/text()="' + "Multi-Batch Pay-Bill Report" + '"]/following-sibling::td/a'     
        ,func: function(spec){
            this.waitForSelector(ETZ_VIEW_REPORT_SEL,function(){
                
                fillEtzSelect(this,'Pay Run Type',spec.report.settings['Pay Run Type']);

                this.waitWhileVisible(ETZ_WAIT_SEL,function() {
                    
                    fillEtzMultiSelect(this,'Payroll Run',spec.report.settings['Payroll Run']);
                    submitEtzReport(this,spec,"multibatch pay report");

                });
            }); 
        }
    }
    ,"Sales Invoice Day Book" : {
        linkXpath: '//a[text()="Sales Invoice Day Book"]'     
        ,func: function(spec){
            this.echo("starting sales invoice day book form entry");
            this.waitForSelector(ETZ_VIEW_REPORT_SEL,function(){

                fillEtzForm(this,'Start Date',spec.report.settings['Start Date']);

                this.waitWhileVisible(ETZ_WAIT_SEL,function() {
                    fillEtzForm(this,'End Date',spec.report.settings['End Date']);

                    this.waitWhileVisible(ETZ_WAIT_SEL,function() {
                        submitEtzReport(this,spec);
                    });
                });
            }); 
        }
    }
    ,"Invoice & Payment Summary": {
        linkXpath: '//td[span/text()="' + "Invoice & Payment Summary" + '"]/following-sibling::td/a'     
        ,func: function(spec){
            this.echo("starting Invoice & Payment Summary form entry");
            this.waitForSelector(ETZ_VIEW_REPORT_SEL,function(){
                submitEtzReport(this,spec);
               
            }); 
        }
    }
    ,"Timesheet by Hours Detailed": {
        linkXpath: '//td[span/text()="' + "Timesheet by Hours Detailed" + '"]/following-sibling::td/a'     
        ,func: function(spec){
            this.waitForSelector(ETZ_VIEW_REPORT_SEL,function(){
                
                fillEtzSelect(this,'Client',spec.report.settings['Client'],{spaceToNbsp:true});

                this.waitWhileVisible(ETZ_WAIT_SEL,function() {
                    
                    fillEtzForm(this,"StartDate",spec.report.settings['StartDate']);
                    
                    this.waitWhileVisible(ETZ_WAIT_SEL,function() {
                        
                        fillEtzForm(this,"EndDate",spec.report.settings['EndDate']);
                        fillEtzSelect(this,'Report By',spec.report.settings['Report By'],{spaceToNbsp:true});
                        fillEtzSelect(this,'Show Pay or Bill',spec.report.settings['Show Pay or Bill'],{spaceToNbsp:true});
                        submitEtzReport(this,spec);
                    });
 
                });
            }); 
        }
    }
};


// constant selectors
var ETZ_WAIT_SEL = '#rvw_AsyncWait_Wait';
var ETZ_VIEW_REPORT_SEL = 'input[value="View Report"]';

// helper functions for ETZ report input completion
function fillEtzSelect(cas,label,value,options){
    
    options = options || {};
    value = runOptionsOnValue(value,options);
    
    var index = cas.getElementAttribute(
        xpath("//select[@id=(//label[contains(.,'" + label + "')]/@for)]/option[contains(.,'" + value + "')]")
        , 'value'
    );
    cas.echo ("Setting " + label + ": " + index + ' (' +  value + ')');
    var settings = {};
    settings["//select[@id=(//label[contains(.,'" + label + "')]/@for)]"] = '' + index + '';
    cas.fillXPath('form', settings, false);
};

function fillEtzForm(cas,label,value,options){

    options = options || {};
    value = runOptionsOnValue(value,options);

    cas.echo("Setting " + label + ": " + value);
    var settings = {};
    settings["//input[@id=(//label[contains(.,'" + label + "')]/@for)]"] = value;
    cas.fillXPath('form',settings, false);
};

function submitEtzReport(cas,spec, name,options){
    name = name || '';
    cas.echo("Submitting report " + name); 
    capture(cas,"presubmit.png");
    cas.click(ETZ_VIEW_REPORT_SEL);   // submit button
    waitForReport.call(cas,spec);
}
function fillEtzMultiSelect(cas,label,valuearray,options){

    options = options || {};
    
    cas.click(xpath("//td[label[contains(.,'" + label + "')]]/following-sibling::td//input[@type='image']")); 
    
    //td[label[contains(.,'Payroll Run')]]//following-sibling::td//input[@type="image"]
   //  this.clickLabel('(Select All)');                  
    
    cas.echo("" + label + " settings: " + valuearray.join(" ")); 
 
    var clicked = {};
    valuearray.forEach(function(value){
        
        value = runOptionsOnValue(value,options);
        // find all labels that have the value in them, then click them
        cas.echo("" + label + " value: " + value);
        var x = xpath('//label[contains(.,"' + value + '")]');
        if (cas.exists(x)){
            var labels = cas.getElementsInfo(x);
            labels.forEach(function(label){
                var target = '#' + label.attributes.for;
                if(clicked[target]){
                    cas.echo("NOT clicking " + label.text + " (already clicked)");
                } else {
                    clicked[target] = true;
                    cas.echo("clicking " + label.text);
                    cas.click(target);                    
                }
            });
        } else {
            cas.echo("[nothing found to click]");
        }
    });
}

// function to allow values to be tidied before using in xpaths
// useful to replace characters - e.g. &nbsp;
function runOptionsOnValue(value,options){
    options = options || {};
    if(options.spaceToNbsp){value = value.replace(/ /g, '\u00a0')};
    return value;
}

// Procedure FUNCTIONS 



function login(spec){
    var loginURL = URLBASE + '/EtzWeb/Account/Login';
    casper.start(loginURL, function() {
        
        if (spec.login == undefined || spec.login.email == undefined){
            this.echo("No logon email specified!");
            this.echo("Try something like:");
            this.echo(" casperjs.exe casetzrun.js --email=horse@mybpos.net --password=xxxxxxxxx --outputdir=c:/temp/output/fake --capturedir=c:/temp/output/captures");
            //this.echo('Or use etzdownload.checkForPwFile and put details in a cred.js file containing something like this:');
            //this.echo('exports.cred = {\"email\":\"xxxx@mybpos.net","password":""};');
            this.exit();
        }
        this.echo("Sending logon to " + loginURL +" as " + spec.login.email + "");
        this.viewport(1000,1000);

        this.fillSelectors('#A0Login', {
            '#password': spec.login.password 
            ,'#email': spec.login.email
        }, true);
        //this.echo(this.exists('a[href="/EtzWeb/EtzApp/EtzReportMenu.aspx"]'));

    });
    //once the report link is there, main page is loaded
    casper.waitForSelector('a[href="/EtzWeb/EtzApp/EtzReportMenu.aspx"]',function() {
        this.echo("Waited for main page.");
    });
}
function changeAgency(spec){
    //change agency
    // Net Talent Ltd
    casper.then(function() {
        this.echo("Clicking to change agency");
        this.click( '#LoginView1_lblSystemIdent2');
        this.click( '#LoginView1_hypChangeAgencyContext');
    });
    casper.waitForSelector('table#SelectMember' ,function() {
        this.echo("Selecting agency: " + spec.agency);
        this.click( xpath('//td[contains(.,"' + spec.agency + '")]/following-sibling::td/form/button' ) );  //My BPOS Ltd            //mybpos button 
        //capture(this,'casEtz5.png');
    });
    //wait for main page
    casper.waitForSelector('a[href="/EtzWeb/EtzApp/EtzReportMenu.aspx"]',function() {
        currentAgencyName = this.fetchText('#LoginView1_lblSystemIdent2');
        this.echo("On main page. Agency is: " + currentAgencyName);
    });
};

//open report menu
function runReport(spec){
    casper.thenOpen(URLBASE + '/EtzWeb/EtzApp/EtzReportMenu.aspx/', function() {
        this.echo("Opened report menu.");
        capture(this,'casEtz_reportMenu.png');
    });

    //choose report
    // td:has(span:contains("Multi-Batch Pay-Bill Report")) + td>a
    casper.then(function() {   
        this.echo("Selecting report: "+ spec.report.name);

        this.click( xpath(reportSpecs[spec.report.name].linkXpath ) );   // report link
    });

    //get ahold of report popup and run report function
    casper.waitForPopup(/.+/,function() {
        this.echo("Waited for report popup");
        this.withPopup(/.+/,function() {
            reportSpecs[spec.report.name].func.call(this,spec);
        });
    });
};

function waitForReport(spec){
    spec = spec || {};
    spec.reportTimout= spec.reportTimout || 15 * 60 * 1000;
    this.echo("Report running..."); 
    
    this.waitWhileVisible('#rvw_AsyncWait_Wait',function() {
        //capture(this,'casEtz_afterSubmit.png')
        this.echo("Report completed"); 
        saveExcelReport.call(this,spec);

    } ,null,spec.reportTimeout);
}


//this saves an excel report
function saveExcelReport(spec){
    spec = spec || {};
    //spec.filename= spec.filename || "report.xls";
    spec.onComplete = spec.onComplete || function(){};    
    
    var exportUrlBase = this.evaluate(function(){ 
        return $find("rvw")._getInternalViewer().ExportUrlBase; 
    });
    var excelUrl = URLBASE + exportUrlBase + 'EXCEL';
    this.echo("Excel URL: " + excelUrl ); 

    //this.download(excelUrl, 'report.xls'); //times out after 30s
    var xhrs = this.evaluate(function(url,timeout){

        var xhr = new XMLHttpRequest();
        xhr.timeout = timeout;
        xhr.overrideMimeType("text/plain; charset=x-user-defined");
        xhr.open("GET", url);  // synchronous request banned

        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4) {
                //if (xhr.status == 200) { 
                window.xhrResponseText = __utils__.encode(xhr.responseText);
                window.xhrstatus = xhr.status;
                //}
            }
        };                        

        xhr.send(null);
        return true;
    },excelUrl,spec.reportTimeout);
    
    this.echo("XHR download initiated: " + xhrs);

    this.waitFor(function() { 
        return this.getGlobal('xhrstatus') != undefined; 
    }, function() { 
        this.echo('XHR status: ' + this.getGlobal('xhrstatus')); 
        var fn = getFileName(spec);
        //http://phantomjs.org/api/fs/method/is-writable.html 
        if (spec.outputDir){
            fs.changeWorkingDirectory(spec.outputDir);
        }
        this.echo('Working directory: ' + fs.workingDirectory); 
        this.echo('Saving report to: ' + fn); 
        fs.write(fn, decode(this.getGlobal('xhrResponseText')), 'wb');
        //spec.onComplete();
        /*this.echo("test");
        this.echo(this.findAllPopupsByRegExp(/.+/).join("\n"));
        var test = this.evaluate(function(){
            return window.close();
        });
        
        casper.wait(3000, function() {
            this.echo("Waited 3000ms");
            this.echo("AFTER " + test); 
            this.echo("POPUPS: " + casper.findAllPopupsByRegExp(/.+/).join("\n"));
        });*/        
        //this.exit();
        //this.close();

        
    },null,spec.reportTimeout);
}


var getFileName = function(spec){

    var filename = spec.filename || 'ETZ_{{reportname}}_{{agency}}_{{date}}_{{time}}.xls';
    var reportName = spec.report.name.replace(/[^a-z0-9A-Z]/gi,'') || 'report';
    var agency = spec.agency.replace(/[^a-z0-9A-Z]/gi,'') || 'defaultagency';
    // reportName
    return filename
        .replace(/{{reportname}}/gi,reportName)
        .replace(/{{agency}}/gi,agency)
        .replace(/{{date}}/gi,moment().format('YYYYMMDD'))
        .replace(/{{time}}/gi,moment().format('HHmmss'));

};

var capture = function(casper, filename){
    if (captureEnabled || casper.cli.options['capturedir']){      // casper.cli.options['capturedir']
        
        var cdir = casper.cli.options['capturedir'] || captureDir;
        var dir = cdir ? cdir + "/": "" ;
        casper.echo("Capturing to " + dir + filename);
        casper.capture(dir + filename);
    }
} ;

var BASE64_ENCODE_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
var BASE64_DECODE_CHARS = [
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, 62, -1, -1, -1, 63,
    52, 53, 54, 55, 56, 57, 58, 59, 60, 61, -1, -1, -1, -1, -1, -1,
    -1,  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13, 14,
    15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, -1, -1, -1, -1, -1,
    -1, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
    41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, -1, -1, -1, -1, -1
];

/**
 * Decodes a base64 encoded string. Succeeds where window.atob() fails.
 *
 * @param  String  str  The base64 encoded contents
 * @return string
 */
var decode = function decode(str) {
    /*eslint max-statements:0, complexity:0 */
    var c1, c2, c3, c4, i = 0, len = str.length, out = "";
    while (i < len) {
        do {
            c1 = BASE64_DECODE_CHARS[str.charCodeAt(i++) & 0xff];
        } while (i < len && c1 === -1);
        if (c1 === -1) {
            break;
        }
        do {
            c2 = BASE64_DECODE_CHARS[str.charCodeAt(i++) & 0xff];
        } while (i < len && c2 === -1);
        if (c2 === -1) {
            break;
        }
        out += String.fromCharCode(c1 << 2 | (c2 & 0x30) >> 4);
        do {
            c3 = str.charCodeAt(i++) & 0xff;
            if (c3 === 61) {
                return out;
            }
            c3 = BASE64_DECODE_CHARS[c3];
        } while (i < len && c3 === -1);
        if (c3 === -1) {
            break;
        }
        out += String.fromCharCode((c2 & 0XF) << 4 | (c3 & 0x3C) >> 2);
        do {
            c4 = str.charCodeAt(i++) & 0xff;
            if (c4 === 61) {
                return out;
            }
            c4 = BASE64_DECODE_CHARS[c4];
        } while (i < len && c4 === -1);
        if (c4 === -1) {
            break;
        }
        out += String.fromCharCode((c3 & 0x03) << 6 | c4);
    }
    return out;
};


