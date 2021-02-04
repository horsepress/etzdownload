/*  
  This script uses casper.js (http://casperjs.org/) to download reports from ETZ timesheets 
  (http://timesheetz.net/)

  you can run this like to specify a username and password on the command line:

  bin\casperjs.exe etzDownloadSample.js --email=xxxxx@xxxx.xxx --password=xxxxxxxxx --outputdir=c:/temp/output/fake --capturedir=c:/temp/output/captures

  or just run with no options and enable settings here as per the comments below
  bin\casperjs.exe casetzrun.js
  
*/
var casper = require('casper');
var etzdownload = require("etzdownload");
var etzcasper = etzdownload.setCasper(casper);    //pass reference to casper. Necessary.
//var moment = require('moment');    // might be useful for dynamic date settings

//check for a password file. Default location is ./cred.js (ie a file called cred.js in the working dir)
// it would contain something like this:
// exports.cred = {"email":"phughson@mybpos.net" ,"password":"xxxxxxx"};
etzdownload.checkForPwFile('./cred.js');    

/*  // alternatively we can specify a password here (or in command line as above)
etzdownload.setLogin({
    email:'xxxxx@xxxx.xxx'    
    ,password:'xxxxxxxx'    
});
// */

//might be useful to set output dir in the file
//etzdownload.setOutputDir("c:/temp/output");

//You can also turn on and set a directory to put screen prints in. Useful for debugging.
//etzdownload.setCapture(true,"c:/temp/output/captures");


//to do the same report for multiple agencies:    
["International worldwide solutions","ESPC Direct Limited"].forEach(function(agency){

   etzdownload.getReport({
        //agency:"First People Solutions Limited"
        agency:agency
        ,report: {
            name:"Timesheet by Hours Detailed"
            ,settings:{
                'Client': '[TRUMP]'                    // best just to use code as these are unique.
                ,'StartDate': '21/05/2016'   // or to set from command line like this --startdate=22/05/2016: etzcasper.cli.options["startdate"]   
                ,'EndDate': '27/05/2016'
                ,'Report By': 'Period End Date'
                ,'Show Pay or Bill': 'Bill'
            }
        }
    });

    /*
    etzdownload.getReport({
        agency:agency
        ,report:{
            name:"Invoice & Payment Summary"
            ,settings:{
                // go with defaults
            }
        }
    });
    
    etzdownload.getReport({
        agency:agency
        ,report: {
            name:"Multi-Batch Pay-Bill Report"
            ,settings:{
                'Pay Run Type': 'Sales'
                ,'Payroll Run' : ['05/2016','252314','05/2016']      // can use a fragment here to match multiply
            }
        }
    });

 

    etzdownload.getReport({
        agency:agency
        ,report: {
            name:"Sales Invoice Day Book"
            ,settings:{
                'Start Date': '01/05/2016'    // could use moment.js to generate dates
                ,'End Date': '30/05/2016'     // e.g.  moment().day(6-(2*7)).format('DD/MM/YYYY')
                                                // or to set from command line like this --startdate=22/05/2016: 
                                                // etzcasper.cli.options["startdate"]
                
                
            }
        }   
    });
    // */

});


etzdownload.exit();     //include this line to make casper exit once all reports finished



/* DETAILS OF REPORT PARAMS
etzdownload.getReport({
    login:{
        email: 'blah@domain'
        ,password: 'xxxxxxxx'
    }
    ,agency:"Org Tennis"    //All or unique part of an agency name, or can put agency number e.g. "207"
    ,filename:"test_salesInvoiceDayBook_{{date}}_{{time}}.xls"    // default is ETZ_{{reportname}}_{{agency}}_{{date}}_{{time}}
    ,outputDir: "C:/temp"           // where to put saved excel file. This will override a command line / global setting
    ,report:{
        name:"Sales Invoice Day Book"           //
        ,settings:{
            'Start Date': '01/05/2016'
            ,'End Date': '30/05/2016'
        }
    }
});
*/

