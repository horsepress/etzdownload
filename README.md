# etzdownload

This Casper.js module handles the automation of report downloads from http://www.timesheetz.net/, the ETZ timesheets site (http://timesheetz.net)

This is useful because the site doesn't (seem to) provide a REST interface or otherwise for automating downloads, and they become tedious to download regularly en masse.

Essentially you can make a javascript file which specifies the parameters for various reports and then run this to automate the download.

To do
=====

* With a little effort the functions for handling the various input types could be generalised a bit, allowing the download spec to handle any kind of report, rather than just those with their own handling function.
