# spfx-time-mngm-filter-to-excel

This SPFX is a completions to my Time Management App

The report app was too simple to some clients so I created this SPFX with better results

Has download to excel button

# ASSETS

* `EmployeesHoursApp.msapp` - super simple Employees app to click "start", "end", by subject (subject can be employee name, client name, ect.)
* `EmployeeReportApp.msapp` - report app, simple
* `EmployeesHoursApp-Start_20210107021941.zip` - start button flow for Employees app
* `EmployeesHoursApp-End_20210107021926.zip` - end button flow for Employees app

# Setup
* SPSite should have 2 lists, 
    * `TimeMngApp-WorkSubject`, no cols (only `Title`)
    * `TimeMngApp-Hours`
        * `EndTime` - DateTime
        * `TotalTime` - calculated with formula `=TEXT(EndTime-Created,"hh:mm")`
* set proper regional settings 



# Deploy

1. run `gulp build`
2. run `gulp bundle --ship`
3. run `gulp package-solution --ship`

or use `gulp serve` with `https://{tenant}.sharepoint.com/sites/{yourSiteName}/_layouts/15/workbench.aspx`

# I'm in YouTube

[link to playlist with this app](https://www.youtube.com/watch?v=B1st9aDk_FU&list=PLbZpz8SE2dlceqH0kuwSjHMTfn5PWzLGp&index=2&ab_channel=ArielRubinstein)

### jump to code

[SpfxTimeMngmFilterToExcelWebPart.ts](https://github.com/bresleveloper/SPFX-TimeMngmApp/blob/master/src/webparts/spfxTimeMngmFilterToExcel/SpfxTimeMngmFilterToExcelWebPart.ts)

#### fix for RTL in excel   
in `ExcellentExport` I added option for RTL, [see issue](https://github.com/jmaister/excellentexport/issues/393)

in our `ASSETS` folder i added the modified `excellentexport.js` minified version, you can change in your `node_module` folder

