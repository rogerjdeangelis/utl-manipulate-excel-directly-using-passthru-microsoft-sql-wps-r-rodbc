%let pgm=utl-manipulate-excel-directly-using-passthru-microsoft-sql-wps-r-rodbc;

PROBLEM

   Given two excel tables(named ranges) ,'males' and 'females' do the following
   Calculate then mean age by sex using MS SQL inside an excel workbook

   /**************************************************************************************************************************/
   /*                                                              |                                                         */
   /* EXCEL WORKBOOKd:/xls/have.xlsx (two named ranges)            | PROCESS (use passthru to excel SQL for mean age by sex) */
   /*                                                              |                                                         */
   /*   +-------------                 +-------------              |                                                         */
   /* 1 |  FEMALES   | =>named range 1 |  MALES     | =>named range| select sex, avg(age) as avgAge from males group by sex  */
   /*   +------------+                 +------------+              | union                                                   */
   /*                                                              | select sex, avg(age) as avgAge from females group by sex*/
   /*   +-------------------------+    +-------------------------+ |                                                         */
   /*   |     A    |   B   |   C  |    |     A    |   B   |   C  | | OUTPUT                                                  */
   /*   +-------------------------+    +-------------------------+ |                                                         */
   /* 1 | NAME     |  SEX  |  AG  |  1 | NAME     |  SEX  |  AG  | | SD1.WANT total obs=2                                    */
   /*   +----------+-------+------+    +----------+-------+------+ |                                                         */
   /* 2 | Alice    |   F   |  13  |  2 | Alice    |   M   |  14  | | Obs    SEX     AVGAGE                                   */
   /*   +----------+-------+------+    +----------+-------+------+ |                                                         */
   /* 3 | Barbara  |   F   |  13  |  3 | Barbara  |   M   |  14  | |  1      F     13.3333                                   */
   /*   +----------+-------+------+    +----------+-------+------+ |  2      M     13.3333                                   */
   /* 4 | Carol    |   F   |  14  |  4 | Carol    |   M   |  12  | |                                                         */
   /*   +----------+-------+------+    +----------+-------+------+ |                                                         */
   /*                    mean=13.33                     mean=13.33 |                                                         */
   /*                                                              |                                                         */
   /**************************************************************************************************************************/

  PROCESS

     Create wps datasets to load into an excel workbook
     Create excel workbook with named ranges
     Programatically create an ODBC Data Source File DSN
     Connect R ro data source
     Exceute Microsoft SQL inside sql
     MS SQL documentations on end

github
https://tinyurl.com/374jvs57
https://github.com/rogerjdeangelis/utl-manipulate-excel-directly-using-passthru-microsoft-sql-wps-r-rodbc

/*                   _                       _                                  _
(_)_ __  _ __  _   _| |_    _____  _____ ___| |  _ __   __ _ _ __ ___   ___  __| |  _ __ __ _ _ __   __ _  ___
| | `_ \| `_ \| | | | __|  / _ \ \/ / __/ _ \ | | `_ \ / _` | `_ ` _ \ / _ \/ _` | | `__/ _` | `_ \ / _` |/ _ \
| | | | | |_) | |_| | |_  |  __/>  < (_|  __/ | | | | | (_| | | | | | |  __/ (_| | | | | (_| | | | | (_| |  __/
|_|_| |_| .__/ \__,_|\__|  \___/_/\_\___\___|_| |_| |_|\__,_|_| |_| |_|\___|\__,_| |_|  \__,_|_| |_|\__, |\___|
        |_|                                                                                         |___/
*/

/*----                                                                   ----*/
/*---- create wps datasets to load into excel named ranges(tables)       ----*/
/*----                                                                   ----*/

%utl_submit_wps64x('
options validvarname=upcase;
libname sd1 "d:/sd1";
proc sql;
      create
        table have
           (
            NAME  Char(8)
           ,SEX   Char(1)
           ,AGE   NUMERIC
           );
      insert into have
    values("Alfred ","M",14)
    values("Alice  ","F",13)
    values("Barbara","F",13)
    values("Carol  ","F",14)
    values("Henry  ","M",14)
    values("James  ","M",12)
    values("Jane   ","F",12)
;quit;
data sd1.males sd1.females;
  set have;
  if sex='M' then output sd1.males;
             else output sd1.females;

run;quit;
');

/**************************************************************************************************************************/
/*                                                                                                                        */
/*  WPS DATASETS TO LOAD INTO EXCEL                                                                                       */
/*                                                                                                                        */
/*  SD1.FEMALES total obs=3                  SD1.MALES total obs=3                                                        */
/*                                                                                                                        */
/*  Obs     NAME      SEX    AGE             Obs     NAME     SEX    AGE                                                  */
/*                                                                                                                        */
/*   1     Alice       F      13              1     Alfred     M      14                                                  */
/*   2     Barbara     F      13              2     Henry      M      14                                                  */
/*   3     Carol       F      14 mean=13.33   3     James      M      12  mean=13.33                                      */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*----                                                                   ----*/
/*---- Create named ramges in excel workbook d:/xls/have.xlsx            ----*/
/*----                                                                   ----*/

%utlfkil(d:/xls/have.xlsx); * delete if exist - it works with an existing workbook;

%utl_submit_wps64x('
libname sd1 "d:/sd1";
proc r;
 export data=sd1.females  r=females;
 export data=sd1.males  r=males;
submit;
library(openxlsx);
wb <- createWorkbook("d:/xls/have.xlsx");
addWorksheet(wb, "sheet 1");
writeData(wb, sheet = 1, x = females, startCol = 1, startRow = 1);
createNamedRegion(
  wb = wb,
  sheet = 1,
  name = "females",
  rows = 1:(nrow(females) + 1),
  cols = 1:ncol(females)
);
writeData(wb, sheet = 1, x = males, startCol = 5, startRow = 1);
createNamedRegion(
  wb = wb,
  sheet = 1,
  name = "males",
  rows = 1:(nrow(males) + 1),
  cols = 5:(ncol(males) + 4)
);
saveWorkbook(wb,"d:/xls/have.xlsx", overwrite = TRUE);
endsubmit;
');

/**************************************************************************************************************************/
/*                                                                                                                        */
/* EXCEL WORKBOOKd:/xls/have.xlsx (two named ranges)                                                                      */
/*                                                                                                                        */
/*   +-------------                 +-------------                                                                        */
/* 1 |  FEMALES   | =>named range 1 |  MALES     | =>named range                                                          */
/*   +------------+                 +------------+                                                                        */
/*                                                                                                                        */
/*   +-------------------------+    +-------------------------+                                                           */
/*   |     A    |   B   |   C  |    |     A    |   B   |   C  |                                                           */
/*   +-------------------------+    +-------------------------+                                                           */
/* 1 | NAME     |  SEX  |  AG  |  1 | NAME     |  SEX  |  AG  |                                                           */
/*   +----------+-------+------+    +----------+-------+------+                                                           */
/* 2 | Alice    |   F   |  13  |  2 | Alice    |   M   |  14  |                                                           */
/*   +----------+-------+------+    +----------+-------+------+                                                           */
/* 3 | Barbara  |   F   |  13  |  3 | Barbara  |   M   |  14  |                                                           */
/*   +----------+-------+------+    +----------+-------+------+                                                           */
/* 4 | Carol    |   F   |  14  |  4 | Carol    |   M   |  12  |                                                           */
/*   +----------+-------+------+    +----------+-------+------+                                                           */
/*                    mean=13.33                     mean=13.33                                                           */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

/*----                                                                   ----*/
/*---- CREATE ODBC DSN                                                   ----*/
/*----                                                                   ----*/

/*---- HANDLE THE LONG LINE ISSUE IN POWERSHELL                          ----*/

%let longline=%str(Add-OdbcDsn -Name 'have' -DriverName 'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)'
 -DsnType 'User' -Platform '64-bit' -SetPropertyValue 'Dbq=d:\xls\have.xlsx');

options ls=255;
%put "&=longline";

/*---- CREATE ODBC DSN                                                   ----*/

%utl_submit_ps64("
Remove-OdbcDsn -Name 'have' -DsnType 'User' -Platform '64-bit';
&longline;
Get-OdbcDsn;
");

/*----                                                                   ----*/
/*---- EXECUTE MS SQL INSIDE EXCEL                                       ----*/
/*----                                                                   ----*/

proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

%utl_submit_wps64x('
libname sd1 "d:/sd1";
proc r;
submit;
library(RODBC);
ch <- odbcConnect("have");
sqlResult <- sqlQuery(ch,
  "select sex, avg(age) as avgAge from males group by sex
   union
   select sex, avg(age) as avgAge from females group by sex
                           ";);
sqlResult;
odbcClose(ch);
endsubmit;
import data=sd1.want r=sqlResult;
');

proc print data=sd1.want;
run;quit;


*                          _
 _ __ ___  ___   ___  __ _| |
| '_ ` _ \/ __| / __|/ _` | |
| | | | | \__ \ \__ \ (_| | |
|_| |_| |_|___/ |___/\__, |_|
                        |_|
;


https://ss64.com/access/

a
  Abs             The absolute value of a number (nore negative sn).
 .AddMenu         Add a custom menu bar/shortcut bar.
 .AddNew          Add a new record to a recordset.
 .ApplyFilter     Apply a filter clause to a table, form, or report.
  Array           Create an Array.
  Asc             The Ascii code of a character.
  AscW            The Unicode of a character.
  Atn             Display the ArcTan of an angle.
  Avg (SQL)       Average.
b
 .Beep (DoCmd)    Sound a tone.
 .BrowseTo(DoCmd) Navate between objects.
c
  Call            Call a procedure.
 .CancelEvent (DoCmd) Cancel an event.
 .CancelUpdate    Cancel recordset changes.
  Case            If Then Else.
  CBool           Convert to boolean.
  CByte           Convert to byte.
  CCur            Convert to currency (number)
  CDate           Convert to Date.
  CVDate          Convert to Date.
  CDbl            Convert to Double (number)
  CDec            Convert to Decimal (number)
  Choose          Return a value from a list based on position.
  ChDir           Change the current directory or folder.
  ChDrive         Change the current drive.
  Chr             Return a character based on an ASCII code.
 .ClearMacroError (DoCmd) Clear MacroError.
 .Close (DoCmd)           Close a form/report/window.
 .CloseDatabase (DoCmd)   Close the database.
  CInt                    Convert to Integer (number)
  CLng                    Convert to Long (number)
  Command                 Return command line option string.
 .CopyDatabaseFile(DoCmd) Copy to an SQL .mdf file.
 .CopyObject (DoCmd)      Copy an Access database object.
  Cos                     Display Cosine of an angle.
  Count (SQL)             Count records.
  CSng             Convert to Single (number.)
  CStr             Convert to String.
  CurDir           Return the current path.
  CurrentDb        Return an object variable for the current database.
  CurrentUser      Return the current user.
  CVar             Convert to a Variant.
d
  Date             The current date.
  DateAdd          Add a time interval to a date.
  DateDiff         The time difference between two dates.
  DatePart         Return part of a given date.
  DateSerial       Return a date given a year, month, and day.
  DateValue        Convert a string to a date.
  DAvg             Average from a set of records.
  Day              Return the day of the month.
  DCount           Count the number of records in a table/query.
  Delete (SQL)          Delete records.
 .DeleteObject (DoCmd)  Delete an object.
  DeleteSetting         Delete a value from the users registry
 .DoMenuItem (DoCmd)    Display a menu or toolbar command.
  DFirst           The first value from a set of records.
  Dir              List the files in a folder.
  DLast            The last value from a set of records.
  DLookup          Get the value of a particular field.
  DMax             Return the maximum value from a set of records.
  DMin             Return the minimum value from a set of records.
  DoEvents         Allow the operating system to process other events.
  DStDev           Estimate Standard deviation for domain (subset of records)
  DStDevP          Estimate Standard deviation for population (subset of records)
  DSum             Return the sum of values from a set of records.
  DVar             Estimate variance for domain (subset of records)
  DVarP            Estimate variance for population (subset of records)
e
 .Echo             Turn screen updating on or off.
  Environ          Return the value of an OS environment variable.
  EOF              End of file input.
  Error            Return the error message for an error No.
  Eval             Evaluate an expression.
  Execute(SQL/VBA) Execute a procedure or run SQL.
  Exp              Exponential e raised to the nth power.
f
  FileDateTime      Filename last modified date/time.
  FileLen           The size of a file in bytes.
 .FindFirst/Last/Next/Previous Record.
 .FindRecord(DoCmd) Find a specific record.
  First (SQL)       Return the first value from a query.
  Fix               Return the integer portion of a number.
  For               Loop.
  Format            Format a Number/Date/Time.
  FreeFile          The next file No. available to open.
  From              Specify the table(s) to be used in an .
  FV                Future Value of an annuity.
g
  GetAllSettings    List the settings saved in the registry.
  GetAttr           Get file/folder attributes.
  GetObject         Return a reference to an ActiveX object
  GetSetting        Retrieve a value from the users registry.
  form.GoToPage     Move to a page on specific form.
 .GoToRecord (DoCmd)Move to a specific record in a dataset.
h
  Hex               Convert a number to Hex.
  Hour              Return the hour of the day.
 .Hourglass (DoCmd) Display the hourglass icon.
  HyperlinkPart     Return information about data stored as a hyperlink.
i
  If Then Else      If-Then-Else
  IIf               If-Then-Else function.
  Input             Return characters from a file.
  InputBox          Prompt for user input.
  Insert (SQL)      Add records to a table (append query).
  InStr             Return the position of one string within another.
  InstrRev          Return the position of one string within another.
  Int               Return the integer portion of a number.
  IPmt              Interest payment for an annuity
  IsArray           Test if an expression is an array
  IsDate            Test if an expression is a date.
  IsEmpty           Test if an expression is Empty (unassned).
  IsError           Test if an expression is returning an error.
  IsMissing         Test if a missing expression.
  IsNull            Test for a NULL expression or Zero Length string.
  IsNumeric         Test for a valid Number.
  IsObject          Test if an expression is an Object.
L
  Last (SQL)        Return the last value from a query.
  LBound            Return the smallest subscript from an array.
  LCase             Convert a string to lower-case.
  Left              Extract a substring from a string.
  Len               Return the length of a string.
  LoadPicture       Load a picture into an ActiveX control.
  Loc               The current position within an open file.
 .LockNavationPane(DoCmd) Lock the Navation Pane.
  LOF               The length of a file opened with Open()
  Log               Return the natural logarithm of a number.
  LTrim             Remove leading spaces from a string.
m
  Max (SQL)         Return the maximum value from a query.
 .Maximize (DoCmd)  Enlarge the active window.
  Mid               Extract a substring from a string.
  Min (SQL)         Return the minimum value from a query.
 .Minimize (DoCmd)  Minimise a window.
  Minute            Return the minute of the hour.
  MkDir             Create directory.
  Month             Return the month for a given date.
  MonthName         Return  a string representing the month.
 .Move              Move through a Recordset.
 .MoveFirst/Last/Next/Previous Record
 .MoveSize (DoCmd)  Move or Resize a Window.
  MsgBox            Display a message in a dialogue box.
n
  Next              Continue a for loop.
  Now               Return the current date and time.
  Nz                Detect a NULL value or a Zero Length string.
o
  Oct               Convert an integer to Octal.
  OnClick, OnOpen   Events.
 .OpenForm (DoCmd)  Open a form.
 .OpenQuery (DoCmd) Open a .
 .OpenRecordset         Create a new Recordset.
 .OpenReport (DoCmd)    Open a report.
 .OutputTo (DoCmd)      Export to a Text/CSV/Spreadsheet file.
p
  Partition (SQL)       Locate a number within a range.
 .PrintOut (DoCmd)      Print the active object (form/report etc.)
q
  Quit                  Quit Microsoft Access
r
 .RefreshRecord (DoCmd) Refresh the data in a form.
 .Rename (DoCmd)        Rename an object.
 .RepaintObject (DoCmd) Complete any pending screen updates.
  Replace               Replace a sequence of characters in a string.
 .Re               Re the data in a form or a control.
 .Restore (DoCmd)       Restore a maximized or minimized window.
  RGB                   Convert an RGB color to a number.
  Rht                 Extract a substring from a string.
  Rnd                   Generate a random number.
  Round                 Round a number to n decimal places.
  RTrim                 Remove trailing spaces from a string.
 .RunCommand            Run an Access menu or toolbar command.
 .RunDataMacro (DoCmd)  Run a named data macro.
 .RunMacro (DoCmd)      Run a macro.
 .RunSavedImportExport (DoCmd) Run a saved import or export specification.
 .RunSQL (DoCmd)        Run an SQL .
s
 .Save (DoCmd)          Save a database object.
  SaveSetting           Store a value in the users registry
 .SearchForRecord(DoCmd) Search for a specific record.
  Second                Return the seconds of the minute.
  Seek                  The position within a file opened with Open.
  Select (SQL)          Retrieve data from one or more tables or queries.
  Select Into (SQL)     Make-table .
  Select-Sub (SQL) Sub.
 .SelectObject (DoCmd)  Select a specific database object.
 .SendObject (DoCmd)    Send an email with a database object attached.
  SendKeys              Send keystrokes to the active window.
  SetAttr               Set the attributes of a file.
 .SetDisplayedCategories (DoCmd)  Change Navation Pane display options.
 .SetFilter (DoCmd)     Apply a filter to the records being displayed.
  SetFocus              Move focus to a specified field or control.
 .SetMenuItem (DoCmd)   Set the state of menubar items (enabled /checked)
 .SetOrderBy (DoCmd)    Apply a sort to the active datasheet, form or report.
 .SetParameter (DoCmd)  Set a parameter before opening a Form or Report.
 .SetWarnings (DoCmd)   Turn system messages on or off.
  Sgn                   Return the sn of a number.
 .ShowAllRecords(DoCmd) Remove any applied filter.
 .ShowToolbar (DoCmd)   Display or hide a custom toolbar.
  Shell                 Run an executable program.
  Sin                   Display Sine of an angle.
  SLN                   Straht Line Depreciation.
  Space                 Return a number of spaces.
  Sqr                   Return the square root of a number.
  StDev (SQL)           Estimate the standard deviation for a population.
  Str                   Return a string representation of a number.
  StrComp               Compare two strings.
  StrConv               Convert a string to Upper/lower case or Unicode.
  String                Repeat a character n times.
  Sum (SQL)             Add up the values in a  result set.
  Switch                Return one of several values.
  SysCmd                Display a progress meter.
t
  Top 1 *               Get first rpw
  Tan                   Display Tangent of an angle.
  Time                  Return the current system time.
  Timer                 Return a number (single) of seconds since midnht.
  TimeSerial            Return a time given an hour, minute, and second.
  TimeValue             Convert a string to a Time.
 .TransferDatabase (DoCmd)      Import or export data to/from another database.
 .TransferSharePointList(DoCmd) Import or link data from a SharePoint Foundation site.
 .TransferSpreadsheet (DoCmd)   Import or export data to/from a spreadsheet file.
 .TransferSQLDatabase (DoCmd)   Copy an entire SQL Server database.
 .TransferText (DoCmd)          Import or export data to/from a text file.
  Transform (SQL)       Create a crosstab .
  Trim                  Remove leading and trailing spaces from a string.
  TypeName              Return the data type of a variable.
u
  UBound                Return the largest subscript from an array.
  UCase                 Convert a string to upper-case.
  Undo                  Undo the last data edit.
  Union (SQL)           Combine the results of two SQL queries.
  Update (SQL)          Update existing field values in a table.
 .Update                Save a recordset.
v
  Val                   Extract a numeric value from a string.
  Var (SQL)             Estimate variance for sample (all records)
  VarP (SQL)            Estimate variance for population (all records)
  VarType               Return a number indicating the data type of a variable.
w
  Weekday               Return the weekday (1-7) from a date.
  WeekdayName           Return the day of the week.
y
  Year                  Return the year for a given date.

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/






































































































































































































































































































































































































































































































































































































           ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;













%utlfkil(d:/xls/have.xlsx); * delete if exist - it works with an existing workbook;

%utl_submit_wps64x('
libname sd1 "d:/sd1";
proc r;
 export data=sd1.females  r=females;
 export data=sd1.males  r=males;
submit;
library(openxlsx);
wb <- createWorkbook("d:/xls/have.xlsx");
addWorksheet(wb, "sheet 1");
writeData(wb, sheet = 1, x = iris, startCol = 2, startRow = 2);
createNamedRegion(
  wb = wb,
  sheet = 1,
  name = "iris1",
  rows = 1:(nrow(iris) + 1),
  cols = 1:5
);
saveWorkbook(wb,"d:/xls/gender.xlsx", overwrite = TRUE);
endsubmit;
');





writeData(wb, sheet = 1, x = iris, name = "iris1", startCol = 10);
getNamedRegions(wb)
getNamedRegions(out_file)

## delete one
deleteNamedRegion(wb = wb, name = "iris2")
getNamedRegions(wb)

## read named regions
df <- read.xlsx(wb, namedRegion = "iris")
head(df)

df <- read.xlsx(out_file, namedRegion = "iris2")
head(df)

## End(Not run)









































































options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.males sd1.females;
  set sashelp.class(obs=6);
   keep name sex age;
  if sex="M" then output sd1.males;
  else output sd1.females;
run;quit;

SD1.FEMALES total obs=3

Obs     NAME      SEX    AGE

 1     Alice       F      13
 2     Barbara     F      13
 3     Carol       F      14

SD1.MALES total obs=3

Obs     NAME     SEX    AGE

 1     Alfred     M      14
 2     Henry      M      14
 3     James      M      12


%utlfkil(d:/xls/gender.xlsx); * delete if exist - it works with an existing workbook;

%utl_submit_wps64x('
libname sd1 "d:/sd1";
proc r;
 export data=sd1.females  r=females;
 export data=sd1.males  r=males;
submit;
Sys.setenv(JAVA_HOME="C:/Program Files/Java/jre-1.8");
system("java -version");
endsubmit;');

options(java.home="C:\\Program Files\\Java\\jre-1.8");
Sys.setenv(JAVA_HOME="C:\\Program Files\\Java\\jre-1.8\\");
library(rJava);
library(XLConnect);
wb <- loadWorkbook("d:/xls/genders.xlsx", create = TRUE);
createSheet(wb, name = "gender");
createName(wb, name = "females", formula = "gender!$B$3");
writeNamedRegion(wb, females, name = "females");

createName(wb, name = "males", formula = "gender!$G$3");
writeNamedRegion(wb, males, name = "males");
saveWorkbook(wb);
endsubmit;
');


Sys.setenv(JAVA_HOME="C:\\Program Files\\Java\\jre-1.8\\")

JAVA_VERSION="1.8.0_381"

JAVA_RUNTIME_VERSION="1.8.0_381-b09
"
OS_NAME="Windows"
OS_VERSION="5.2"
OS_ARCH="amd64"
SOURCE=".:git:543f7df00d44+"
BUILD_TYPE="commercial"



 needed to install the Java SE Development Kit for rJava to work (should have read the package's documents)
and then set the JAVA_HOME path to the jre folder inside "jdk1.8.0_121". Finally restart RStudio and everything works fine (I can load the rJava package).

Sorry for the duplicate.


 http://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html
 https://login.oracle.com/mysso/signon.jsp






 WANT excel sheet, GENDER, with two named ranges males starting at A3 and females at G3
 ==============================================================

 d:/xls/gender.xlsx

     +---------------------+---------------------------------+-------+
     |  A  |  B    |  C    |  D  |  E  |  F    |  G    |  H  |  D    |
     +---------------------+---------------------------------+-------+
 1   |     |       |       |     |     |       |       |     |       |
     |-----+-------+-------|-----+-----+-------+-------+-----|-------|
 2   |     |       |       |     |     |       |       |     |       |
     |-----+-------+-------+-----+-----+-------+-------+-----+-------+
 3   |     |NAME   |SEX    |AGE  |     |       |NAME   |SEX  |AGE    |
     |-----+-------+-------|-----+-----+-------+-------+-----|-------|
 4   |     |Alfred |M      |13   |     |       |Alice  |F    |14     |
     |-----+-------+-------+-----+-----+-------+-------+-----+-------+
 5   |     |Alex   |M      |13   |     |       |Barbara|F    |14     |
     |-----+-------+-------+-----+-----+-------+-------+-----+-------+
 6   |     |JAMES  |M      |14   |     |       |Carol  |F    |12     |
     -----------------------------------------------------------------
 ...

 [GENDER]























































https://github.com/rogerjdeangelis/utl-create-sas-table-using-an-excel-rang

%utlfkil(d:/xls/have.xls);

%utl_submit_wps64x('
options validvarname=upcase;
libname xls excel "d:/xls/have.xls";
proc sql;
      create
        table xls.have
           (
            NAME  Char(8)
           ,SEX   Char(1)
           ,AGE   NUMERIC
           );
      insert into xls.have
    values("Alfred ","M",14)
    values("Alice  ","F",13)
    values("Barbara","F",13)
    values("Carol  ","F",14)
    values("Henry  ","M",14)
    values("James  ","M",12)
    values("Jane   ","F",12)
;quit;
libname xls clear;
');


/**************************************************************************************************************************/
/*                                                                                                                        */
/*    EXCEL WORKBOOKd:/xls/want.xlsx                                                                                      */
/*                                                                                                                        */
/*     +-------------                                                                                                     */
/*  1  |  have      |  ==> named range                                                                                    */
/*     +------------+                                                                                                     */
/*                                                                                                                        */
/*     +--------------------------------------+                                                                           */
/*     |     A      |    B       |     C      |                                                                           */
/*     +--------------------------------------+                                                                           */
/*  1  | NAME       |   SEX      |    AGE     |                                                                           */
/*     +------------+------------+------------+                                                                           */
/*  2  | Alice      |    F       |    12      |                                                                           */
/*     +------------+------------+------------+                                                                           */
/*  3  | Mary       |    M       |    16      |                                                                           */
/*     +------------+------------+------------+                                                                           */
/*  4  | Tom        |    M       |    15      |                                                                           */
/*     +------------+------------+------------+                                                                           */
/*  5  | John       |    M       |    14      |                                                                           */
/*     +------------+------------+------------+                                                                           */
/*  6  | Jane       |    F       |    13      |                                                                           */
/*     +------------+------------+------------+                                                                           */
/*                                                                                                                        */
/**************************************************************************************************************************/

%utl_submit_ps64('
Add-OdbcDsn -Name "s8fin" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue "Dbq=d:\xls\s8fin.xlsx";
');

%utl_submit_wps64x('
proc r;
submit;
library(RODBC);
ch <- odbcConnect("s8fin");
sqlResult <- sqlQuery(ch, "select * from have";);
sqlResult;
odbcClose(ch);
endsubmit;
');


options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.males sd1.females;
  set sashelp.class(obs=6);
   keep name sex age;
  if sex="M" then output sd1.males;
  else output sd1.females;
run;quit;


%utlfkil(d:/xls/gender.xlsx); * delete if exist - it works with an existing workbook;

%utl_submit_wps64x('
library(XLConnect);
proc r;
 export data=sd1,females  r=females;
 export data=sd1,males    r=males  ;
submit;
wb <- loadWorkbook("d:/xls/genders.xlsx", create = TRUE);
createSheet(wb, name = "gender");
createName(wb, name = "females", formula = "gender!$B$3");
writeNamedRegion(wb, females, name = "females");

createName(wb, name = "males", formula = "gender!$G$3");
writeNamedRegion(wb, males, name = "males");
saveWorkbook(wb);
');











































%utl_submit_wps64x('
proc r;
submit;
library(RODBC);
qry<-"select * from have";
MyExcelData <- sqlQuery(odbcConnect("s8fin"),qry, na.strings = "NA", as.is = T);  odbcCloseAll();
MyExcelData;
endsubmit;
');


%utl_submit_wps64x('
proc r;
submit;
library(RODBC);
MyExcelData <- sqlQuery(odbcConnect("s8fin"),"select * from have", na.strings = "NA", as.is = T);  odbcCloseAll();
MyExcelData;
endsubmit;
');



                                    ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;

































conn <- odbcConnect("CData Excel Source")



STEP 1: Search ODBC in the Start Menu search and open ODBC Data Source Administrator (64-bit)
Step 2: Select Add under the User DSN.
Step 3: Select Excel 64bit ODBC driver for which you wish to set up a data source
Step 4: Select workbook
Step 5. Ok then finish


PS C:\>

Add-OdbcDsn -Name "sql1" -DriverName "Microsoft Excel Driver(*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue "Dbq=d:\xls\s8fin.xlsx"
Add-OdbcDsn -Name "sql1" -DriverName "Microsoft Excel Driver(*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue @("KeyFilePath=d:\xls\s8fin.xlsx")
Set-OdbcDsn -Name "excel" -DriverName "Microsoft Excel Driver(*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -Platform "64-bit"


Add-OdbcDsn -Name "mdb64" -DriverName "Microsoft Access Driver (*.mdb, *.accdb)" -DsnType "User" -Platform "64-bit" -SetPropertyValue 'Dbq=d:\mdb\demo.mdb'
















*WORKS;
%utl_pybegin;
parmcards4;
import pyodbc
# MS Access DB connection
pyodbc.lowercase = False
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};' +
    r'DBQ=d:\mdb\demo.MDB;')

# Open cursor and execute SQL
cursor = conn.cursor()
cursor.execute('select name FROM cls');

for row in cursor.fetchall():
    print (row)
;;;;
%utl_pyend;

*works;
%utl_pybegin;
parmcards4;
import pyodbc
spreadsheet_path = "d:\\xls\\s8fin.xlsx"
cnxn = (r'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};'
            r'DBQ={}; ReadOnly=0').format(spreadsheet_path);
print("Opened Excel successfully")
cnxn = pyodbc.connect(cnxn, autocommit=True)
cursor= cnxn.cursor()
cursor.execute("Select isnumeric(a),isnumeric(b) from [sheet1$]")
row = cursor.fetchone()
print(row)
;;;;
%utl_pyend;


*works;
%utl_pybegin;
parmcards4;
import pandas
import pyodbc
spreadsheet_path = "d:\\xls\\s8fin.xlsx"
cnxn = (r'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};'
            r'DBQ={}; ReadOnly=0').format(spreadsheet_path);
print("Opened Excel successfully")
cnxn = pyodbc.connect(cnxn, autocommit=True)
sql = 'select count(*) as cnt, sum(isnumeric(a)) as cnta from [sheet1$]'
data = pandas.read_sql(sql, cnxn)
print(data);
;;;;
%utl_pyend;

         ,IIF( ( isnumeric(a)=0,'chr','num') as type
        ,count(*) + sum(isnumeric(a))   as NumChrName

*works;
%utl_pybegin;
parmcards4;
import pandas
import pyodbc
spreadsheet_path = "d:\\xls\\s8fin.xlsx"
cnxn = (r'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};'
            r'DBQ={}; ReadOnly=0').format(spreadsheet_path);
print("Opened Excel successfully")
cnxn = pyodbc.connect(cnxn, autocommit=True)
sql = 'select a, 2*b, c from [sheet1$]'
data = pandas.read_sql(sql, cnxn)
print(data);
;;;;
%utl_pyend;


            ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;
















import pyodbc



cnn = pyodbc.connect(...)

data = pandas.read_sql(sql, cnn)


cursor= cnxn.cursor()
cursor.execute("Select isnumeric(a),isnumeric(b) from [sheet1$]")
row = cursor.fetchone()
print(row)















%utl_submit_wps64x('
proc sql dquote=ansi;
   connect to excel (Path="d:\xls\have.xlsx");
     select
        *
     from
        connection to Excel
         (
          Select
               sex
              ,age
          from
               have
         );
       disconnect from Excel;
quit;
');


%let qq=%quotProvider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=d:\xls\have.xlsx;Mode=Share Deny Write;Jet OLEDB:Engine Type=37;

%utl_submit_wps64x(
proc sql dquote=ansi;
   connect to oledb as excel
      (Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=d:\xls\have.xlsx;Mode=Share Deny Write;Jet OLEDB:Engine Type=37;);
     select
        *
     from
        connection to excel
         (
          Select
               sex
              ,age
          from
               have
         );
       disconnect from Excel;
quit;
'));
         ;;;;%end;%mend;/*'*/ *);*};*];*/;/*"*/;run;quit;%end;end;run;endcomp;%utlfix;

%let str="Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=d:\xls\have.xlsx;Mode=Share Deny Write;Jet OLEDB:Engine Type=37;"

%utl_submit_ps64(%tslit('
$cnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=d:\xls\have.xlsx;Mode=Share Deny Write;Jet OLEDB:Engine Type=37;";
$cnStr
$cn = New-Object System.Data.OleDb.OleDbConnection $cnStr;
$cn.Open()

$cmd = $cn.CreateCommand()

$cmd.CommandText = "SELECT * FROM [Sheet1$]"
$rdr = $cmd.ExecuteReader();

$dt = new-object System.Data.DataTable

$dt.Load($rdr)

$dt | Out-GridView
'));

$cnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=d:\xls\s8fin.xlsx;Mode=Share Deny Write;Jet OLEDB:Engine Type=37;";
$cnStr
$cn = New-Object System.Data.OleDb.OleDbConnection $cnStr;
$cn.Open()

$cmd = $cn.CreateCommand()

$cmd.CommandText = "SELECT * FROM [sheet1$]"
$rdr = $cmd.ExecuteReader();

$dt = new-object System.Data.DataTable

$dt.Load($rdr)

$dt | Out-GridView


https://community.esri.com/t5/python-questions/ms-access-tables-to-excel/td-p/332167

%utl_pybegin;
parmcards4;
import pyodbc
# MS Access DB connection
pyodbc.lowercase = False
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};' +
    r'DBQ=d:\mdb\demo.MDB;')

# Open cursor and execute SQL
cursor = conn.cursor()
cursor.execute('select name FROM cls');

for row in cursor.fetchall():
    print (row)
;;;;
%utl_pyend;


con = ("Driver={Microsoft Excel Driver (*.xls)};DBQ="+ExcelPath+";ReadOnly = True;")




                d:\xls\s8fin.xlsx








cursor = conn.cursor()
cursor.execute('select name FROM [have$]');
;;;;
%utl_pyend;

https://www.red-gate.com/simple-talk/databases/sql-server/database-administration-sql-server/getting-data-between-excel-and-sql-server-using-odbc/


for row in cursor.fetchall():
    print (row)



import pyodbc

# Setup path and driver connection string
spreadsheet_path = "C:\\temp\\test_spreadsheet.xlsx"
conn_str = (r'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};'
            r'DBQ={}; ReadOnly=0').format(spreadsheet_path)

with pyodbc.connect(conn_str, autocommit=True) as conn:
    # Create table
    cursor = conn.cursor()
    query = "create table sheet1 (COL1 TEXT, COL2 NUMBER);"
    cursor.execute(query)
    cursor.commit()

    # Insert a row
    query = "insert into sheet1 (COL1, COL2) values (?, ?);"
    cursor.execute(query, "apples", 10)
    cursor.commit()

    # Check the row is there
    query = "select * from sheet1;"
    cursor.execute(query)
    for r in cursor.fetchall():
        print(r)

print("done")



























%utl_pybegin;
parmcards4;
import pyodbc
# MS Access DB connection
pyodbc.lowercase = False
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};' +
    r'DBQ=d:\mdb\demo.MDB;')

# Open cursor and execute SQL
cursor = conn.cursor()
cursor.execute('select name FROM cls');

for row in cursor.fetchall():
    print (row)
;;;;
%utl_pyend;

























https://search.r-project.org/CRAN/refmans/RODBC/html/sqlQuery.html
https://stackoverflow.com/questions/47008836/use-isnumeric-sql-query-via-vba-excel
https://stackoverflow.com/questions/20245647/rodbc-and-sqlqueries-with-r-objects?rq=4
https://learn.microsoft.com/en-us/sql/connect/python/pyodbc/step-3-proof-of-concept-connecting-to-sql-using-pyodbc?view=sql-server-ver16

*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/
*/



















    values('',   'HRR_BENE            ', )
            Alfred      M      14
            Alice       F      13
            Barbara     F      13
            Carol       F      14
            Henry       M      14
            James       M      12
            Jane        F      12








proc datasets lib=work nolist nodetails mt=cat;
  delete sasmac1 sasmac2 sasmac3 sasmac4;
run;quit;

proc datasets lib=sd1 nolist nodetails;delete want; run;quit;

options validvarname=any;
libname sd1 "d:/sd1";

%utl_submit_wps64x('
libname sd1 "d:/sd1";
proc r;
export data=sd1.have r=have;
submit;
have<-structure(list(buffer = c(100L, 200L, 300L, 400L, 500L, 600L,
700L, 800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L), date = structure(c(15984, 15984, 15984,
15984, 15984, 15984, 15984, 15984, 15984, 15984, 15984, 15984,
15984, 15984, 15984, 15984, 15984, 15984, 15984, 16000, 16000,
16000, 16000, 16000, 16000, 16000, 16000, 16000, 16000, 16000,
16000, 16000, 16000, 16000, 16000, 16000, 16000, 16000, 16064,
16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064,
16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064,
16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080,
16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080,
16080, 16112, 16112, 16112, 16112, 16112, 16112, 16112, 16112,
16112, 16112, 16112, 16112, 16112, 16112, 16112, 16112, 16112,
16112, 16112, 16144, 16144, 16144, 16144, 16144, 16144, 16144,
16144, 16144, 16144, 16144, 16144, 16144, 16144, 16144, 16144,
16144, 16144, 16144, 16176, 16176, 16176, 16176, 16176, 16176,
16176, 16176, 16176, 16176, 16176, 16176, 16176, 16176, 16176,
16176, 16176, 16176, 16176, 16192, 16192, 16192, 16192, 16192,
16192, 16192, 16192, 16192, 16192, 16192, 16192, 16192, 16192,
16192, 16192, 16192, 16192, 16192, 16240, 16240, 16240, 16240,
16240, 16240, 16240, 16240, 16240, 16240, 16240, 16240, 16240,
16240, 16240, 16240, 16240, 16240, 16240, 16272, 16272, 16272,
16272, 16272, 16272, 16272, 16272, 16272, 16272, 16272, 16272,
16272, 16272, 16272, 16272, 16272, 16272, 16272, 16304, 16304,
16304, 16304, 16304, 16304, 16304, 16304, 16304, 16304, 16304,
16304, 16304, 16304, 16304, 16304, 16304, 16304, 16304, 16336,
16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336,
16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336,
17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120,
17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120,
17120, 17152, 17152, 17152, 17152, 17152, 17152, 17152, 17152,
17152, 17152, 17152, 17152, 17152, 17152, 17152, 17152, 17152,
17152, 17152, 17184, 17184, 17184, 17184, 17184, 17184, 17184,
17184, 17184, 17184, 17184, 17184, 17184, 17184, 17184, 17184,
17184, 17184, 17184, 17223, 17223, 17223, 17223, 17223, 17223,
17223, 17223, 17223, 17223, 17223, 17223, 17223, 17223, 17223,
17223, 17223, 17223, 17223, 17232, 17232, 17232, 17232, 17232,
17232, 17232, 17232, 17232, 17232, 17232, 17232, 17232, 17232,
17232, 17232, 17232, 17232, 17232, 17264, 17264, 17264, 17264,
17264, 17264, 17264, 17264, 17264, 17264, 17264, 17264, 17264,
17264, 17264, 17264, 17264, 17264, 17264, 17312, 17312, 17312,
17312, 17312, 17312, 17312, 17312, 17312, 17312, 17312, 17312,
17312, 17312, 17312, 17312, 17312, 17312, 17312, 17328, 17328,
17328, 17328, 17328, 17328, 17328, 17328, 17328, 17328, 17328,
17328, 17328, 17328, 17328, 17328, 17328, 17328, 17328, 17360,
17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360,
17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360,
17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392,
17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392,
17392, 17424, 17424, 17424, 17424, 17424, 17424, 17424, 17424,
17424, 17424, 17424, 17424, 17424, 17424, 17424, 17424, 17424,
17424, 17424, 17456, 17456, 17456, 17456, 17456, 17456, 17456,
17456, 17456, 17456, 17456, 17456, 17456, 17456, 17456, 17456,
17456, 17456, 17456), class = "Date"), mean = c(1.215838882,
1.088822109, 0.989531359, 0.823662403, 0.711081399, 0.538787295,
0.433376764, 0.264958626, 0.25004379, 0.470761656, 0.388691884,
0.098983408, 0.2033786, 0.186319853, 0.316626707, 0.384678207,
0.151754468, 0.021208597, 0.070374404, 1.642164596, 1.5240609,
1.356737414, 1.138406446, 0.989445483, 0.775460764, 0.64431253,
0.457905314, 0.403775943, 0.520688199, 0.370364428, 0.147411867,
0.177712147, 0.175085484, 0.237685702, 0.318550327, 0.165086768,
0.047179424, 0.083480265, 0.959871169, 0.861126848, 0.715940454,
0.55645199, 0.42624073, 0.274809049, 0.1777997, 0.084744839,
0.054235277, 0.143284816, 0.094401748, 0.013418036, 0.02932515,
0.016985843, 0.033923091, 0.109454413, 0.079167565, 0.011850503,
0.015270404, 1.115561879, 1.118547733, 1.008445249, 0.807292202,
0.619912547, 0.279178175, 0.139024045, 0.012907406, -0.009447278,
0.049632182, 0.00678033, -0.092529329, -0.012337128, 0.033859182,
0.081019116, 0.107821823, 0.008842851, -0.011484568, 0.037363486,
1.186130896, 1.021365666, 0.857613199, 0.643899541, 0.477285226,
0.308509163, 0.255213401, 0.117960822, 0.124952278, 0.293408528,
0.217805546, -0.025199489, 0.004961657, 0.02917934, 0.129920121,
0.236894044, 0.079171313, -0.043470064, 0.006356774, 1.64361703,
1.54039779, 1.411086458, 1.236237665, 1.066643799, 0.766247715,
0.633964581, 0.468390994, 0.450625728, 0.599879161, 0.470545048,
0.231224345, 0.266935106, 0.243939704, 0.361542188, 0.472305981,
0.21139414, 0.062885848, 0.070500871, 2.303472759, 2.157337917,
1.969416805, 1.742266165, 1.536159994, 1.251516299, 1.066625294,
0.775546646, 0.66155575, 0.808701489, 0.656754134, 0.332386128,
0.410712469, 0.394033103, 0.513500572, 0.600072554, 0.335750906,
0.113518036, 0.129407806, 2.943621788, 2.751945666, 2.458050683,
2.188826179, 1.956691499, 1.692109995, 1.468835698, 1.168147981,
1.041755112, 1.226755177, 0.981195245, 0.576804209, 0.685852767,
0.660669348, 0.775221712, 0.828703129, 0.463769664, 0.15422009,
0.162982951, 2.554575623, 2.372103254, 2.290408464, 2.054517841,
1.861615329, 1.592217004, 1.401797377, 1.051598412, 1.002812695,
1.267668649, 1.115415035, 0.603692868, 0.710852895, 0.705805591,
0.875490839, 0.916332795, 0.473202815, 0.164548251, 0.133674457,
2.496432459, 2.405860455, 2.332276626, 2.154314468, 1.931537295,
1.593605353, 1.35674335, 1.064953145, 0.948308816, 1.117156309,
0.902296258, 0.499734413, 0.568312216, 0.536336683, 0.646246448,
0.703519511, 0.360827234, 0.116159248, 0.121843156, -0.193683744,
-0.150883103, -0.163360656, -0.22786412, -0.257963676, -0.307309721,
-0.305761362, -0.394724905, -0.407210918, -0.177085124, -0.140428942,
-0.201550107, -0.019512414, 0.036890416, -0.036720632, -0.010561275,
-0.016140439, 0.004593684, 0.066561324, 1.380684443, 1.250706119,
1.173965001, 0.914397036, 0.760097916, 0.509322326, 0.391089492,
0.164465793, 0.148020166, 0.313878606, 0.154586799, -0.11587368,
-0.031067692, 0.00211778, 0.203495391, 0.340638935, 0.127880642,
-0.025747548, 0.033259082, 0.056043885, 0.271880696, 0.292647284,
0.184580925, 0.104690094, -0.111611973, -0.1701464, -0.265878888,
-0.275633824, -0.155178802, -0.204414597, -0.319376869, -0.271366792,
-0.237639718, -0.146844029, -0.035182789, -0.059645122, -0.047620826,
0.017376059, 1.392743445, 1.417764028, 1.415585906, 1.327725057,
1.261706787, 1.088591395, 1.012315648, 0.903954473, 0.849904531,
0.937846514, 0.814444927, 0.642600376, 0.619143953, 0.545698885,
0.51165879, 0.464725757, 0.290636394, 0.171534375, 0.085728616,
-0.106535134, -0.00742211, 0.011385109, -0.069584823, -0.11296655,
-0.249723692, -0.273480409, -0.276303109, -0.226114738, -0.167877177,
-0.110217493, -0.089293596, -0.101253827, -0.094656478, -0.07174728,
0.030344275, 0.045722273, 0.013324444, -0.003683338, 2.747240329,
2.836421225, 2.705051751, 2.526198267, 2.204361162, 1.488168719,
1.180385927, 0.952864236, 0.803868624, 0.643424534, 0.503747022,
0.469753831, 0.447443175, 0.336910431, 0.287241996, 0.254510912,
0.229997111, 0.136083221, 0.062082349, 2.116996183, 2.310411315,
2.282293682, 2.074499782, 1.915750444, 1.677553698, 1.506378314,
1.229625557, 1.146444017, 1.226392345, 1.046117019, 0.699447822,
0.619510061, 0.522634969, 0.61499077, 0.629987256, 0.309672406,
0.132654223, 0.118852919, 1.282711399, 1.744908103, 1.398968321,
1.034402073, 0.701381408, 0.118671792, -0.200273726, -0.385770537,
-0.507956839, -0.517404552, -0.52280134, -0.49241506, -0.383178089,
-0.345206625, -0.313454904, -0.235689519, -0.14605215, -0.116389156,
-0.071567936, 2.905459757, 2.976478462, 2.930253766, 2.752992955,
2.568241121, 2.291361426, 2.070697474, 1.698827001, 1.652046296,
1.8812795, 1.658990564, 1.150503479, 1.193308549, 1.088186829,
1.276389399, 1.271844, 0.690707389, 0.320918302, 0.263052932,
0.163346787, 0.323565285, 0.355617739, 0.299568497, 0.339650405,
0.30381453, 0.213207318, -0.015656403, 0.016348774, 0.309246444,
0.252628522, -0.04072425, 0.119878134, 0.202926357, 0.448243018,
0.651444884, 0.339529504, 0.053395622, 0.059782062, 3.109011659,
3.293034789, 3.148142518, 2.847359107, 2.620571798, 2.354282192,
2.154662371, 1.799961345, 1.659077036, 1.779987021, 1.523015934,
1.007879685, 1.020642289, 0.933913797, 1.050122458, 1.112308875,
0.708600292, 0.328096844, 0.238061694, 1.471810915, 1.645113526,
1.651099529, 1.477998942, 1.340008181, 1.084850388, 0.909319795,
0.662641391, 0.600637409, 0.857133307, 0.612528334, 0.244987389,
0.365849339, 0.35461597, 0.548239393, 0.652339938, 0.261362181,
0.08712113, 0.091655186, 3.324442272, 3.368835357, 3.226617548,
2.959897028, 2.754235554, 2.521647178, 2.327990844, 1.997280721,
1.940654329, 2.127070455, 1.865786378, 1.305294559, 1.211158642,
1.05880232, 1.178376533, 1.198792776, 0.708653958, 0.355491786,
0.257578385, -0.143639278, 0.558572512, 0.007050647, -0.268860295,
-0.703218426, -1.308560058, -1.085585002, -0.510504714, -0.397561478,
-0.00358799, 0.085958807, -0.237357804, -0.162937684, -0.127590858,
0.206757209, 0.432435649, 0.410980363, 0.363155446, 0.092740666
), Month = c(10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L,
10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 11L, 11L, 11L, 11L,
11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L,
11L, 11L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L,
12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L,
3L, 3L, 3L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L,
4L, 4L, 4L, 4L, 4L, 4L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L,
5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 6L, 6L, 6L, 6L, 6L, 6L, 6L,
6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 7L, 7L, 7L, 7L,
7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 8L,
8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L,
8L, 8L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L,
9L, 9L, 9L, 9L, 9L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L,
11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 12L, 12L, 12L,
12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L,
12L, 12L, 12L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 3L, 3L, 3L, 3L, 3L, 3L,
3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 4L, 4L, 4L,
4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L,
5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L,
5L, 5L, 5L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L,
6L, 6L, 6L, 6L, 6L, 6L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L,
7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 8L, 8L, 8L, 8L, 8L, 8L, 8L,
8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 9L, 9L, 9L, 9L,
9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 10L,
10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L,
10L, 10L, 10L, 10L, 10L), Year = c(2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L), JulianDay = c(279L, 279L, 279L, 279L, 279L, 279L,
279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L,
279L, 279L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L,
295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 359L,
359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L,
359L, 359L, 359L, 359L, 359L, 359L, 359L, 10L, 10L, 10L, 10L,
10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L,
10L, 10L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L,
42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 74L, 74L, 74L, 74L, 74L,
74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L,
74L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L,
106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 122L, 122L,
122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L,
122L, 122L, 122L, 122L, 122L, 122L, 170L, 170L, 170L, 170L, 170L,
170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L,
170L, 170L, 170L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L,
202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L,
234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L,
234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 266L, 266L, 266L,
266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L,
266L, 266L, 266L, 266L, 266L, 320L, 320L, 320L, 320L, 320L, 320L,
320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L,
320L, 320L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L,
352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 18L,
18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L,
18L, 18L, 18L, 18L, 18L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L,
57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 66L, 66L,
66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L,
66L, 66L, 66L, 66L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L,
98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 146L, 146L,
146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L,
146L, 146L, 146L, 146L, 146L, 146L, 162L, 162L, 162L, 162L, 162L,
162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L,
162L, 162L, 162L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L,
194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L,
226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L,
226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 258L, 258L, 258L,
258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L,
258L, 258L, 258L, 258L, 258L, 290L, 290L, 290L, 290L, 290L, 290L,
290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L,
290L, 290L), TimePeriod = structure(c(1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L), levels = c("1", "2"), class = "factor")), row.names = c(NA,
-456L), class = "data.frame");
endsubmit;
import data=sd1.have r=have;
run;quit;
');

I want to calculate an average for each buffer across all 12 months.

%utl_submit_wps64x('
libame xls "d:\xls\tempChange.xlsx";
data xls.hav;
  set sd1.have(keep=buffer mean rename=mean=tempChange);
run;quit;
');

proc print data=sd1.want width=min;
run;quit;

proc sql;
  create
    table bufSvg as
  select
    buffer
   ,mean(buffer) as avgBuf
  from
    hav
  group
   by buffer
;quit;


https://github.com/rogerjdeangelis/utl-PASSTHRU-to-mysql-and-select-rows-based-on-a-SAS-dataset-without-loading-the-SAS-daatset-into-my
https://github.com/rogerjdeangelis/utl-excel-fixing-bad-formatting-using-passthru
https://github.com/rogerjdeangelis/utl-fix-excel-columns-with-mutiple-datatypes-on-the-excel-side-using-ms-sql-and-passthru
https://github.com/rogerjdeangelis/utl-fix-excel-date-fields-on-the-excel-side-using-ms-sql-and-passthru
https://github.com/rogerjdeangelis/utl_passthru_to_excel_to_fix_column_names
https://github.com/rogerjdeangelis/utl_using_a_macro_variable_in_a_passthru_where_clause_to_a_foreign_database




data class;
  set sashelp.class;

proc sql dquote=ansi;
   connect to excel (Path="d:\xls\dates.xlsx" mixed=yes);
     create
         table dates as
     select
        dates
       ,input(dteChr,mmddyy10.)  as SAS_dates
       ,put(calculated sas_dates,mmddyy10.) as text_dates
     from
        connection to Excel
         (
          Select
               dates
               ,iif(isnumeric(dates),format(dates,"mm/dd/yy"),cvdate(dates) ) as dteChr
          from
               dates
         );
       disconnect from Excel;
quit;






























structure(list(buffer = c(100L, 200L, 300L, 400L, 500L, 600L,
700L, 800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L, 100L, 200L, 300L, 400L, 500L, 600L, 700L,
800L, 900L, 1000L, 1100L, 1200L, 1300L, 1400L, 1500L, 1600L,
1700L, 1800L, 1900L), date = structure(c(15984, 15984, 15984,
15984, 15984, 15984, 15984, 15984, 15984, 15984, 15984, 15984,
15984, 15984, 15984, 15984, 15984, 15984, 15984, 16000, 16000,
16000, 16000, 16000, 16000, 16000, 16000, 16000, 16000, 16000,
16000, 16000, 16000, 16000, 16000, 16000, 16000, 16000, 16064,
16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064,
16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064, 16064,
16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080,
16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080, 16080,
16080, 16112, 16112, 16112, 16112, 16112, 16112, 16112, 16112,
16112, 16112, 16112, 16112, 16112, 16112, 16112, 16112, 16112,
16112, 16112, 16144, 16144, 16144, 16144, 16144, 16144, 16144,
16144, 16144, 16144, 16144, 16144, 16144, 16144, 16144, 16144,
16144, 16144, 16144, 16176, 16176, 16176, 16176, 16176, 16176,
16176, 16176, 16176, 16176, 16176, 16176, 16176, 16176, 16176,
16176, 16176, 16176, 16176, 16192, 16192, 16192, 16192, 16192,
16192, 16192, 16192, 16192, 16192, 16192, 16192, 16192, 16192,
16192, 16192, 16192, 16192, 16192, 16240, 16240, 16240, 16240,
16240, 16240, 16240, 16240, 16240, 16240, 16240, 16240, 16240,
16240, 16240, 16240, 16240, 16240, 16240, 16272, 16272, 16272,
16272, 16272, 16272, 16272, 16272, 16272, 16272, 16272, 16272,
16272, 16272, 16272, 16272, 16272, 16272, 16272, 16304, 16304,
16304, 16304, 16304, 16304, 16304, 16304, 16304, 16304, 16304,
16304, 16304, 16304, 16304, 16304, 16304, 16304, 16304, 16336,
16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336,
16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336, 16336,
17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120,
17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120, 17120,
17120, 17152, 17152, 17152, 17152, 17152, 17152, 17152, 17152,
17152, 17152, 17152, 17152, 17152, 17152, 17152, 17152, 17152,
17152, 17152, 17184, 17184, 17184, 17184, 17184, 17184, 17184,
17184, 17184, 17184, 17184, 17184, 17184, 17184, 17184, 17184,
17184, 17184, 17184, 17223, 17223, 17223, 17223, 17223, 17223,
17223, 17223, 17223, 17223, 17223, 17223, 17223, 17223, 17223,
17223, 17223, 17223, 17223, 17232, 17232, 17232, 17232, 17232,
17232, 17232, 17232, 17232, 17232, 17232, 17232, 17232, 17232,
17232, 17232, 17232, 17232, 17232, 17264, 17264, 17264, 17264,
17264, 17264, 17264, 17264, 17264, 17264, 17264, 17264, 17264,
17264, 17264, 17264, 17264, 17264, 17264, 17312, 17312, 17312,
17312, 17312, 17312, 17312, 17312, 17312, 17312, 17312, 17312,
17312, 17312, 17312, 17312, 17312, 17312, 17312, 17328, 17328,
17328, 17328, 17328, 17328, 17328, 17328, 17328, 17328, 17328,
17328, 17328, 17328, 17328, 17328, 17328, 17328, 17328, 17360,
17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360,
17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360, 17360,
17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392,
17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392, 17392,
17392, 17424, 17424, 17424, 17424, 17424, 17424, 17424, 17424,
17424, 17424, 17424, 17424, 17424, 17424, 17424, 17424, 17424,
17424, 17424, 17456, 17456, 17456, 17456, 17456, 17456, 17456,
17456, 17456, 17456, 17456, 17456, 17456, 17456, 17456, 17456,
17456, 17456, 17456), class = "Date"), mean = c(1.215838882,
1.088822109, 0.989531359, 0.823662403, 0.711081399, 0.538787295,
0.433376764, 0.264958626, 0.25004379, 0.470761656, 0.388691884,
0.098983408, 0.2033786, 0.186319853, 0.316626707, 0.384678207,
0.151754468, 0.021208597, 0.070374404, 1.642164596, 1.5240609,
1.356737414, 1.138406446, 0.989445483, 0.775460764, 0.64431253,
0.457905314, 0.403775943, 0.520688199, 0.370364428, 0.147411867,
0.177712147, 0.175085484, 0.237685702, 0.318550327, 0.165086768,
0.047179424, 0.083480265, 0.959871169, 0.861126848, 0.715940454,
0.55645199, 0.42624073, 0.274809049, 0.1777997, 0.084744839,
0.054235277, 0.143284816, 0.094401748, 0.013418036, 0.02932515,
0.016985843, 0.033923091, 0.109454413, 0.079167565, 0.011850503,
0.015270404, 1.115561879, 1.118547733, 1.008445249, 0.807292202,
0.619912547, 0.279178175, 0.139024045, 0.012907406, -0.009447278,
0.049632182, 0.00678033, -0.092529329, -0.012337128, 0.033859182,
0.081019116, 0.107821823, 0.008842851, -0.011484568, 0.037363486,
1.186130896, 1.021365666, 0.857613199, 0.643899541, 0.477285226,
0.308509163, 0.255213401, 0.117960822, 0.124952278, 0.293408528,
0.217805546, -0.025199489, 0.004961657, 0.02917934, 0.129920121,
0.236894044, 0.079171313, -0.043470064, 0.006356774, 1.64361703,
1.54039779, 1.411086458, 1.236237665, 1.066643799, 0.766247715,
0.633964581, 0.468390994, 0.450625728, 0.599879161, 0.470545048,
0.231224345, 0.266935106, 0.243939704, 0.361542188, 0.472305981,
0.21139414, 0.062885848, 0.070500871, 2.303472759, 2.157337917,
1.969416805, 1.742266165, 1.536159994, 1.251516299, 1.066625294,
0.775546646, 0.66155575, 0.808701489, 0.656754134, 0.332386128,
0.410712469, 0.394033103, 0.513500572, 0.600072554, 0.335750906,
0.113518036, 0.129407806, 2.943621788, 2.751945666, 2.458050683,
2.188826179, 1.956691499, 1.692109995, 1.468835698, 1.168147981,
1.041755112, 1.226755177, 0.981195245, 0.576804209, 0.685852767,
0.660669348, 0.775221712, 0.828703129, 0.463769664, 0.15422009,
0.162982951, 2.554575623, 2.372103254, 2.290408464, 2.054517841,
1.861615329, 1.592217004, 1.401797377, 1.051598412, 1.002812695,
1.267668649, 1.115415035, 0.603692868, 0.710852895, 0.705805591,
0.875490839, 0.916332795, 0.473202815, 0.164548251, 0.133674457,
2.496432459, 2.405860455, 2.332276626, 2.154314468, 1.931537295,
1.593605353, 1.35674335, 1.064953145, 0.948308816, 1.117156309,
0.902296258, 0.499734413, 0.568312216, 0.536336683, 0.646246448,
0.703519511, 0.360827234, 0.116159248, 0.121843156, -0.193683744,
-0.150883103, -0.163360656, -0.22786412, -0.257963676, -0.307309721,
-0.305761362, -0.394724905, -0.407210918, -0.177085124, -0.140428942,
-0.201550107, -0.019512414, 0.036890416, -0.036720632, -0.010561275,
-0.016140439, 0.004593684, 0.066561324, 1.380684443, 1.250706119,
1.173965001, 0.914397036, 0.760097916, 0.509322326, 0.391089492,
0.164465793, 0.148020166, 0.313878606, 0.154586799, -0.11587368,
-0.031067692, 0.00211778, 0.203495391, 0.340638935, 0.127880642,
-0.025747548, 0.033259082, 0.056043885, 0.271880696, 0.292647284,
0.184580925, 0.104690094, -0.111611973, -0.1701464, -0.265878888,
-0.275633824, -0.155178802, -0.204414597, -0.319376869, -0.271366792,
-0.237639718, -0.146844029, -0.035182789, -0.059645122, -0.047620826,
0.017376059, 1.392743445, 1.417764028, 1.415585906, 1.327725057,
1.261706787, 1.088591395, 1.012315648, 0.903954473, 0.849904531,
0.937846514, 0.814444927, 0.642600376, 0.619143953, 0.545698885,
0.51165879, 0.464725757, 0.290636394, 0.171534375, 0.085728616,
-0.106535134, -0.00742211, 0.011385109, -0.069584823, -0.11296655,
-0.249723692, -0.273480409, -0.276303109, -0.226114738, -0.167877177,
-0.110217493, -0.089293596, -0.101253827, -0.094656478, -0.07174728,
0.030344275, 0.045722273, 0.013324444, -0.003683338, 2.747240329,
2.836421225, 2.705051751, 2.526198267, 2.204361162, 1.488168719,
1.180385927, 0.952864236, 0.803868624, 0.643424534, 0.503747022,
0.469753831, 0.447443175, 0.336910431, 0.287241996, 0.254510912,
0.229997111, 0.136083221, 0.062082349, 2.116996183, 2.310411315,
2.282293682, 2.074499782, 1.915750444, 1.677553698, 1.506378314,
1.229625557, 1.146444017, 1.226392345, 1.046117019, 0.699447822,
0.619510061, 0.522634969, 0.61499077, 0.629987256, 0.309672406,
0.132654223, 0.118852919, 1.282711399, 1.744908103, 1.398968321,
1.034402073, 0.701381408, 0.118671792, -0.200273726, -0.385770537,
-0.507956839, -0.517404552, -0.52280134, -0.49241506, -0.383178089,
-0.345206625, -0.313454904, -0.235689519, -0.14605215, -0.116389156,
-0.071567936, 2.905459757, 2.976478462, 2.930253766, 2.752992955,
2.568241121, 2.291361426, 2.070697474, 1.698827001, 1.652046296,
1.8812795, 1.658990564, 1.150503479, 1.193308549, 1.088186829,
1.276389399, 1.271844, 0.690707389, 0.320918302, 0.263052932,
0.163346787, 0.323565285, 0.355617739, 0.299568497, 0.339650405,
0.30381453, 0.213207318, -0.015656403, 0.016348774, 0.309246444,
0.252628522, -0.04072425, 0.119878134, 0.202926357, 0.448243018,
0.651444884, 0.339529504, 0.053395622, 0.059782062, 3.109011659,
3.293034789, 3.148142518, 2.847359107, 2.620571798, 2.354282192,
2.154662371, 1.799961345, 1.659077036, 1.779987021, 1.523015934,
1.007879685, 1.020642289, 0.933913797, 1.050122458, 1.112308875,
0.708600292, 0.328096844, 0.238061694, 1.471810915, 1.645113526,
1.651099529, 1.477998942, 1.340008181, 1.084850388, 0.909319795,
0.662641391, 0.600637409, 0.857133307, 0.612528334, 0.244987389,
0.365849339, 0.35461597, 0.548239393, 0.652339938, 0.261362181,
0.08712113, 0.091655186, 3.324442272, 3.368835357, 3.226617548,
2.959897028, 2.754235554, 2.521647178, 2.327990844, 1.997280721,
1.940654329, 2.127070455, 1.865786378, 1.305294559, 1.211158642,
1.05880232, 1.178376533, 1.198792776, 0.708653958, 0.355491786,
0.257578385, -0.143639278, 0.558572512, 0.007050647, -0.268860295,
-0.703218426, -1.308560058, -1.085585002, -0.510504714, -0.397561478,
-0.00358799, 0.085958807, -0.237357804, -0.162937684, -0.127590858,
0.206757209, 0.432435649, 0.410980363, 0.363155446, 0.092740666
), Month = c(10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L,
10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 11L, 11L, 11L, 11L,
11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L,
11L, 11L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L,
12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L,
3L, 3L, 3L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L,
4L, 4L, 4L, 4L, 4L, 4L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L,
5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 6L, 6L, 6L, 6L, 6L, 6L, 6L,
6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 7L, 7L, 7L, 7L,
7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 8L,
8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L,
8L, 8L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L,
9L, 9L, 9L, 9L, 9L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L,
11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 11L, 12L, 12L, 12L,
12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L, 12L,
12L, 12L, 12L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 3L, 3L, 3L, 3L, 3L, 3L,
3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 3L, 4L, 4L, 4L,
4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L, 4L,
5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L, 5L,
5L, 5L, 5L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L, 6L,
6L, 6L, 6L, 6L, 6L, 6L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L,
7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 7L, 8L, 8L, 8L, 8L, 8L, 8L, 8L,
8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 8L, 9L, 9L, 9L, 9L,
9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 9L, 10L,
10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L,
10L, 10L, 10L, 10L, 10L), Year = c(2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L,
2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2013L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L,
2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2014L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L, 2016L,
2016L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L, 2017L,
2017L, 2017L), JulianDay = c(279L, 279L, 279L, 279L, 279L, 279L,
279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L, 279L,
279L, 279L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L,
295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 295L, 359L,
359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L, 359L,
359L, 359L, 359L, 359L, 359L, 359L, 359L, 10L, 10L, 10L, 10L,
10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L, 10L,
10L, 10L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L,
42L, 42L, 42L, 42L, 42L, 42L, 42L, 42L, 74L, 74L, 74L, 74L, 74L,
74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L, 74L,
74L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L,
106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 106L, 122L, 122L,
122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L, 122L,
122L, 122L, 122L, 122L, 122L, 122L, 170L, 170L, 170L, 170L, 170L,
170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L, 170L,
170L, 170L, 170L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L,
202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L, 202L,
234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L,
234L, 234L, 234L, 234L, 234L, 234L, 234L, 234L, 266L, 266L, 266L,
266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L, 266L,
266L, 266L, 266L, 266L, 266L, 320L, 320L, 320L, 320L, 320L, 320L,
320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L, 320L,
320L, 320L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L,
352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 352L, 18L,
18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L, 18L,
18L, 18L, 18L, 18L, 18L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L,
57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 57L, 66L, 66L,
66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L, 66L,
66L, 66L, 66L, 66L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L,
98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 98L, 146L, 146L,
146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L, 146L,
146L, 146L, 146L, 146L, 146L, 146L, 162L, 162L, 162L, 162L, 162L,
162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L, 162L,
162L, 162L, 162L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L,
194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L, 194L,
226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L,
226L, 226L, 226L, 226L, 226L, 226L, 226L, 226L, 258L, 258L, 258L,
258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L, 258L,
258L, 258L, 258L, 258L, 258L, 290L, 290L, 290L, 290L, 290L, 290L,
290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L, 290L,
290L, 290L), TimePeriod = structure(c(1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L,
1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 1L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L, 2L,
2L, 2L), levels = c("1", "2"), class = "factor")), row.names = c(NA,
-456L), class = "data.frame")
*
123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234
;
*/ /*************************************************************************************************************************
*/ /*
*/ /*  The WPS System
*/ /*
*/ /*    sex   avgAge
*/ /*  1   F 13.33333
*/ /*  2   M 13.33333
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*
*/ /*************************************************************************************************************************
