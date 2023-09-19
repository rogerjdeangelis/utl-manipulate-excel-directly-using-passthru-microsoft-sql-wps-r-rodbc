# utl-manipulate-excel-directly-using-passthru-microsoft-sql-wps-r-rodbc
Given two excel tables(named ranges) ,'males' and 'females' do the following  Calculate then mean age by sex using MS SQL inside an excel workbook         
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
