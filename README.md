# Visual-Basic-For-Application-Projects
This repository contains all projects, coding standards, boilerplate, and best practices. This is intended to share my experience and expertise so everyone can fork or download everything in this repository even without permission.

#### `VBA CODING STANDARDS AND BEST PRACTICES`
- Enable `Option Explicit`
  - This Forces explicit declaration of all variables in a file, or allows implicit declarations of variables.
- Use `Title blocks for each Macro`
  - This title blocks will give you information about the module what it does and how it will be called.
```
        PROGRAM:         UpdateCustomerCur
        DESCRIPTION:     This is to update the list of customer currency in the database
        CALL:            Call thru Macro UpdateCustomerCur
```
- Get Advantage of `Modular Programming` Technique
  - Modular programming is a software design technique that emphasizes seperating the functionality of a program into dependent, interchangeable modules, such that each contains everything necessary to execute only one aspect of the desired functionality.
  - Categorize and divide your code based from what it does and create one module for similar task. Follow Modular Programming technique due to its advantage when refactoring and scaling your code. Also, it will give you less headache when you're debugging your code, easy to understand, and really helpful for big and complex projects.

- Declaring `Constants and Variables`
  - The scope of a variable in Excel VBA determines where that variable may be used. You determine the `scope of a variable` when you declare it. There are three scoping levels: procedure level, module level, and public module level.
  - Make use the advantage of declaring variables based from scope (i.e. Local, Module, Public)
  - `Constants` are coded in ALL_UPPER_CASE with words seperated by underscores.
```
    Global Const WDGT_STATUS_OK = 0
    Global Const WDGT_STATUS_BUSY = 1
    Global Const WDGT_STATUS_FAIL = 2
    Global Const WDGT_STATUS_OFF = 3
    Global Const WDGT_STATUS_START = 4
    Global Const WDGT_STATUS_STOP = 5
```
- Naming Convention for SubProcedures and Functions >>> `VERB.NOUN.ADJECTIVE`
```
    Sub UpdateCustomerCur()
      'Code goes here
    End Sub
```
- Use `HungarianNotation` for Variables, Constants, SubProcedures and Functions
  - Hungarian notation is an identifier naming convention in computer programming, in which the name of a variable or function indicates its intention or kind, and in some dialects its type. The original Hungarian Notation uses intention or kind in its naming convention and is sometimes called Apps Hungarian as it became popular in the Microsoft Apps division in the development of Word, Excel and other apps. As the Microsoft Windows division adopted the naming convention, they used the actual data type for naming, and this convention became widely spread through the Windows API; this is sometimes called Systems Hungarian notation.

```
      Dim strMyName As String
      Dim intMyNumber As Integer
```

```
DECLARING VARIABLES
        VARIABLE       TAG             NOTES
        BOOLEAN        bln             blnFound
        BYTE           byt             bytRasterData
        CURRENCY       cur             curRevenue
        DATE (Time)    dat             datStart
        DOUBLE         dbl             dblTolerance
        ENUM           enm             enmColours
        INTEGERS       int             intQuantity
        LONG           lng             lngDistance
        SINGLE         sng             sngAverage
        STRING         str             strCustName
        USERDEFINED    udt             udtEmployee
        Variant        var             varCheckSum

OTHER PREFIXES
        cbo ComboBox
        chk CheckBox
        cmd CommandButton
        frm Form
        img Image
        lbl Lable
        lst ListBox
        rpt Report
        shp Shape
        txt TextBox
        tbl Table
        ole OLE Control
        pic Picture
        pnl Panel
        qry Query

        Db Database
        ws Workspace
        rs Recordset
        xl Excel Object
        wd Word Object
```
