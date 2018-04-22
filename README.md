# utl_excel_import_data_from_a_xlsx_file_where_first_2_rows_are_header
Import data from a xlsx file where first 2 rows are header. Keywords: sas sql join merge big data analytics macros oracle teradata mysql sas communities stackoverflow statistics artificial inteligence AI Python R Java Javascript WPS Matlab SPSS Scala Perl C C# Excel MS Access JSON graphics maps NLP natural language processing machine learning igraph DOSUBL DOW loop stackoverflow SAS community.
    Import data from a xlsx file where first 2 rows are header

    github
    https://tinyurl.com/ya59feug
    https://github.com/rogerjdeangelis/utl_excel_import_data_from_a_xlsx_file_where_first_2_rows_are_header

    see
    https://tinyurl.com/y7r5pm49
    https://communities.sas.com/t5/Base-SAS-Programming/I-am-trying-to-import-data-from-a-xlsx-file-where-first-2-rows/m-p/455820

      Three solutions (none of which use proc import)
      (only 1 and 3 have the correct names)

        1.  Failed in WPS  Data set "XEL.sheet1$a3:z99" not found
            libname xel "d:/xls/class.xlsx";
            set xel.'sheet1$a3:z99'n;

        2.  Same result SAS and WPS
            libname xel "d:/xls/class.xlsx" header=no;
            set xel.'sheet1$'n(firstobs=4);

        3.  Same result SAS and WPS
            libname xel "d:/xls/class.xlsx" header=no;
            data class(
                rename=(
                  f1=name
                  f2=sex
                  f3=age
                  f4=height
                  f5=weight
                ));
              set xel.'sheet1$'n(firstobs=4);


    INPUT
    =====

    SHEET CLASS IN WORKBOOK D:/XLS/CLASS.XLSX


      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
    1 | HEADER 1                                                       |
      +------------+------------+------------+------------+------------+
    2 | HEADER 2                                                       |
      +------------+------------+------------+------------+------------+
    3 |  NAME      |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
    4 | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
    N | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

    [SHEET1]

    Example table I want

    WORK.CLASS total obs=19

    Obs    NAME       SEX    AGE    HEIGHT    WEIGHT

      1    Alfred      M      14     69.0      112.5
      2    Alice       F      13     56.5       84.0
      3    Barbara     F      13     65.3       98.0
      4    Carol       F      14     62.8      102.5
      5    Henry       M      14     63.5      102.5


    PROCESS

    * read row 3 with names and then data;
    * note only cells in class ar read in even though z99 is used;

    libname xel "d:/xls/class.xlsx";
    data class;
      set xel.'sheet1$a3:z99'n;
    run;quit;
    libname xel clear;

    libname xel "d:/xls/class.xlsx" header=no;
    data class;
      set xel.'sheet1$'n(firstobs=4);
    run;quit;
    libname xel clear;

    libname xel "d:/xls/class.xlsx" header=no;
    data class(
        rename=(
          f1=name
          f2=sex
          f3=age
          f4=height
          f5=weight
        ));
      set xel.'sheet1$'n(firstobs=4);
    run;quit;
    libname xel clear;

    *                _              _       _
     _ __ ___   __ _| | _____    __| | __ _| |_ __ _
    | '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
    | | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
    |_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

    ;

    * create sheet1 with two header rows;

    ods excel file="d:/xls/class.xlsx"
       options(sheet_name="sheet1");
    * created the code below using the list option on proc report;

    PROC REPORT DATA=SASHELP.CLASS LS=171 PS=65  SPLIT="/" NOCENTER ;
    title1 "Header 1";
    title2 "Header 2";
    COLUMN  ("Header1" ("Header2"  NAME SEX AGE HEIGHT WEIGHT));

    DEFINE  NAME / DISPLAY FORMAT= $8. WIDTH=8     SPACING=2   LEFT "NAME" ;
    DEFINE  SEX / DISPLAY FORMAT= $1. WIDTH=1     SPACING=2   LEFT "SEX" ;
    DEFINE  AGE / SUM FORMAT= BEST9. WIDTH=9     SPACING=2   RIGHT "AGE" ;
    DEFINE  HEIGHT / SUM FORMAT= BEST9. WIDTH=9     SPACING=2   RIGHT "HEIGHT" ;
    DEFINE  WEIGHT / SUM FORMAT= BEST9. WIDTH=9     SPACING=2   RIGHT "WEIGHT" ;

    RUN;
    ods excel close;

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __  ___
    / __|/ _ \| | | | | __| |/ _ \| '_ \/ __|
    \__ \ (_) | | |_| | |_| | (_) | | | \__ \
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|___/

    ;

    *SAS;

    * read row 3 with names and then data;
    * note only cells in class ar read in even though z99 is used;

    libname xel "d:/xls/class.xlsx";
    data class;
      set xel.'sheet1$a3:z99'n;
    run;quit;
    libname xel clear;

    libname xel "d:/xls/class.xlsx" header=no;
    data class(
        rename=(
          f1=name
          f2=sex
          f3=age
          f4=height
          f5=weight
        ));
      set xel.'sheet1$'n(firstobs=4);
    run;quit;
    libname xel clear;



    libname xel "d:/xls/class.xlsx" header=no;
    data class;
      set xel."sheet1$"n(firstobs=4);
    run;quit;
    libname xel clear;

    *WPS;
    %utl_submit_wps64('
    libname wrk sas7bdat "%sysfunc(pathname(work))";
    libname xel excel "d:/xls/class.xlsx" header=no;
    data wrk.class;
      set xel."sheet1$"n(firstobs=4);
    run;quit;
    data wrk.classren(
        rename=(
          f1=name
          f2=sex
          f3=age
          f4=height
          f5=weight
        ));
      set xel."sheet1$"n(firstobs=4);
    run;quit;
    libname xel clear;
    proc print data=wrk.classren;
    run;quit;
    ');
