/* --- Step 1: Import the CSV File --- */
/* Adjust the file path below as per your local setup */
proc import datafile="C:\Path\to\million_rows.csv"
            out=million
            dbms=csv
            replace;
     getnames=yes;
run;

/* --- Step 2: Heavy Processing with Runtime Logging --- */
options fullstimer;  /* Enables detailed timing info in the SAS log */

/* Record the start time */
%let start = %sysfunc(datetime());

data processed;
    set million;
    
    /* Simulate heavy computation: inner loop of 100 iterations */
    computed = 0;
    do j = 1 to 100;
        computed + sin(value1) * cos(value2) / j;
    end;
    
    output;
    
    /* Debug message: print every 100,000 rows processed */
    if mod(_N_, 100000) = 0 then do;
        put "DEBUG: Processed " _N_ " rows";
    end;
run;

/* Record the end time and print the elapsed time */
%let end = %sysfunc(datetime());
%put NOTE: Elapsed time for heavy processing: %sysevalf(&end - &start) seconds;