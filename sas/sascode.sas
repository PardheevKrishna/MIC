options fullstimer; /* Enables detailed timing info */

/* Record start time */
%let start = %sysfunc(datetime());

/* Import the CSV file "million_rows.csv" */
proc import datafile="million_rows.csv"
    out=million
    dbms=csv
    replace;
    guessingrows=MAX;
run;

data million_processed;
    set million;
    /* Initialize computed column for heavy computation */
    computed = 0;
    /* Simulate heavy per‑row computation with an inner loop of 100 iterations */
    do j = 1 to 100;
        computed + sin(value1) * cos(value2) / j;
    end;
    
    /* Debug message: print progress every 100,000 rows */
    if mod(_N_, 100000) = 0 then do;
        put "DEBUG: Processed " _N_ " rows";
    end;
run;

/* Record end time and compute elapsed seconds */
%let end = %sysfunc(datetime());
%put NOTE: Elapsed time for SAS processing: %sysevalf(&end - &start) seconds;