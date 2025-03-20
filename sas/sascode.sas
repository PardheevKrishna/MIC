options fullstimer;

/* Assume dataset 'million' already exists with columns: id, value1, value2 */

/* Record start time for heavy processing */
%let start = %sysfunc(datetime());

data processed;
    set million;
    computed = 0;
    /* Simulate heavy computation: inner loop of 100 iterations */
    do j = 1 to 100;
        computed + sin(value1) * cos(value2) / j;
    end;
    output;
    
    /* Debug: Log every 100,000 rows processed */
    if mod(_N_, 100000) = 0 then do;
        put "DEBUG: Processed " _N_ " rows";
    end;
run;

/* Record end time and print elapsed seconds */
%let end = %sysfunc(datetime());
%put NOTE: Elapsed time for heavy processing: %sysevalf(&end - &start) seconds;