options fullstimer; /* Enables detailed timing info in the log */

/* Record start time */
%let start = %sysfunc(datetime());

data million;
    call streaminit(123); /* Initialize random number generator */
    do id = 1 to 1000000;
        /* Generate random values similar to Python */
        value1 = rand("Uniform");
        value2 = floor(rand("Uniform") * 100);
        
        /* Simulate heavy computation: inner loop of 100 iterations */
        computed = 0;
        do j = 1 to 100;
            computed + sin(value1) * cos(value2) / j;
        end;
        
        output;
        
        /* Debug message every 100,000 rows */
        if mod(id, 100000) = 0 then do;
            put "DEBUG: Processed " id " rows";
        end;
    end;
run;

/* Record end time and calculate elapsed seconds */
%let end = %sysfunc(datetime());
%put NOTE: Elapsed time for SAS processing: %sysevalf(&end - &start) seconds;