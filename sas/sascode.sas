options fullstimer;  /* Enables detailed resource usage statistics in the log */

libname myxpt xport 'C:\path\to\dummy_data.xpt';

/* Step 1: Extract dataset name dynamically */
proc contents data=myxpt._all_ out=dsnames noprint;
run;

/* Step 2: Load the dataset (update dataset name if needed) */
data dummy;
   set myxpt.dummy_data; /* If different, replace dummy_data with correct name */
run;

/* Step 3: Print first few rows to verify */
proc print data=dummy (obs=5);
run;

/* Step 4: Capture start time */
%let start_time = %sysfunc(datetime());

/* Step 5: Perform complex calculations */
data complex_calcs;
   set dummy;
   complex_calc = 0;
   array var[*] _numeric_; /* Automatically selects all numeric columns */
   do i = 1 to dim(var);
      temp = var[i];
      do j = 1 to 50;
         temp = sin(temp) * cos(temp) + log(abs(temp) + 1);
      end;
      complex_calc = complex_calc + temp;
   end;
   drop i j temp;
run;

/* Step 6: Capture end time and calculate elapsed time */
%let end_time = %sysfunc(datetime());
%put NOTE: Total SAS runtime in seconds: %sysevalf(&end_time - &start_time);

/* Step 7: Show sample results */
proc print data=complex_calcs (obs=5);
run;