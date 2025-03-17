options fullstimer;  /* Enables detailed resource usage statistics in the log */

/* Record the start time */
%let start_time = %sysfunc(datetime());

data complex_calcs;
   set dummy;
   complex_calc = 0;
   array var[200] var1-var200;
   /* Loop over each column and perform 50 iterations of nested calculations */
   do i = 1 to 200;
      temp = var[i];
      do j = 1 to 50;
         temp = sin(temp) * cos(temp) + log(abs(temp) + 1);
      end;
      complex_calc = complex_calc + temp;
   end;
run;

/* Record the end time */
%let end_time = %sysfunc(datetime());

/* Compute and print the elapsed time in seconds */
%put NOTE: Total SAS runtime in seconds: %sysevalf(&end_time - &start_time);

/* Print first 5 rows for verification */
proc print data=complex_calcs (obs=5);
   title "Sample Output from complex_calcs Dataset";
run;