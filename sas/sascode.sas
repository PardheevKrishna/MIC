options fullstimer;  /* Enables detailed resource usage statistics in the SAS log */

/* Step 1: Read the XPORT file */
/* Adjust the file path below to the location of your dummy_data.xpt file */
libname myxpt xport 'C:\path\to\dummy_data.xpt';

/* Create a SAS dataset from the XPORT file. 
   (Assumes the dataset inside the XPORT file is named dummy_data) */
data dummy;
   set myxpt.dummy_data;
run;

/* Step 2: Record the start time */
%let start_time = %sysfunc(datetime());

/* Step 3: Perform complex calculations on each row */
/* Assumes the dataset dummy contains 200 numeric columns named var1 - var200 */
data complex_calcs;
   set dummy;
   complex_calc = 0;
   array var[200] var1-var200;
   /* Loop over each of the 200 columns */
   do i = 1 to 200;
      temp = var[i];
      /* Perform 50 iterations of complex operations on each column value */
      do j = 1 to 50;
         temp = sin(temp) * cos(temp) + log(abs(temp) + 1);
      end;
      complex_calc = complex_calc + temp;
   end;
   drop i j temp;
run;

/* Step 4: Record the end time and calculate the elapsed time */
%let end_time = %sysfunc(datetime());
%put NOTE: Total SAS runtime in seconds: %sysevalf(&end_time - &start_time);

/* Step 5: Print the first 5 rows for verification */
proc print data=complex_calcs (obs=5);
   title "Sample Output from complex_calcs Dataset";
run;