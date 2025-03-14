/* ----------------------------------------------------------
   SAS Code: Import dummy_data.xpt and perform complex tasks
   ---------------------------------------------------------- */

/* Start timer */
%let start_time = %sysfunc(datetime());

/* --- Import the SAS XPORT file --- */
libname myxpt xport 'dummy_data.xpt';

/* Create a working dataset from the XPT file (assumes one dataset inside) */
data dummy_data;
   set myxpt._all_;
run;

/* --- Task 1: Group Aggregation ---
     Compute means for numeric variables num_1 - num_100 grouped by cat_1 */
proc means data=dummy_data noprint;
   class cat_1;
   var num_1 - num_100;
   output out=group_means mean=;
run;

/* --- Task 2: Join ---
     Join the average of num_1 back to the original dataset */
proc sql;
   create table joined_data as
   select a.*, b.mean_num_1
   from dummy_data as a
   left join
      (select cat_1, mean(num_1) as mean_num_1
       from dummy_data
       group by cat_1) as b
   on a.cat_1 = b.cat_1;
quit;

/* --- Task 3: Data Transformation ---
     Create new columns doubling num_1 to num_10 */
data transformed;
   set dummy_data;
   %do i = 1 %to 10;
      double_num_&i = num_&i * 2;
   %end;
run;

/* --- Task 4: Sorting ---
     Sort the transformed data by num_1 */
proc sort data=transformed out=sorted_data;
   by num_1;
run;

/* --- Task 5: Pivoting ---
     Transpose num_1 by cat_1 (one row per category) */
proc transpose data=dummy_data out=pivot_data;
   by cat_1;
   var num_1;
run;

/* Stop timer and report elapsed time */
%let end_time = %sysfunc(datetime());
%let elapsed = %sysevalf(&end_time - &start_time);
%put NOTE: Total SAS runtime (seconds): &elapsed;