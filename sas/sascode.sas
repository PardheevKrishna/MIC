/* run_complex_sql_timed.sas */

/* Enable detailed timing */
options fullstimer;

/* 1. Point libname to folder where table1.csv/table2.csv/table3.csv live */
libname mycsv csv '/path/to/your/csv/folder';

/* 2. Capture start time */
%let t0 = %sysfunc(datetime());

/* 3. Complex PROC SQL */
proc sql;
  create table work.joined as
  select 
    a.key,
    count(distinct b.col1)   as cnt_col1,
    sum(c.col2)              as sum_col2,
    case when avg(a.col3)>50 then 'High' else 'Low' end as Category,
    std(c.col5)              as stddev_col5
  from mycsv.table1 as a
    inner join mycsv.table2 as b on a.key = b.key
    left  join mycsv.table3 as c on a.key = c.key
  where a.col4 between 10 and 90
  group by a.key, calculated Category
  having count(*) > 100
  order by sum_col2 desc
  ;
quit;

/* 4. Capture end time and print */
%let t1      = %sysfunc(datetime());
%let elapsed = %sysevalf(&t1 - &t0);
%put NOTE: >>> SAS PROC SQL elapsed time: &elapsed seconds. <<<;