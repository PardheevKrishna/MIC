/* run_import_and_sql_timed.sas */

/* 0. Enable detailed timing */
options fullstimer;

/* 1. Capture timestamp */
%let t0_import = %sysfunc(datetime());

/* 2. IMPORT table1.csv */
proc import
  datafile="/path/to/csv/folder/table1.csv"
  out=work.table1
  dbms=csv
  replace;
  getnames=yes;
  guessingrows=10000;  /* scan first 10k rows for column types */
run;

/* 3. IMPORT table2.csv */
proc import
  datafile="/path/to/csv/folder/table2.csv"
  out=work.table2
  dbms=csv
  replace;
  getnames=yes;
  guessingrows=10000;
run;

/* 4. IMPORT table3.csv */
proc import
  datafile="/path/to/csv/folder/table3.csv"
  out=work.table3
  dbms=csv
  replace;
  getnames=yes;
  guessingrows=10000;
run;

/* 5. Report import elapsed time */
%let t1_import = %sysfunc(datetime());
%let import_secs = %sysevalf(&t1_import - &t0_import);
%put NOTE: >>> TOTAL IMPORT TIME: &import_secs seconds. <<<;

/* 6. Capture start of PROC SQL */
%let t0_sql = %sysfunc(datetime());

/* 7. Complex PROC SQL */
proc sql;
  create table work.joined as
  select 
    a.key,
    count(distinct b.col1)   as cnt_col1,
    sum(c.col2)              as sum_col2,
    case when avg(a.col3)>50 then 'High' else 'Low' end as Category,
    std(c.col5)              as stddev_col5
  from work.table1 as a
    inner join work.table2 as b on a.key = b.key
    left  join work.table3 as c on a.key = c.key
  where a.col4 between 10 and 90
  group by a.key, calculated Category
  having count(*) > 100
  order by sum_col2 desc
  ;
quit;

/* 8. Report PROC SQL elapsed time */
%let t1_sql = %sysfunc(datetime());
%let sql_secs = %sysevalf(&t1_sql - &t0_sql);
%put NOTE: >>> PROC SQL elapsed time: &sql_secs seconds. <<<;

/* 9. Grand total */
%let total_secs = %sysevalf(&t1_sql - &t0_import);
%put NOTE: >>> TOTAL IMPORT + SQL TIME: &total_secs seconds. <<<;