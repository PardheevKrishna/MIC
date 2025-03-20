options fullstimer;

data million;
    call streaminit(123); /* Initialize random number generator */
    do id = 1 to 1000000;
        /* Generate random numbers */
        value1 = rand("Uniform");
        value2 = floor(rand("Uniform") * 100);
        
        /* Simulate heavy computation */
        computed = 0;
        do j = 1 to 100;
            computed + sin(value1) * cos(value2) / j;
        end;
        
        output;
        
        /* Log progress every 100,000 rows */
        if mod(id, 100000) = 0 then do;
            put "Processed " id " rows";
        end;
    end;
run;