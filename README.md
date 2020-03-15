# algoTest

**Assumption**<br/>

The metadata file should be **xlsx** file and its structure should be like first column would be an identifier and second column should be a key column. Rest columns wil be non-key column and number can be veried. This file should contain atleast 3 columns.<br/>

**Structure**<br/>

**data_files**<br/>
     It contains xlsx files to run the test. New test dataset can be stored here. It also contains result files for the test.
     For Wilcoxon Signed Rank Test, result file name will be **result.xlsx** and for permutations it will be **permutation_result.xlsx**<br/>
     `For changing the file names; you have to pass it whle object intialization. These are named parameters.`<br/>
     `To change metadata file name pass file name to excel_name=<new_name>`<br>
     `To change result file name pass it as result_file_name=<file_name>`

**execute_permutation_test.py**<br/>
     To run Wilcoxon Signed Rank Test of permutations; run below command:-<br/>
>          python execute_permutation_test.py

**execute_test.py**<br/>
     To run test Wilcoxon Signed Rank Test; run below command:-<br/>
>         python execute_test.py

**permutations.py**<br/>
     It consists of a logic for getting permutations of data given in xlsx.

**rank_test.py**<br/>
     Logic for performing Wilcoxon Signed Rank         

**read_excel.py**<br/>
     Logic to read xlsx file         

**write_excel.py**<br/>
     Logic to write xlsx file
