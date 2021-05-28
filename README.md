Requirement:
Compare sequenced / unsequenced data from two different excels

Flow:

Step 1: Excel initilization, Fetch cell 0 for all rows of file 1--> Add it in hashmap, if any repetative skip that value: else add
Step 2: Iterate over hashmap and store keys in a list
Step 3: Initilize FILLO class objects and pass excel path of both files
Step 4: Run select query for all available keys from file 1for file2 with where condition(firstCellNameOfFile2='KeyFromFile1')
if(keyfound, store retrived data in object array)
else(Exception handled to continue with next key)

Result:

If key 1 is available for any of the data in file 2 then we will get all cells for that row
Ex. "Select * from Sheet2 where file2Cell1Name ='keyFromFileOne'"
OP: 110022, ABCD11, A (my excel1 has 4 rows 3 cell and 3 rows 2 cells in file2)
File2:
Key: 110022 from file1
[110022, VWDE11, A]
[110022, VWDE11, Blank] ( column3 of file2 is blank )

Key: 110023 from file1
[For the key from file1 '110023' No records available in file2]


Files

Cell1        Cell2           Cell3                             Cell1          Cell2          Cell3                     

11022       1234              A                                 11022       1234              A
11022       5678              B                                 11022       5678              
11023       9012              C                                 
