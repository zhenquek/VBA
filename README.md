# VBA
VBA workpapers created to ease office work 


## OracleConverter
Used in financial institution making 50x faster and efficient extraction of Oracle txt data into usable spreadsheet

###Worksheet "Input"

Col A - checks type of transaction and whether it will be included or excluded in output of below
```=IF(ISNUMBER(SEARCH("Default.",B18)),1,IF(AND(LEN(B18)>160,LEN(B18)<165,NOT(ISNUMBER(SEARCH("---------------------------------------------------------",B18)))),2,""))```

Row C to V - split by TextToColumn function

Row1
```
6,4,6,3,1,6,2,5,3,6,4,0,12,18,21,15,33,28,15,142
```
Row2
```
Segment,Cost Centre,Account,Categories,State,Product,Class,Channel,Project,Intercompany,Spare,n,Source,Category,Batch,Number,Description,Line Item,Dates,Amount
```

Determine initial nil value
```
Cell(C5) = 0
```

Replicate across C6 to M6
```
=IF($A6=1,IFERROR(MID($B6,SUM($B$1:B$1)+COUNTA($B$1:B$1),C$1),""),C5)
```

Col N is intentionally blank (acted as buffer area - not necessary)
Col O to U
```
=TRIM(IF($A6=2,IFERROR(MID($B6,SUM($N$1:N$1)+COUNTA($N$1:N$1),O$1),""),""))
```
Col V = Amount
```
=IFERROR(VALUE(IF($A6=2,(MID($B6,143,V$1)),"")),"")
```

### Worksheet "WIP"
Used as intemediary before outputting data to clear all blank data points
Row1
```
Segment,Cost Centre,Account,Categories,State,Product,Class,Channel,Project,Intercompany,Spare,n,Source,Category,Batch,Number,Description,Line Item,Dates,Amount
```

### Worksheet "Output"
Used as output spreadsheet for extraction purposes
Row1
```
Segment,Cost Centre,Account,Categories,State,Product,Class,Channel,Project,Intercompany,Spare,n,Source,Category,Batch,Number,Description,Line Item,Dates,Amount
```


## MakeCSV





