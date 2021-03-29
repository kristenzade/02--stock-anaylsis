# Refactoring Stock Anaylsis

## Overview of Project
#Testing efficiency of VBA

### Purpose
#This analysis will show the logical errors easily appear in well structure code that contains nested conditionals and loops.


### Analysis of Outcomes
#Refactoring process can affect the testing outcomes. 

### Results
This is a long procedure may contain the same line of code in several locations. A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions. A complex unstructured code may be better broken down.

