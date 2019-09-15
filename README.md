# Linear_Cut_Calculator
2D linear cut optimizer: gives number of Lengths required and instructions to cut | C# Application


Prerequisites: The aaplication uses SAP Crystal Report so please download the version below.
http://downloads.businessobjects.com/akdlm/crnetruntime/clickonce/CRRuntime_32bit_13_0_23.msi


Welcome to 2D Linear Cut Optimizer Application

Problem: Suppose you have infinite number of 10" length rods and you need 20 of 2", 40 of 5", and 200 of 3" lengths!
         How would you cut them so that you use the least number of 10" rods thereby assuring least wastage of steel!!

Solution:

Trivial: My Application LOL!!

Non-Trivial:
1) The best way to solve this is by treating this as Knapsack Problem! Read more here: https://en.wikipedia.org/wiki/Knapsack_problem
2) I wish I knew how to implement that but unfortunately I don't. If someone can teach me in Mathematics, I can for sure implement in C# or any language.
3) I am using some part of Knapsack algorithm and it works and is better than hundreds of out there.

Now a little bit about the application:

Features: 1) Lets you import Lengths and Quantities of cut required from Excel sheet!
          2) Spits out a PDF of instructions that you need to follow to use least number of rods.
          3) Lets you enter "SAW WIDTH". In mills, the saw blade does the cutting and with every cut, a part of steel is lost.

Running Instructions: There are two folders in the repository-
                      1) Application - This has the .exe of the application with a config file. Please copy both files to the
                                       same folder before running the application. Config contains all libraries' dlls.
                      2) Code - This is self-explanatory! 

This is an open source application!!

I would love to hear any reviews, suggestion, or anything about this application!!
HELP ME MAKE THIS BETTER!!!!
