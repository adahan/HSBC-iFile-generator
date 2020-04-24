# HSBC-iFile-generator
Python based script that take a .CSV with  Name, BankID, BranchNo, AccountNo, Amount, Reference as columns and generate a .TXT file in HSBC ifile format

This script implement ACH-CR format of ifile, can be used for exemple for Payroll batch payment with HSBC

Could easily be modify to implement other format or add email notification for customers

Also added frozen version of the .py script => .exe that can be executed on windiws for those who dont know how to do it. Never recommended as you cant really know whats inside and you are dealing with payment data...  use at your own risks. 
