# VBA-challenge
## Overview of Project

### Purpose
Given multiple years' data of stock information, produce a VBA script that loops through each year's stock information and output each ticker's yearly change, percent change, total stock volume, and the greatest increases and decreases annually. 

## Analysis and Challenges
While I was able to correctly code to aggregate ticker information to show yearly change, percent change, and total stock volume for the alphabetical_testing workbook, the code did not work the same in the Multiple_year_stock_data workbook. The error showed "End With without With" and could not figure out what in my argumentation resulted in that. I checked for missing "End If" statements and double-checked to start with a With statement and have been left dumbfounded.
Another challenge I had was displaying the Greatest % Increase + Greatest % Decrease. For whatever reason, I would only get 100% for the greatest and either 0% or -100% for the biggest decrease. I tried using the max and min functions but neither gave me what I needed. Below is the last version of my worksheet loop before submitting
<img width="708" alt="Sample VBA Code" src="https://user-images.githubusercontent.com/114324871/229019412-5b4736dc-d614-49a4-9d37-57fbf3b320f9.png">



