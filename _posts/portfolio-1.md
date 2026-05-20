---
title: "International Migration and Climate Change"
excerpt: "**Methodology**: Logitistic Regression.<br/><img src='/images/portfolio_image1.png'>"
collection: portfolio
lang: "STATA"
date: 2025-02-28
---

You can read the abstract from here:  [👀📖✨ Read](/projects/project10)


```stata
** For checking current directory

pwd

** For changing the directory to desired one

cd "D:\Munem\BRAC_CC_IntMig\Data Analysis\BRAC_CC_IntMIg_29012025"

** Importing excel file

import excel using "Clean Data_BRAC_CC_IntMig_29012025", firstrow clear
 
** Opening a log file

log using BRACresult30012025

** Renaming Variables
rename NameofEnumeraotor enu_name
rename Id id
rename NameoftheRespondent res_name
rename CellNumber cell
rename District district
rename Union union
rename Age age
rename Gender gender
rename EducationalQualification edu
rename MaritalStatus marital
rename Religion religion


rename q19_Salinity q19_1
rename q19_Rivererosion q19_2
rename q19_Flooding q19_3
rename q19_Cyclones q19_4
rename q19_Drought q19_5
rename q19_WaterLogging q19_6 
rename q19_Tidalwave q19_7


*** new variable creation

** district_new 
encode district , gen (district_new)
tab district_new
tab district_new , nol

** marital_new
encode marital , gen (marital_new)
tab marital_new
tab marital_new , nol

** religion_new
encode religion , gen (religion_new)
tab religion_new
tab religion_new , nol

** age_cat
recode age (18/25 = 1 "18 to 25 years") (26/35 = 2 "26 to 35 years") (36/45 = 3 "36 to 45 years") (46/55 = 4 "46 to 55 years") (56/65 = 5 "56 to 65 years") (66/75 = 6 "66 to 75 years") (76/80 = 7 "76 or above") , gen (age_cat)
tab age_cat 


** q41_newv2

encode q41 , gen (q41_new)
tab q41_new
tab q41_new , nol

gen q41_newv2 = (q41_new == 2)

su q41_newv2
tab q41_newv2

label define q41_newv2_l 1 "Yes" 0 "No" 
label values q41_newv2 q41_newv2_l


** q39_newv2

encode q39 , gen (q39_new)

tab q39_new
tab q39_new, nol 

gen q39_newv2 = .
replace q39_newv2 = 1 if q39_new <= 2
replace q39_newv2 = 0 if q39_new > 2

label define q39_newv2 0 "No impact or Increased" 1 "Decreased"
label values q39_newv2 q39_newv2

tab q39_newv2
tab q39_newv2, nol


** q40_new 

encode q40 , gen (q40_new)
tab q40_new
tab q40_new, nol

** q43_newv2

encode q43 , gen (q43_new)
tab q43_new
tab q43_new, nol 
gen q43_newv2 = .
replace q43_newv2 = 1 if q43_new == 3
replace q43_newv2 = 0 if q43_new <= 2
label define q43_newv2 0 "No or Not sure" 1 "Yes"
label values q43_newv2 q43_newv2
tab q43_newv2
tab q43_newv2, nol



* q10_new

gen q10_new = .

replace q10_new = 1 if q10 <= 3
replace q10_new = 2 if q10 > 3 & q10 <= 6
replace q10_new = 3 if q10 > 6 & q10 <= 9
replace q10_new = 4 if q10 > 9

label define q10_new 1 "Less than or equal to 3" 2 "4 to 6" 3 "7 to 9" 4 "Above 9" 
label values q10_new q10_new

tab q10_new
tab q10_new, nol


* q11_new

gen q11_new = .

replace q11_new = 1 if q11 <= 2
replace q11_new = 2 if q11 > 2 & q11 <= 4
replace q11_new = 3 if q11 > 4 & q11 <= 6


label define q11_new 1 "Less than or equal to 2" 2 "3 to 4" 3 "5 to 6" 
label values q11_new q11_new

tab q11_new
tab q11_new, nol


**** Economic class for the households in the current situation

*HH per capita income 
gen per_capita_income_current = q13/ q10 

gen eco_status_current = .

replace eco_status_current = 1 if per_capita_income_current <= 1800
replace eco_status_current = 2 if per_capita_income_current > 1801 & per_capita_income_current <= 5010
replace eco_status_current = 3 if per_capita_income_current > 5011 & per_capita_income_current <= 8370
replace eco_status_current = 4 if per_capita_income_current > 8371


label define eco_status_current 1 "Poor" 2 "Lower Middle Income" 3 "Upper Middle Income" 4 "Upper Class" 
label values eco_status_current eco_status_current

tab eco_status_current
tab eco_status_current, nol


*** labeling variables

label variable district_new "Numeric value district"
label variable marital_new "Numeric value of marital status categories"
label variable religion_new "Numeric value of religion category"
label variable age_cat "Age category"


*** Labeling variable values  

** q17_Gender

label define q17_Gender 1 "Male" 2 "Female"
label values q17_Gender q17_Gender
labelbook q17_Gender

** q42 
label define q42 1 "Salinity intrusion in agricultural land" 2 "River erosion" 3 "Cyclones and storm surges" 4 "Flooding or waterlogging" 5 "Drought" 6 "Tidal Wave"
label values q42 q42
tab q42
tab q42, nol



** Descriptive statistics

tab gender
tab q12
tab q16
tab q17_Gender




* install mrtab for multiple response analysis
ssc install mrtab

* multiple response of q18

mrtab q18_1 q18_2 q18_3 q18_4 q18_5 q18_6
mrtab q18_1 q18_2 q18_3 q18_4 q18_5 q18_6 , by (district_new) column 
 
* multiple response of q19

mrtab q19_1 q19_2 q19_3 q19_4 q19_5 q19_6 q19_7  
mrtab q19_1 q19_2 q19_3 q19_4 q19_5 q19_6 q19_7 , by (district_new) column 

* Cross table q20 & q21 with district_new

tab q20 district_new , column
tab q21 district_new , column

* multiple response of q22 

mrtab q22_1 q22_2 q22_3 q22_4 q22_5 q22_6 q22_7 q22_8 
mrtab q22_1 q22_2 q22_3 q22_4 q22_5 q22_6 q22_7 q22_8 , by (district_new) column

* Cross table q23 & q24 with district_new

tab q23 district_new , column
tab q24 district_new , column

* multiple response of q25
 
mrtab q25_1 q25_2 q25_3 q25_4 q25_5 q25_6 q25_7 q25_8 q25_9 q25_10  
mrtab q25_1 q25_2 q25_3 q25_4 q25_5 q25_6 q25_7 q25_8 q25_9 q25_10 , by (district_new) column

* Cross table q26 & q27 with district_new

tab q26 district_new , column
tab q27 district_new , column

* multiple response of q28
mrtab q28_1 q28_2 q28_3 q28_4 q28_5 q28_6 q28_7 q28_8
mrtab q28_1 q28_2 q28_3 q28_4 q28_5 q28_6 q28_7 q28_8 , by (district_new) column

* Cross table q29 & q30 with district_new
tab q29 district_new , column
tab q30 district_new , column

* multiple response of q31

mrtab q31_1 q31_2 q31_3 q31_4 q31_5 q31_6 q31_7 q31_8 
mrtab q31_1 q31_2 q31_3 q31_4 q31_5 q31_6 q31_7 q31_8 , by (district_new) column

* Cross table q32 & q33 with district_new

tab q32 district_new , column
tab q33 district_new , column

* multiple response of q34

mrtab q34_1 q34_2 q34_3 q34_4 q34_5 q34_6 q34_7
mrtab q34_1 q34_2 q34_3 q34_4 q34_5 q34_6 q34_7 , by (district_new) column 

* Cross table q35 & q36 with district_new

tab q35 district_new , column
tab q36 district_new , column

* multiple response of q37

mrtab q37_1 q37_2 q37_3 q37_4 q37_5 q37_6 q37_7  
mrtab q37_1 q37_2 q37_3 q37_4 q37_5 q37_6 q37_7  , by (district_new) column

* multiple response of q38

mrtab q38_1 q38_2 q38_3 q38_4 q38_5 q38_6 q38_7 q38_8 q38_9 q38_10 q38_11
mrtab q38_1 q38_2 q38_3 q38_4 q38_5 q38_6 q38_7 q38_8 q38_9 q38_10 q38_11 , by (district_new) column 

* Cross table q39 t0 q45 with district_new

tab q39 district_new , column
tab q40 district_new , column
tab q41 district_new , column
tab q42 district_new , column
tab q43 district_new , column
tab q44 district_new , column
tab q44_other district_new , column
tab q45 district_new , column

*multiple response of q46

mrtab q46_1 q46_2 q46_3 q46_4 q46_5 q46_6 q46_7 q46_8 q46_9 q46_10
mrtab q46_1 q46_2 q46_3 q46_4 q46_5 q46_6 q46_7 q46_8 q46_9 q46_10 , by (district_new) column 

* multiple response of q47  
mrtab q47_1 q47_2 q47_3 q47_4 q47_5
mrtab q47_1 q47_2 q47_3 q47_4 q47_5 , by (district_new) column 


* Summary Statistics of q48 
su q48 



* multiple response of q50 

mrtab q50_1 q50_2 q50_3 q50_4 q50_5 q50_6 q50_7 q50_8 q50_9 q50_10 q50_11 q50_other 

tab q50_other_res district_new , column

* Cross table q51 t0 q58 with district_new


tab q51 district_new , column
tab q52 district_new , column
tab q53 district_new , column
tab q54 district_new , column
tab q55 district_new , column
tab q56 district_new , column
tab q57 district_new , column
tab q58 district_new , column


* multiple response of q59

mrtab q59_1 q59_2 q59_3 q59_4 q59_5 q59_6 q59_Other 
mrtab q59_1 q59_2 q59_3 q59_4 q59_5 q59_6 q59_Other , by (district_new) column 
tab q59_Other_res district_new , column

* Cross table q60 t0 q61 with district_new

tab q60 district_new , column
tab q61 district_new , column


* New Variable for q60

encode q60 , gen (q60_new)
tab q60_new
tab q60_new , nol

* multiple response of 62

mrtab q62_1 q62_2 q62_3 q62_4 q62_5 q62_6 q62_other
mrtab q62_1 q62_2 q62_3 q62_4 q62_5 q62_6 q62_other , by (district_new) column 
tab q62_other_res district_new , column

* Summary Statistics of q63 
su q63

* Cross table q64 t0 q65 with district_new

tab q64 district_new , column
tab q65 district_new , column





*** create index by Multiple Correspondence Analysis (MCA)

*For climate events q19 
ssc install mca
mca q19_1 q19_2 q19_3 q19_4 q19_5 q19_6 q19_7 


predict mca_sc1 mca_sc2

gen composite_index_cc = 0.56 * mca_sc1 + 0.08 * mca_sc2


* For Push factors q46 
mca q46_1 q46_2 q46_3 q46_4 q46_5 q46_6 q46_7 q46_8 q46_9 q46_10

predict mca_sc3 mca_sc4

gen composite_index_pus = 0.498 * mca_sc3 + 0.137 * mca_sc4

* For Pull factors q47 
mca q47_1 q47_2 q47_3 q47_4 q47_5 

predict mca_sc5 mca_sc6 

gen composite_index_pul = 0.538 * mca_sc5 + 0.116 * mca_sc6 



**** Economic class for the households before international migration

* HH per capita income 
gen per_capita_income_bintmig = q48/ q10 

gen eco_status_bintmig = .

replace eco_status_bintmig = 1 if per_capita_income_bintmig <= 1800
replace eco_status_bintmig = 2 if per_capita_income_bintmig > 1801 & per_capita_income_bintmig <= 5010
replace eco_status_bintmig = 3 if per_capita_income_bintmig > 5011 & per_capita_income_bintmig <= 8370
replace eco_status_bintmig = 4 if per_capita_income_bintmig > 8371


label define eco_status_bintmig 1 "Poor" 2 "Lower Middle Income" 3 "Upper Middle Income" 4 "Upper Class" 
label values eco_status_bintmig eco_status_bintmig

tab eco_status_bintmig
tab eco_status_bintmig, nol





*** logit individual climate impacts

* For Salinity

logit q41_newv2 q19_1 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig, or 
margins , dydx (*)
margins, at(eco_status_bintmig =(1 2 3 4))
ssc install fitstat
fitstat

* For River Erosion

logit q41_newv2 q19_2 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig, or
margins , dydx (*)
margins, at(eco_status_bintmig = (1 2 3 4))
fitstat



* For Flooding 

logit q41_newv2 q19_3 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig, or 
margins , dydx (*)
margins, at(eco_status_bintmig =(1 2 3 4))
fitstat


* For Cyclones

logit q41_newv2 q19_4 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig , or 
margins , dydx (*)
margins, at(eco_status_bintmig =(1 2 3 4))
fitstat


* For Drought

logit q41_newv2 q19_5 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig , or 
margins , dydx (*)
margins, at(eco_status_bintmig =(1 2 3 4))
fitstat


* For Water Logging

logit q41_newv2 q19_6 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig , or 
margins , dydx (*)
margins, at(eco_status_bintmig=(1 2 3 4))
fitstat


* For Tidal Wave

logit q41_newv2 q19_7 composite_index_pul composite_index_pus  q43_newv2 i.eco_status_bintmig , or
margins , dydx (*)
margins, at(eco_status_bintmig =(1 2 3 4))
fitstat


** cross table for retunees income: present income of returnee X income before recieving reintegration services
table q60_new district_new, statistic(mean q48) statistic(mean q13)


log close
save "data30012025.dta", replace
export excel "data30012025.xlsx" , firstrow(variables) replace

```

