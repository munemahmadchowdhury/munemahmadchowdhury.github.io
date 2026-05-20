---
title: "Readiness for the Fourth Industrial Revolution (4IR)"
excerpt: "**Methodology**: Likeart Scale Based Mean Score. <br/><img src='/images/portfolio_image2.png'>"
collection: portfolio
lang: "STATA"
date: 2025-08-29
---

You can read the abstract from here:  [👀📖✨ Read](/projects/project9)


```stata

//Please see the related questionnaire for students for better understanding the variables. For example, q7 variable here indication the question 7 in the respective questionnaire which is "Gender related"

// Know working directory
pwd

// Change working directory
cd "D:\Munem\SSRC_4.0\Data\Analysis\18032025"

//Importing excel file

import excel using "Student_31082025_semiclean_revised", firstrow clear


/******************************************************
*				Treating Missing Values				  *
******************************************************/

//https://www.statalist.org/forums/forum/general-stata-discussion/general/1466362-plug-in-a-mean-into-the-missing-values


*** Error in q17_e. Though it's a numeraic variable the stata identifies it as string. Hence the variable is going to be converted from string to numeric

describe q17_e 
destring q17_e, gen (q17_e_numeric)



foreach var in q6 q8 q11 q12 q13 q14 q15 q16  q17_a q17_b q17_c q17_d q17_e_numeric q17_f q17_g q17_h q17_i q18_a q18_b q18_c q18_j q18_k q18_l q18_m q18_n q18_o q18_p q19 q21 q22_1 q22_2 q22_3 q22_3 q22_4 q22_other q23 q24 q25 q26 q27 q29 q30 q31 q32 q33_a q33_b q33_c q34_a q34_b q34_c q34_d q34_e q35_a q35_b q35_c q35_d q36_1 q36_2 q36_3 q36_4 q36_5 q36_6 q36_other q37_1 q37_2 q37_3 q37_other q38 q39 q40 q41 q42 q43 q44 q45 q46 q47 q48 q49 {
    quietly summarize `var', meanonly
    local roundedmean = round(r(mean),1)
    replace `var' = `roundedmean' if missing(`var')
}


misstable summarize q6 q8 q11 q12 q13 q14 q15 q16  q17_a q17_b q17_c q17_d  q17_f q17_g q17_h q17_i q18_a q18_b q18_c q18_j q18_k q18_l q18_m q18_n q18_o q18_p q19 q21 q22_1 q22_2 q22_3 q22_3 q22_4 q22_other q23 q24 q25 q26 q27 q29 q30 q31 q32 q33_a q33_b q33_c q34_a q34_b q34_c q34_d q34_e q35_a q35_b q35_c q35_d q36_1 q36_2 q36_3 q36_4 q36_5 q36_6 q36_other q37_1 q37_2 q37_3 q37_other q38 q39 q40 q41 q42 q43 q44 q45 q46 q47 q48 q49

misstable summarize




// Recoding Institution name variables
tab q3 
encode q3, gen (q3_new)
tab q3_new
describe q3_new
labelbook q3_new


gen q3_type = . 
replace q3_type  = 1 if q3_new == 2 | q3_new == 4 | q3_new == 7 | q3_new == 8
replace q3_type  = 2 if q3_new == 1 | q3_new == 3 | q3_new == 5 | q3_new == 6
label define q3_type  1 "Public Institute" 2 "Private Institute" 
label values q3_type q3_type
label variable q3_type "Type of Institution"
tab q3_type


// Recoding academic year 
* (academic year 1st to 3rd & 2nd to 4th converted))
tab q4 
encode q4, gen (q4_new)
tab q4_new
describe q4_new
labelbook q4_new



//Recoding age variable, from numeric to category

tabstat q6, statistics (min max) 
recode q6 (17 = 1 "Below 18 years") (18/19 = 2 "18 to 19 years") (20/21 = 3 "20 to 21 years") (22/23 = 4 "22 to 23 years") (24/25 = 5 "24 and Above") , gen (q6_cat)
tab q6_cat 
label variable q6_cat "Age category"



//Re-defining Gender variable 
label define gender 1 "Female" 2 "Male" 3 "Third Gender" 
label values q7 gender 
label variable q7 "Gender of the Student"
tab q7


//Recoding income variable, from numeric to category
gen q8_cat = . 
replace q8_cat = 1 if q8 < 5000
replace q8_cat = 2 if q8 >= 5000 & q8 <= 14999
replace q8_cat = 3 if q8 >= 15000 & q8 <= 29999
replace q8_cat = 4 if q8 >= 30000 & q8 <= 44999
replace q8_cat = 5 if q8 >= 45000 & q8 <= 59999 
replace q8_cat = 6 if q8 >= 60000
label define income_category 1 "Less than 5000" 2 "5000 to 14999" 3 "15000 to 29999" 4 "30000 to 44999" 5 "45000 to 59000" 6 "60000 and above"
label values q8_cat income_category
label variable q8_cat "Category of Monthly income of Student's Family"
tab q8_cat 

// Rcoding device owned variables

label variable q9_1 "Android Mobile"
label variable q9_2 "Laptop"
label variable q9_3 "Desktop Computer"
label variable q9_4 "Tab"
label variable q9_other "Other"

label variable q10_1 "For using social media (Facebook, YouTube, TikTok, Instagram, and others)" 
label variable q10_2 "For visiting various websites" 
label variable q10_3 "For watching movies, series or dramas" 
label variable q10_4 "For playing games" 
label variable q10_5 "For formal education" 
label variable q10_6  "For acquiring online education (Coursera, Bohubrihi, Udemy, Code Academy and others)" 
label variable q10_7 "For freelancing" 
label variable q10_other "Other"



//// Re-defining label values of quality of consumed internet related variable

label define q11 1 "Yes, Good" 0 "No, Bad" 2 "Workable, not bad"
label values q11 q11


// Re-defining label values of student readiness (knowledge & attitude) variables

label define q12 1 "Yes" 0 "No" 
label values q12 q12 

label define q13 1 "Yes" 0 "No" 
label values q13 q13

label define q14 1 "Yes" 0 "No" 
label values q14 q14

label define q15 1 "Yes" 0 "No" 
label values q15 q15

label define q16 1 "Yes" 0 "No" 
label values q16 q16

//Label defining for variables of Tecnical (Software) skills 4.0 

label define Data_Science 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_a Data_Science
label variable q17_a "Data Science"

label define Internet_of_Things 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_b Internet_of_Things
label variable q17_b "Internet of Things"

label define Cyber_security_Encription 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_c Cyber_security_Encription
label variable q17_c "Cyber Security Encription"

label define Artificial_Intelligence 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_d Artificial_Intelligence
label variable q17_d "Artificial Intelligence"


label define Cloud_Computing 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_e_numeric Cloud_Computing
label variable q17_e_numeric "Cloud Computing"
tab q17_e_numeric


label define Augmented_Virtual_Reality 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_f Augmented_Virtual_Reality
label variable q17_f "Augmented Virtual Reality"


label define Nonhumanoid_Robot 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_g Nonhumanoid_Robot
label variable q17_g "Nonhumanoid & Robot"

label define Printing_3D 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_h Printing_3D
label variable q17_h "3D Printing"

label define Generative_AI 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q17_i Generative_AI
label variable q17_i "Generative AI"


//Label defining for variables of Digital literacy


label define MS_Office 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q18_a MS_Office
label variable q18_a "MS Office"


label define Email 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q18_b Email
label variable q18_b "Email"


label define Video_Conferencing 5 "Very Familiar (Expert)" 4 "Familiar (Regular user)" 3 "Moderately familiar (Occasional User)" 2 "Somewhat familiar (Knows but hasn't used)" 1 "Completely unfamiliar"
label values q18_c Video_Conferencing
label variable q18_c "Video Conferencing"


label define SEO 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_j SEO
label variable q18_j "SEO"


label define Password 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_k Password
label variable q18_k "Password"


label define Scam_Virus 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_l Scam_Virus 
label variable q18_l "Scam & Virus Attack"


label define Fact_checking 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_m Fact_checking
label variable q18_m "Fact Checking"


label define Device_problem 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_n Device_problem 
label variable q18_n "Device Problem"


label define Video_tutorial 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_o Video_tutorial
label variable q18_o "Video Tutorial"


label define Copyright 5 "Well aware (Expert)" 4 "Aware (Regular user)" 3 "Moderately aware (Occasional User)" 2 "Somewhat aware (Knows but hasn't used)" 1 "Not aware"
label values q18_p Copyright
label variable q18_p "Copyright"


//Label defining for variables of freelancing

label define freelancing 1 "Yes" 0 "No" 
label values q19 freelancing
label variable q19 "Freelancing"

label define work_type 1 "Freelancing" 2 "Tech oriented job" 88 "Other"
label values q20 work_type
label variable q20 "Type of Freelancing"


//Label defining for variables of Teacher's Readiness

label define q21 1 "Yes" 0 "No" 
label values q21 q21

label variable q22_1 "Lecture Based"
label variable q22_2 "Interactive"
label variable q22_3 "Problem Solving"
label variable q22_4 "Case Study Based"
label variable q22_other "Other"


label define q23 1 "Aware" 2 "Somewhat aware" 3 "Not aware" 
label values q23 q23

label define q24 1 "Always do" 2 "Sometimes do" 3 "Do not do at all" 
label values q24 q24

label define q25 1 "Supportive" 2 "Somewhat supportive" 3 "Not supportive" 
label values q25 q25

label define q26 1 "Agree" 2 "Somewhat agree" 3 "Not agree" 
label values q26 q26



// Label defining for variables of Institutional Readiness


label define q27 5 "Very much" 4 "Much" 3 "Moderate" 2 "A little" 1 "Not at all"
label values q27 q27


label define q28 5 "Very much" 4 "Much" 3 "Moderate" 2 "A little" 1 "Not at all"
label values q28 q28

label define q29 5 "Very much" 4 "Much" 3 "Moderate" 2 "A little" 1 "Not at all"
label values q29 q29

label define q30 5 "Very much" 4 "Much" 3 "Moderate" 2 "A little" 1 "Not at all"
label values q30 q30


label define q31 5 "Very much" 4 "Much" 3 "Moderate" 2 "A little" 1 "Not at all"
label values q31 q31

label define q32 5 "Very much" 4 "Much" 3 "Moderate" 2 "A little" 1 "Not at all"
label values q32 q32


label define q33_a 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q33_a q33_a

label define q33_b 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do"
label values q33_b q33_b

label define q33_c 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do"
label values q33_c q33_c

label define q34_a 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q34_a q34_a

label define q34_b 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do"
label values q34_b q34_b

label define q34_c 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q34_c q34_c

label define q34_d 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q34_d q34_d

label define q34_e 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do"
label values q34_e q34_e

label define q35_a 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do"
label values q35_a q35_a

label define q35_b 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q35_b q35_b

label define q35_c 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q35_c q35_c

label define q35_d 3 "Has done / Do" 2 "Has done somewhat / Do somewhat" 1 "Has not done / Do not do" 
label values q35_d q35_d


// Recoding comment and plan variables


label variable q36_1 "Limited resources"
label variable q36_2 "Inadequate training"
label variable q36_3 "Limitations in the use of technology (none or little)" 
label variable q36_4 "Outdated curriculum"
label variable q36_5 "Lack of industry exposure"
label variable q36_6 "Low skill level of teachers" 
label variable q36_other "Other"


label variable q37_1 "Attainment of higher education"
label variable q37_2 "Jobs"
label variable q37_3 "Entrepreneurship"
label variable q37_other "Other"




//recoding for Soft skill variables as it was kinda test  

gen q38_new = . 
replace q38_new  = 1 if q38 == 4
replace q38_new  = 2 if q38 == 1
replace q38_new  = 3 if q38 == 2
replace q38_new  = 4 if q38 == 3
label define q38_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q38_new q38_new
tab q38_new



gen q39_new = . 
replace q39_new  = 1 if q39 == 1
replace q39_new  = 2 if q39 == 2
replace q39_new  = 3 if q39 == 4
replace q39_new  = 4 if q39 == 3
label define q39_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q39_new q39_new
tab q39_new


gen q40_new = . 
replace q40_new  = 1 if q40 == 3
replace q40_new  = 2 if q40 == 4
replace q40_new  = 3 if q40 == 1
replace q40_new  = 4 if q40 == 2
label define q40_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q40_new q40_new
tab q40_new


gen q41_new = . 
replace q41_new  = 1 if q41 == 2
replace q41_new  = 2 if q41 == 1
replace q41_new  = 3 if q41 == 3
replace q41_new  = 4 if q41 == 4
label define q41_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q41_new q41_new
tab q41_new


gen q42_new = . 
replace q42_new  = 1 if q42 == 1
replace q42_new  = 2 if q42 == 2
replace q42_new  = 3 if q42 == 4
replace q42_new  = 4 if q42 == 3
label define q42_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q42_new q42_new
tab q42_new


gen q43_new = . 
replace q43_new  = 1 if q43 == 2
replace q43_new  = 2 if q43 == 1
replace q43_new  = 3 if q43 == 3
replace q43_new  = 4 if q43 == 4
label define q43_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q43_new q43_new
tab q43_new




gen q44_new = . 
replace q44_new  = 1 if q44 == 1
replace q44_new  = 2 if q44 == 4
replace q44_new  = 3 if q44 == 2
replace q44_new  = 4 if q44 == 3
label define q44_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q44_new q44_new
tab q44_new


gen q45_new = . 
replace q45_new  = 1 if q45 == 3
replace q45_new  = 2 if q45 == 1
replace q45_new  = 3 if q45 == 2
replace q45_new  = 4 if q45 == 4
label define q45_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q45_new q45_new
tab q45_new



gen q46_new = . 
replace q46_new  = 1 if q46 == 2
replace q46_new  = 2 if q46 == 3
replace q46_new  = 3 if q46 == 4
replace q46_new  = 4 if q46 == 1
label define q46_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q46_new q46_new
tab q46_new




gen q47_new = . 
replace q47_new  = 1 if q47 == 1
replace q47_new  = 2 if q47 == 3
replace q47_new  = 3 if q47 == 2
replace q47_new  = 4 if q47 == 4
label define q47_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q47_new q47_new
tab q47_new



gen q48_new = . 
replace q48_new  = 1 if q48 == 2
replace q48_new  = 2 if q48 == 3
replace q48_new  = 3 if q48 == 1
replace q48_new  = 4 if q48 == 4
label define q48_new 1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q48_new q48_new
tab q48_new



gen q49_new = . 
replace q49_new  = 1 if q49 == 1
replace q49_new  = 2 if q49 == 3
replace q49_new  = 3 if q49 == 4
replace q49_new  = 4 if q49 == 2
label define q49_new  1 "Low Proficiency (Basic)" 2 "Limited Proficiency (Developing)" 3 "Good Proficiency (Proficient)" 4 "High Proficiency (Advanced)" 
label values q49_new q49_new
tab q49_new






/********************************************************
*					Cross tables						*
********************************************************/

ssc install asdoc
ssc install mrtab, replace 


// Demografic Info

asdoc tab q3_type, title(q3_type frequency) save(student's_result)
asdoc tab q4, title(q4 frequency) save(student's_result) append
asdoc tab q6_cat, title(q6_cat frequency) save(student's_result) append
asdoc tab q7, title(q7 frequency) save(student's_result) append
asdoc tab q8_cat, title(q8_cat frequency) save(student's_result) append



// Device owned
asdoc mrtab q9_1 q9_2 q9_3 q9_4 q9_other , by(q3_type) column title(q9_1 to q9_other X q3_type) save(student's_result) append
asdoc mrtab q9_1 q9_2 q9_3 q9_4 q9_other , by(q7) column title(q9_1 to q9_other X q7) save(student's_result) append

asdoc mrtab q10_1 q10_2 q10_3 q10_4 q10_5 q10_6 q10_7 q10_other , by(q3_type) column title(q10_1 to q10_other X q3_type)  save(student's_result) append
asdoc mrtab q10_1 q10_2 q10_3 q10_4 q10_5 q10_6 q10_7 q10_other , by(q7) column title(q10_1 to q10_other X q7) save(student's_result) append

asdoc tab q11 q3_type , column title(q11 X q3_type) save(student's_result) append 
asdoc tab q11 q7 , column title(q11 X q7) save(student's_result) append



// Readiness of the respondents (students)
asdoc tab q12 q3_type , column title(q12 X q3_type) save(student's_result) append
asdoc tab q13 q3_type , column title(q13 X q3_type) save(student's_result) append 
asdoc tab q14 q3_type , column title(q14 X q3_type) save(student's_result) append
asdoc tab q15 q3_type , column title(q15 X q3_type) save(student's_result) append 
asdoc tab q16 q3_type , column title(q16 X q3_type) save(student's_result) append 

asdoc tab q19 q3_type , column title(q19 X q3_type) save(student's_result) append 
asdoc tab q20 q3_type , column title(q20 X q3_type) save(student's_result) append 
asdoc tab q21 q3_type , column title(q21 X q3_type) save(student's_result) append

asdoc mrtab q22_1 q22_2 q22_3 q22_4 q22_other , by(q3_type) column title(q22_1 to q22_other X q3_type) save(student's_result) append 
asdoc tab q23 q3_type , column title(q23 X q3_type) save(student's_result) append
asdoc tab q24 q3_type , column title(q24 X q3_type) save(student's_result) append
asdoc tab q25 q3_type , column title(q25 X q3_type) save(student's_result) append 
asdoc tab q26 q3_type , column title(q26 X q3_type) save(student's_result) append 

/// Comment, challenge & future plan 
asdoc mrtab q36_1 q36_2 q36_3 q36_4 q36_5 q36_6 q36_other, by(q3_type) column title(q36_1 to q36_other X q3_type) save(student's_result) append
asdoc mrtab q37_1 q37_2 q37_3 q37_other, by(q3_type) column title(q37_1 to q37_other X q3_type) save(student's_result) append

///Soft skill
 
asdoc tabstat q38_new q39_new q40_new q41_new q42_new q43_new q44_new q45_new q46_new q47_new q48_new q49_new, statistics(mean median sd) by(q3_type) title(q38_new to q49_new X q3_type) save(student's_result) append


foreach var in q38_new q39_new q40_new q41_new q42_new q43_new q44_new q45_new q46_new q47_new q48_new q49_new {
    asdoc tab `var' q3_type, column title(`var' X q3_type) save(student's_result) append
}



asdoc tab q38_new q3_type, column title(q38_new X q3_type) save(student's_result) append
asdoc tab q39_new q3_type, column title(q39_new X q3_type) save(student's_result) append
asdoc tab q40_new q3_type, column title(q40_new X q3_type) save(student's_result) append
asdoc tab q41_new q3_type, column title(q41_new X q3_type) save(student's_result) append
asdoc tab q42_new q3_type, column title(q42_new X q3_type) save(student's_result) append
asdoc tab q43_new q3_type, column title(q43_new X q3_type) save(student's_result) append
asdoc tab q44_new q3_type, column title(q44_new X q3_type) save(student's_result) append
asdoc tab q45_new q3_type, column title(q45_new X q3_type) save(student's_result) append
asdoc tab q46_new q3_type, column title(q46_new X q3_type) save(student's_result) append
asdoc tab q47_new q3_type, column title(q47_new X q3_type) save(student's_result) append
asdoc tab q48_new q3_type, column title(q48_new X q3_type) save(student's_result) append
asdoc tab q49_new q3_type, column title(q49_new X q3_type) save(student's_result) append



 
/// Technical skill 
 
asdoc tabstat q17_a q17_b q17_c q17_d q17_e_numeric q17_f q17_g q17_h q17_i, statistics(mean median sd) by(q3_type) title(q17_a to q17_i X q3_type) save(student's_result) append 


///Digital literacy 
asdoc tabstat q18_a q18_b q18_c q18_j q18_k q18_l q18_m q18_n q18_o q18_p, statistics(mean median sd) by(q3_type) title(q18_a to q18_p X q3_type) save(student's_result) append
 

/// Teacher's readiness
asdoc tab q21 q3_type, column title(q21 X q3_type) save(student's_result) append
asdoc mrtab q22_1 q22_2 q22_3 q22_4 q22_other, by(q3_type) column title(q22_1 to q22_other X q3_type) save(student's_result) append
asdoc tab q23 q3_type, column title(q23 X q3_type) save(student's_result) append
asdoc tab q24 q3_type, column title(q24 X q3_type) save(student's_result) append
asdoc tab q25 q3_type, column title(q25 X q3_type) save(student's_result) append
asdoc tab q26 q3_type, column title(q26 X q3_type) save(student's_result) append
 

/// Institution's readiness 
asdoc tab q27 q3_type, column title(q27 X q3_type) save(student's_result) append
asdoc tab q28 q3_type, column title(q28 X q3_type) save(student's_result) append
asdoc tab q29 q3_type, column title(q29 X q3_type) save(student's_result) append
asdoc tab q30 q3_type, column title(q30 X q3_type) save(student's_result) append
asdoc tab q31 q3_type, column title(q31 X q3_type) save(student's_result) append
asdoc tab q32 q3_type, column title(q32 X q3_type) save(student's_result) append

asdoc tab q33_a q3_type, column title(q33_a X q3_type) save(student's_result) append
asdoc tab q33_b q3_type, column title(q33_b X q3_type) save(student's_result) append
asdoc tab q33_c q3_type, column title(q33_c X q3_type) save(student's_result) append
asdoc tab q34_a q3_type, column title(q34_a X q3_type) save(student's_result) append
asdoc tab q34_b q3_type, column title(q34_b X q3_type) save(student's_result) append
asdoc tab q34_c q3_type, column title(q34_c X q3_type) save(student's_result) append
asdoc tab q34_d q3_type, column title(q34_d X q3_type) save(student's_result) append
asdoc tab q34_e q3_type, column title(q34_e X q3_type) save(student's_result) append
asdoc tab q35_a q3_type, column title(q35_a X q3_type) save(student's_result) append
asdoc tab q35_b q3_type, column title(q35_b X q3_type) save(student's_result) append
asdoc tab q35_c q3_type, column title(q35_c X q3_type) save(student's_result) append
asdoc tab q35_d q3_type, column title(q35_d X q3_type) save(student's_result) append
 

save "clean_data_student_31082025.dta", replace

export excel "clean_data_student_31082025.xlsx" , firstrow(variables) replace
```

