---
layout: post
title: "Independent Verification of Enterprise Based Training (EBT)"
date: 2026-09-20 11:46:00
description: Stata codes for data cleaning and demographic information
tags: stata code
categories: sample-posts
thumbnail: assets/img/9.jpg
---

## For Trainee & Graduate Survey Data

```stata

cls
clear all
cd "D:\Munem\All Projects\2026\DTE\15. Data Set\Clean"

*-----------------------------------------------------------------------------
* SECTION 0: IMPORT Excel File
*-----------------------------------------------------------------------------

import excel using "Trainee_Merged", sheet("data") firstrow clear



*-----------------------------------------------------------------------------
* SECTION 1: FULL RENAME (v1-v262)
*-----------------------------------------------------------------------------


rename NameofEnumeraotor            enum_name
rename Whatisyourcurrentstatusint   resp_type_screen
rename Fullnameoftherespondent      resp_name
rename Gender                       q2_gender
rename Gender_Other                 q2_gender_other
rename RespondentStatus             q3_status
rename Disabilitytype               q3_disability_type
rename Weretrainingvenuesandclas    q4_access
rename Couldyouaccesswashroomfac    q5_toilet_access
rename Ethnicity                    q6_ethnicity
rename Ethnicity_Other              q6_ethnicity_other
rename Age                          q7_age
rename MaritalStatus                q8_marital
rename EmailAddress                 q9_email
rename ContactNumber                q10_phone
rename CurrentaddressUpazilaD       q11_address
rename Division                     division
rename Howdidyoucometoknowabo       q12_howheard_raw
rename S                            q12_c01_newspaper
rename T                            q12_c02_govtsite
rename U                            q12_c03_socialmedia
rename V                            q12_c04_workshop
rename W                            q12_c05_employer
rename X                            q12_c06_wordofmouth
rename Y                            q12_c07_cso
rename Z                            q12_c08_poster
rename AA                           q12_c09_sms
rename AB                           q12_c10_other
rename AC                           q12_other_text
rename Whydidyouchoosethistrai      q13_whychose_raw
rename AE                           q13_c01_getjob
rename AF                           q13_c02_techskills
rename AG                           q13_c03_careerprog
rename AH                           q13_c04_newtech
rename AI                           q13_c05_family
rename AJ                           q13_c06_certificate
rename AK                           q13_c07_other
rename AL                           q13_other_text
rename AM                           q2_1_progtype
rename AN                           q2_2_trade
rename Nameofthesector              sector_name
rename AP                           q2_3_level
rename AQ                           q2_5_1_admitcycle
rename AR                           q2_5_2_admitmonth
rename AS                           q2_5_3_admityear
rename AT                           q2_6_status
rename AU                           q2_7_completiondate
rename A1Didyougoforanindustry      q2a_1_attach
rename A21Nameoftheindustryatta     q2a_2_1_attachname
rename A22Addressoftheindustrya     q2a_2_2_attachaddr
rename A3Howlongdidyoustayatth      q2a_3_attachdur
rename A4Didyouusealogbookdurin     q2a_4_logbook
rename A5Howoftendidyouupdateth     q2a_5_logfreq
rename B11Noofclasseseachweek       q2b_1_1_classwk_org
rename BC                           q2b_1_1_classwk_resp
rename B12Noofweeksclassesheld      q2b_1_2_weeksheld_org
rename BE                           q2b_1_2_weeksheld_resp
rename B13Totalnoofclassesheld      q2b_1_3_totalclass_org
rename BG                           q2b_1_3_totalclass_resp
rename B14Averagenoofhoursper       q2b_1_4_hrsperclass_org
rename BI                           q2b_1_4_hrsperclass_resp
rename B15TotalnoofhoursheldO       q2b_1_5_totalhrs_org
rename B15TotalnoofhoursheldA       q2b_1_5_totalhrs_resp
rename B2Didyoureceivethehardco     q2b_2_cblm
rename B3Howmanyuniqueinstructors   q2b_3_ninstructors
rename B4Wasthereoneinstructorfr    q2b_4_dualinstructor
rename PleaseSpecifyfromwherefacto  q2b_4_specify
rename B5Wholedthemajorityofins     q2b_5_wholed
rename Other13                      q2b_5_other_text
rename B6Howwerelessonsprimarily    q2b_6_delivery
rename BS                           hdr_3_1
rename BT                           q3_1a_ppe_trainee
rename BU                           q3_1b_ppe_enum
rename BV                           q3_2_ppe_enforced
rename BW                           q3_3_safetyguard
rename BX                           q3_4_riskbriefed
rename BY                           q4_1_water
rename BZ                           q4_2_toilets
rename CA                           q4_3_complaintinfo
rename CB                           q4_4_complaintbox
rename CC                           q5_1_stipend
rename CD                           q5_2_stipendamt
rename CE                           q5_3_stipendmode
rename CF                           q5_4_stipendtiming_raw
rename CG                           q5_4_c01_start
rename CH                           q5_4_c02_end
rename CI                           q5_4_c03_2wkafter
rename CJ                           q5_5_stipenduse_raw
rename CK                           q5_5_c01_food
rename CL                           q5_5_c02_transport
rename CM                           q5_5_c03_mobiledata
rename CN                           q5_5_c04_household
rename CO                           q5_5_c05_savings
rename CP                           q5_5_c06_other
rename CQ                           q5_5_other_text
rename CR                           hdr_5_6
rename CS                           hdr_5_6_1
rename CT                           q5_6_1a_fee
rename CU                           q5_6_1b_feefunder
rename CV                           q5_6_1b_other_text
rename CW                           q5_6_1c_feepartial
rename CX                           q5_6_1d_feeother
rename CY                           hdr_5_6_2
rename CZ                           q5_6_2a_transport
rename DA                           q5_6_2b_transportfunder
rename DB                           q5_6_2b_other_text
rename DC                           q5_6_2c_transportpartial
rename DD                           q5_6_2d_transportother
rename DE                           hdr_5_6_3
rename DF                           q5_6_3a_equip
rename DG                           q5_6_3b_equipfunder
rename DH                           q5_6_3b_other_text
rename DI                           q5_6_3c_equippartial
rename DJ                           q5_6_3d_equipother
rename DK                           hdr_5_6_4
rename DL                           q5_6_4a_stationery
rename DM                           q5_6_4b_stationeryfunder
rename DN                           q5_6_4b_other_text
rename DO                           q5_6_4c_stationerypartial
rename DP                           q5_6_4d_stationeryother
rename DQ                           q5_6_5a_othercomponent1
rename DR                           q5_6_5b_othertotal1
rename DS                           q5_6_5c_otherfunder1
rename DT                           q5_6_5c_other_text
rename DU                           q5_6_5d_otherpartial1
rename DV                           q5_6_5e_otherother1
rename DW                           q5_6_6a_othercomponent2
rename DX                           q5_6_6b_othertotal2
rename DY                           q5_6_6c_otherfunder2
rename DZ                           q5_6_6c_other_text
rename EA                           q5_6_6d_otherpartial2
rename EB                           q5_6_6e_otherother2
rename EC                           q6_1_attendrate
rename ED                           q6_2_lowattend_raw
rename EE                           q6_2_c01_work
rename EF                           q6_2_c02_family
rename EG                           q6_2_c03_transport
rename EH                           q6_2_c04_illness
rename EI                           q6_2_c05_cancelled
rename EJ                           q6_2_c06_na
rename EK                           q6_2_c07_other
rename EL                           q6_2_other_text
rename EM                           hdr_6_3
rename EN                           q6_3_1_scheduling
rename EO                           q6_3_2_matavail
rename EP                           q6_3_3_matquality
rename EQ                           q6_3_4_overallquality
rename ER                           q6_3_5_trainerquality
rename ES                           q6_3_6_industryfacility
rename ET                           q7_1_otherenrolled
rename EU                           q7_2_progtype
rename EV                           q7_2_educ_specify
rename EW                           q7_2_vocational_specify
rename EX                           q7_2_other_specify
rename EY                           q7_3_aftertraining
rename EZ                           q7_3_other_specify
rename FA                           q7_4_overallsat
rename FB                           q7_5_recommend
rename A1Whatisthehighestlevelo     q8a_1_edulevel
rename A2Inwhichyeardidyoucompl     q8a_2_eduyear
rename A3Haveyoucompletedanyvoca    q8a_3_priortraining
rename A4Durationofthatprevioust    q8a_4_priortraindur
rename A5Wasthattrainingrelatedt    q8a_5_priortrainrelated
rename A6Whatwasyouremployment      q8a_6_prejobstatus
rename FI                           q8a_6_other_text
rename A7Whattypeofworkwereyou      q8a_7_worktype
rename FK                           q8a_7_other_text
rename A8Inwhichsectorwasyourwo     q8a_8_worksector
rename FM                           q8a_8_other_text
rename A9Wasthatworkrelatedtoth     q8a_9_workrelated
rename A10Whatwasyouraveragemont    q8a_10_preincome
rename A11Howlonghadyoubeenunem     q8a_11_unemploydur
rename A12Whatwasyourprimaryreas    q8a_12_joinreason
rename FR                           q8a_12_other_text
rename A13Atthetimeofjoiningwe      q8a_13_primaryearner
rename B1Inthelast7sevendays        q8b_1_worked7d
rename B2Inthelast7daysdidyou       q8b_2_absentjob7d
rename B3Haveyouworkedatleast1      q8b_3_ownuse7d
rename B4Inthelast7sevendays        q8b_4_absentownuse7d
rename B51aHowmanymonthsdidyou      q8b_5_1a_monthswork12m
rename B51bOnaveragehowmanyday      q8b_5_1b_daysmonth12m
rename B51cOnaveragehowmanyhou      q8b_5_1c_hrsday12m
rename B52aHowmanymonthsdidyou      q8b_5_2a_monthsownuse12m
rename B52bOnaveragehowmanyday      q8b_5_2b_daysmonthownuse12m
rename B52cOnaveragehowmanyhou      q8b_5_2c_hrsdayownuse12m
rename B61aHowmanydaysdidyouwo      q8b_6_1a_days_pre
rename B61bOnaveragehowmanyhou      q8b_6_1b_hrs_pre
rename B61cHowmuchdidyouearnin      q8b_6_1c_income_pre
rename B62aHowmanydaysdidyouwo      q8b_6_2a_days_m1
rename B62bOnaveragehowmanyhou      q8b_6_2b_hrs_m1
rename B62cHowmuchdidyouearnin      q8b_6_2c_income_m1
rename B63aHowmanydaysdidyouwo      q8b_6_3a_days_m2
rename B63bOnaveragehowmanyhou      q8b_6_3b_hrs_m2
rename B63cHowmuchdidyouearnin      q8b_6_3c_income_m2
rename B64aHowmanydaysdidyouwo      q8b_6_4a_days_m3
rename B64bOnaveragehowmanyhou      q8b_6_4b_hrs_m3
rename B64cHowmuchdidyouearnin      q8b_6_4c_income_m3
rename B65aHowmanydaysdidyouwo      q8b_6_5a_days_m4
rename B65bOnaveragehowmanyhou      q8b_6_5b_hrs_m4
rename B65cHowmuchdidyouearnin      q8b_6_5c_income_m4
rename B66aHowmanydaysdidyouwo      q8b_6_6a_days_m5
rename B66bOnaveragehowmanyhou      q8b_6_6b_hrs_m5
rename B66cHowmuchdidyouearnin      q8b_6_6c_income_m5
rename B67aHowmanydaysdidyouwo      q8b_6_7a_days_m6
rename B67bOnaveragehowmanyhou      q8b_6_7b_hrs_m6
rename B67cHowmuchdidyouearnin      q8b_6_7c_income_m6
rename C1Whichofthefollowingwoul    q8c_1_aspiration
rename C2Whatistheprimaryreason     q8c_2_furtherreason
rename HA                           q8c_2_other_text
rename A1Didyouappearinaformal      q9a_1_assessed
rename A2Pleasementiontheassessme   q9a_2_assessdate
rename A3WereyouassessedasCompet    q9a_3_competent
rename A4Pleasespecifythemonthan    q9a_4_graddate
rename A5Howmanyattemptsdiditta     q9a_5_attempts
rename A6Whyhaveyounotappearedi     q9a_6_notappeared_raw
rename HH                           q9a_6_c01_notscheduled
rename HI                           q9a_6_c02_didntknow
rename HJ                           q9a_6_c03_healthpersonal
rename HK                           q9a_6_c04_other
rename HL                           q9a_6_other_text
rename A7Howlongaftercompletiono    q9a_7_awaitingdur
rename HN                           q9a_7_other_text
rename A81NameoftheAssessmentCe     q9a_8_1_rtoname
rename A82AddressoftheAssessment    q9a_8_2_rtoaddr
rename HQ                           q9_1_empstatus
rename B1Howlongaftercompletingt    q9b_1_timetojob
rename HS                           q9b_1_other_text
rename B2Isyourcurrentjobinthe      q9b_2_sectormatch
rename B3IfdifferentsectorWhya      q9b_3_whynotsector
rename HV                           q9b_3_other_text
rename B4Whatisyourcurrentemploy    q9b_4_emptype
rename HX                           q9b_4_other_text
rename B5Whereisyourcurrentjobl     q9b_5_joblocation
rename B6Howdidyouobtainthisjob     q9b_6_jobsource
rename IA                           q9b_6_other_text
rename B7Whatisyourcurrentaverag    q9b_7_currentincome
rename B8Hasyourincomechangedcom    q9b_8_incomechange
rename B9DoyoufeelEBTtrainingdi     q9b_9_ebtcontributed
rename C1Howlonghaveyoubeenlook     q9c_1_searchdur
rename C2Whatarethemainbarriers     q9c_2_barriers_raw
rename IG                           q9c_2_c01_novacancy
rename IH                           q9c_2_c02_lowsalary
rename II                           q9c_2_c03_location
rename IJ                           q9c_2_c04_noexperience
rename IK                           q9c_2_c05_nocert
rename IL                           q9c_2_c06_family
rename IM                           q9c_2_c07_other
rename IN                           q9c_2_other_text
rename D1Whyareyounotcurrentlyl     q9d_1_reasons_raw
rename IP                           q9d_1_c01_study
rename IQ                           q9d_1_c02_family
rename IR                           q9d_1_c03_health
rename IS                           q9d_1_c04_ownbusiness
rename IT                           q9d_1_c05_overseas
rename IU                           q9d_1_c06_other
rename IV                           q9d_1_other_text
rename D2Doyouintendtolookforw      q9d_2_planlook
rename PleasecollectyourGPSlocation hdr_gps
rename _PleasecollectyourGPSlocatio gps_lat
rename IZ                           gps_lon
rename JA                           gps_alt
rename JB                           gps_precision

* Quick sanity check: confirm all 262 renamed and nothing left unmatched
quietly ds
local nvars : word count `r(varlist)'
di as result "Variables after rename: `nvars'  (expected: 262)"

*==============================================================================
* SECTION 2: QUALITY CHECKS
*==============================================================================


duplicates tag resp_name q10_phone, gen(dup_flag)

tab dup_flag    //0= Not duplicate; 1= Duplicate

tab enum_name, m
gen gps_flag = (gps_lat<20.5 | gps_lat>26.7 | gps_lon<88 | gps_lon>92.7) if !missing(gps_lat)
tab gps_flag, m
gen age_flag = (q7_age<14 | q7_age>65) if !missing(q7_age)
tab age_flag, m



*==============================================================================
* SECTION 3: NEW VARIABLES GENERATED FROM STRING VARIABLES
*==============================================================================

*** All "_new" valriables are numeric, recoded from their string couterpass



**q2_gender
gen q2_gender_new =.
replace q2_gender_new = 1 if q2_gender == "Male"
replace q2_gender_new = 2 if q2_gender == "Female"
replace q2_gender_new = 3 if q2_gender == "Prefer not to say"
replace q2_gender_new = 88 if q2_gender == "Other"
label define GENDER 1 "Male" 2 "Female" 3 "Prefer not to say" 88 "Other"
label values q2_gender_new GENDER
label variable q2_gender_new "Gender of Respondent"
tab q2_gender_new



**q3_status
gen q3_status_new =.
replace q3_status_new = 1 if q3_status == "Fully abled"
replace q3_status_new = 2 if q3_status == "Disabled"
label define STATUS3 1 "Fully able" 2 "Disabled"
label values q3_status_new STATUS3
label variable q3_status_new "Respondent Status (Able/PWD)"
tab q3_status_new


*---------YES/NO-------------------------------------------------------------------------------------------------------------------------


label define YESNO 0 "No" 1 "Yes"

foreach var in q4_access q5_toilet_access q2a_4_logbook q2b_2_cblm ///
	q3_1a_ppe_trainee q3_1b_ppe_enum q3_2_ppe_enforced ///
	q3_3_safetyguard q3_4_riskbriefed ///
    q4_1_water q4_2_toilets q4_3_complaintinfo q4_4_complaintbox ///
    q5_1_stipend q7_1_otherenrolled q8a_3_priortraining ///
	q8a_5_priortrainrelated q8a_9_workrelated q8b_1_worked7d ///
    q8b_2_absentjob7d q8b_3_ownuse7d q8b_4_absentownuse7d {

    gen `var'_new = .
    replace `var'_new = 1 if `var' == "Yes"
    replace `var'_new = 0 if `var' == "No"
    label values `var'_new YESNO
}


**Doesn't have any data now. It is a graduate specific question (YES/NO)
*gen q9a_1_assessed_new =.
*gen q9a_3_competent_new =.
*gen q9b_2_sectormatch_new =.


label variable q3_status_new "Respondent Status (Able/PWD)"



label variable q4_access_new "q4_access (numeric)"
label variable q5_toilet_access_new "q5_toilet_access (numeric)"
label variable q2a_4_logbook_new "q2a_4_logbook (numeric)"
label variable q2b_2_cblm_new "q2b_2_cblm (numeric)"
label variable q3_1a_ppe_trainee_new "q3_1a_ppe_trainee  (numeric)"
label variable q3_1b_ppe_enum_new "q3_1b_ppe_enum (numeric)"
label variable q3_2_ppe_enforced_new "q3_2_ppe_enforced (numeric)"
label variable q3_3_safetyguard_new "q3_3_safetyguard (numeric)"
label variable q3_4_riskbriefed_new  "q3_4_riskbriefed (numeric)"
label variable q4_1_water_new "q4_1_water (numeric)"
label variable q4_2_toilets_new "q4_2_toilets (numeric)"
label variable q4_3_complaintinfo_new "q4_3_complaintinfo (numeric)"
label variable q4_4_complaintbox_new "q4_4_complaintbox (numeric)"
label variable q5_1_stipend_new "q5_1_stipend (numeric)"
label variable q7_1_otherenrolled_new "q7_1_otherenrolled (numeric)"
label variable q8a_3_priortraining_new "q8a_3_priortraining (numeric)"
label variable q8a_5_priortrainrelated_new "q8a_5_priortrainrelated (numeric)"
label variable q8a_9_workrelated_new "q8a_9_workrelated (numeric)"
label variable q8b_1_worked7d_new "q8b_1_worked7d (numeric)"
label variable q8b_2_absentjob7d_new "q8b_2_absentjob7d (numeric)"
label variable q8b_3_ownuse7d_new "q8b_3_ownuse7d (numeric)"
label variable q8b_4_absentownuse7d_new "q8b_4_absentownuse7d (numeric)"



**q6_ethnicity
gen q6_ethnicity_new =.
replace q6_ethnicity_new = 1 if q6_ethnicity == "Bengali"
replace q6_ethnicity_new = 2 if q6_ethnicity == "Chakma"
replace q6_ethnicity_new = 3 if q6_ethnicity == "Marma"
replace q6_ethnicity_new = 4 if q6_ethnicity == "Garo"
replace q6_ethnicity_new = 5 if q6_ethnicity == "Khassia"
replace q6_ethnicity_new = 6 if q6_ethnicity == "Shaotal"
replace q6_ethnicity_new = 88 if q6_ethnicity == "Other"

label define ETHNIC 1 "Bangali" 2 "Chakma" 3 "Marma" 4 "Garo" ///
                     5 "Khassia" 6 "Shaotal" 88 "Other"
label values q6_ethnicity_new ETHNIC
label variable q6_ethnicity_new "q6_ethnicity (numeric)"
tab q6_ethnicity_new


**q8_marital
gen q8_marital_new =.
replace q8_marital_new = 1 if q8_marital == "Married"
replace q8_marital_new = 2 if q8_marital == "Unmarried"
replace q8_marital_new = 3 if q8_marital == "Divorced"
replace q8_marital_new = 4 if q8_marital == "Separated"
replace q8_marital_new = 5 if q8_marital == "Widowed"

label define MARITAL 1 "Married" 2 "Unmarried" 3 "Divorced" ///
                      4 "Separated" 5 "Widowed"
label values q8_marital_new MARITAL
label variable q8_marital_new "q8_marital (numeric)"
tab q8_marital_new



**division

gen division_new =.
replace division_new = 1 if division == "Barisal"
replace division_new = 2 if division == "Chattogram"
replace division_new = 3 if division == "Dhaka"
replace division_new = 4 if division == "Khulna"
replace division_new = 5 if division == "Mymensingh"
replace division_new = 6 if division == "Rajshahi"
replace division_new = 7 if division == "Rangpur"
replace division_new = 8 if division == "Sylhet"

label define DIVISION 1 "Barisal" 2 "Chattogram" 3 "Dhaka" 4 "Khulna" ///
						5 "Mymensingh" 6 "Rajshahi" 7 "Rangpur" 8 "Sylhet"
label values division_new DIVISION
label variable division_new "division (numeric)"
tab division_new



**q2_1_progtype

gen q2_1_progtype_new =.
replace q2_1_progtype_new = 1 if q2_1_progtype == "Apprenticeship and Dual Training Program (ADTP)"
replace q2_1_progtype_new = 2 if q2_1_progtype == "Technical Skill Training Program (TSTP)"
replace q2_1_progtype_new = 3 if q2_1_progtype == "Skill up-gradation Training Program (SUTP)"
replace q2_1_progtype_new = 4 if q2_1_progtype == "IHTP"
label define PROGTYPE 1 "ADTP" 2 "TSTP" 3 "SUTP" 4 "IHTP"
label values q2_1_progtype_new PROGTYPE
label variable q2_1_progtype_new "q2_1_progtype (numeric)"
tab q2_1_progtype_new


**sector_name
tab  sector_name
gen sector_name_new =.
replace sector_name_new = 1 if sector_name =="Agro-Food Sector" 
replace sector_name_new = 2 if sector_name == "Ceramic" 
replace sector_name_new = 3 if sector_name == "Construction & Civil"
replace sector_name_new = 4 if sector_name == "Furniture"  
replace sector_name_new = 5 if sector_name == "Garment, Textile & Leather"
replace sector_name_new = 6 if sector_name == "Health Sector"
replace sector_name_new = 7 if sector_name == "ICT" 
replace sector_name_new = 8 if sector_name == "Informal"
replace sector_name_new = 9 if sector_name == "Light Engineering & Manufacturing" 
replace sector_name_new = 10 if sector_name == "Tourism and Hospitality" 


label define SECTOR 1 "Agro-Food Sector" 2 "Ceramic"  ///
					3 "Construction & Civil" 4 "Furniture"  ///
					5 "Garment, Textile & Leather" 6 "Health Sector" ///
					7 "ICT" 8 "Informal" ///
					9 "Light Engineering & Manufacturing" ///
					10 "Tourism and Hospitality" 
label values sector_name_new SECTOR
label variable sector_name_new "sector_name (numeric)"
tab sector_name_new



**q2_3_level
gen q2_3_level_new =.
replace q2_3_level_new = 1 if q2_3_level == "Level 1"
replace q2_3_level_new = 2 if q2_3_level == "Level 2"
replace q2_3_level_new = 3 if q2_3_level == "Level 3"
replace q2_3_level_new = 4 if q2_3_level == "Level 4"
label define LEVEL 1 "Level 1" 2 "Level 2" 3 "Level 3" 4 "Level 4"
label values q2_3_level_new LEVEL
label variable q2_1_progtype_new "q2_3_level (numeric)"
tab q2_3_level_new


**q2_5_1_admitcycle
gen q2_5_1_admitcycle_new =.
replace q2_5_1_admitcycle_new = 1 if q2_5_1_admitcycle == 1
replace q2_5_1_admitcycle_new = 2 if q2_5_1_admitcycle == 2
replace q2_5_1_admitcycle_new = 3 if q2_5_1_admitcycle == 3
replace q2_5_1_admitcycle_new = 4 if q2_5_1_admitcycle == 4
label define ADMITCYCLE 1 "1st Cycle" 2 "2nd cycle" 3 "3rd cycle" 4 "4th cycle"
label values q2_5_1_admitcycle_new ADMITCYCLE
label variable q2_5_1_admitcycle_new "q2_5_1_admitcycle (numeric)"
tab q2_5_1_admitcycle_new


**q2_5_2_admitmonth
gen q2_5_2_admitmonth_new =.
replace q2_5_2_admitmonth_new = 1 if q2_5_2_admitmonth == "January"
replace q2_5_2_admitmonth_new = 2 if q2_5_2_admitmonth == "March"
replace q2_5_2_admitmonth_new = 3 if q2_5_2_admitmonth == "April"
replace q2_5_2_admitmonth_new = 4 if q2_5_2_admitmonth == "May"
replace q2_5_2_admitmonth_new = 4 if q2_5_2_admitmonth == "May "
replace q2_5_2_admitmonth_new = 5 if q2_5_2_admitmonth == "June"
replace q2_5_2_admitmonth_new = 5 if q2_5_2_admitmonth == "June "

label define ADMITMONTH 1 "January" 2 "March" 3 "April" 4 "May" 5 "June"
label values q2_5_2_admitmonth_new ADMITMONTH
label variable q2_5_2_admitmonth_new "q2_5_2_admitmonth (numeric)"
tab q2_5_2_admitmonth_new


**q2_6_status

gen q2_6_status_new =.
replace q2_6_status_new = 1 if q2_6_status =="Still undergoing training (TRAINEE)"
replace q2_6_status_new = 2 if q2_6_status =="Completed training (GRADUATE)"
label define STATUSq2_6 1 "Trainee" 2 "Graduate"
label values q2_6_status_new STATUSq2_6
label variable q2_6_status_new "q2_6_status (numeric)"
tab q2_6_status_new


**q2a_1_attach 
gen q2a_1_attach_new =.
replace q2a_1_attach_new = 1 if q2a_1_attach == "Yes"
replace q2a_1_attach_new = 2 if q2a_1_attach == "No"
replace q2a_1_attach_new = 3 if q2a_1_attach == "Not eligible yet"

label define ATTACH 1 "Yes" 2 "No" 3 "Not eligible yet"
label values q2a_1_attach_new ATTACH
label variable q2a_1_attach_new "q2a_1_attach (Numeric)"
tab q2a_1_attach_new

**q2a_5_logfreq
gen q2a_5_logfreq_new =.
replace q2a_5_logfreq_new = 1 if q2a_5_logfreq =="Every day"
replace q2a_5_logfreq_new = 2 if q2a_5_logfreq =="Every three days"
replace q2a_5_logfreq_new = 3 if q2a_5_logfreq =="Every seven days"
replace q2a_5_logfreq_new = 4 if q2a_5_logfreq =="Whenever I wished"

label define LOGFREQ 1 "Every day" 2 "Every three days" ///
						3 "Every seven days" 4 "Whenever I wished"
label values q2a_5_logfreq_new LOGFREQ
label variable q2a_5_logfreq_new "q2a_5_logfreq (Numeric)"
tab q2a_5_logfreq_new



**q2b_4_dualinstructor

gen q2b_4_dualinstructor_new =.
replace q2b_4_dualinstructor_new = 1 if q2b_4_dualinstructor == "Yes, both present"
replace q2b_4_dualinstructor_new = 2 if q2b_4_dualinstructor == "No, only one instructor type"
label define DUALINSTRUCTOR 1 "Both Present" 2 "Only One Type"
label values q2b_4_dualinstructor_new DUALINSTRUCTOR
label variable q2b_4_dualinstructor_new "q2b_4_dualinstructor (numeric)"
tab q2b_4_dualinstructor_new

**q2b_4_specify
gen q2b_4_specify_new =.
replace q2b_4_specify_new = 1 if q2b_4_specify == "Factory"
replace q2b_4_specify_new = 1 if q2b_4_specify == "Factory "
replace q2b_4_specify_new = 2 if q2b_4_specify == "Institute"
label define SPECIFYq2b_4 1 "Factory" 2 "Institute"
label values q2b_4_specify_new SPECIFYq2b_4
label variable q2b_4_specify_new "q2b_4_specify (numeric)"
tab q2b_4_specify_new




**q2b_5_wholed
gen q2b_5_wholed_new =.
replace q2b_5_wholed_new = 1 if q2b_5_wholed == "Factory instructor"
replace q2b_5_wholed_new = 2 if q2b_5_wholed == "Institute instructor"
replace q2b_5_wholed_new = 3 if q2b_5_wholed == "Both equally"
replace q2b_5_wholed_new = 88 if q2b_5_wholed == "Other"
label define WHOLEDq2b_5 1 "Factory instructor" 2 "Institute instructor" ///
							3 "Both equally" 88 "Other"

label values q2b_5_wholed_new WHOLEDq2b_5
label variable q2b_5_wholed_new "q2b_5_wholed (numeric)"
tab q2b_5_wholed_new




**q2b_6_delivery

gen q2b_6_delivery_new =.
replace q2b_6_delivery_new = 1 if q2b_6_delivery == "Mostly practical / hands-on"
replace q2b_6_delivery_new = 2 if q2b_6_delivery == "Mostly theoretical"
replace q2b_6_delivery_new = 3 if q2b_6_delivery == "Equal mix"
label define DELIVERY 1 "Mostly practical / hands-on" 2 "Mostly theoretical" 3 "Equal mix"
label values q2b_6_delivery_new DELIVERY
label variable q2b_6_delivery_new "q2b_6_delivery (numeric)"
tab q2b_6_delivery_new 

**q5_3_stipendmode // no ovservation now
*gen q5_3_stipendmode_new =.
*replace q5_3_stipendmode_new = 1 if q5_3_stipendmode == "In a single tranche"
*replace q5_3_stipendmode_new = 2 if q5_3_stipendmode == "In two tranches"
*replace q5_3_stipendmode_new = 3 if q5_3_stipendmode == "In three tranches"
*replace q5_3_stipendmode_new = 4 if q5_3_stipendmode == "Don't receive yet"

*label define STIPENDMODE 1 "In a single tranche" 2 "In two tranches" 3 "In three tranches" 4 "Don't receive yet"
*label values q5_3_stipendmode_new STIPENDMODE
*label variable q5_3_stipendmode_new "q5_3_stipendmode (numeric)"
*tab q5_3_stipendmode_new




**q5_6_1b_feefunder (course fee)
gen q5_6_1b_feefunder_new =.
replace q5_6_1b_feefunder_new = 1 if q5_6_1b_feefunder == "Myself"
replace q5_6_1b_feefunder_new = 2 if q5_6_1b_feefunder == "Friends/relatives"
replace q5_6_1b_feefunder_new = 3 if q5_6_1b_feefunder == "Fully ASSET-funded"
replace q5_6_1b_feefunder_new = 4 if q5_6_1b_feefunder == "Partially ASSET-funded"
replace q5_6_1b_feefunder_new = 5 if q5_6_1b_feefunder == "Other project"
replace q5_6_1b_feefunder_new = 6 if q5_6_1b_feefunder == "Waiver"
replace q5_6_1b_feefunder_new = 88 if q5_6_1b_feefunder == "Other"

label define FEEFUNDERq5_6_1b 1 "Myself" 2 "Friends/relatives" ///
								3 "Fully ASSET-funded" ///
								4 "Partially ASSET-funded" ///
								5 "Other project" 6 "Waiver" 88 "Other"
label values q5_6_1b_feefunder_new FEEFUNDERq5_6_1b
label variable q5_6_1b_feefunder_new "q5_6_1b_feefunder (numeric)"
tab q5_6_1b_feefunder_new


**q5_6_2b_transportfunder (Conveyance)
gen q5_6_2b_transportfunder_new =.
replace q5_6_2b_transportfunder_new = 1 if q5_6_2b_transportfunder == "Myself"
replace q5_6_2b_transportfunder_new = 2 if q5_6_2b_transportfunder == "Friends/relatives"
replace q5_6_2b_transportfunder_new = 3 if q5_6_2b_transportfunder == "Fully ASSET-funded"
replace q5_6_2b_transportfunder_new = 4 if q5_6_2b_transportfunder == "Partially ASSET-funded"
replace q5_6_2b_transportfunder_new = 5 if q5_6_2b_transportfunder == "Other project"
replace q5_6_2b_transportfunder_new = 6 if q5_6_2b_transportfunder == "Waiver"
replace q5_6_2b_transportfunder_new = 88 if q5_6_2b_transportfunder == "Other"

label define FEEFUNDERq5_6_2b 1 "Myself" 2 "Friends/relatives" ///
								3 "Fully ASSET-funded" 4 "Partially ASSET-funded" ///
								5 "Other project" 6 "Waiver" 88 "Other"
label values q5_6_2b_transportfunder_new FEEFUNDERq5_6_2b
label variable q5_6_2b_transportfunder_new "q5_6_2b_transportfunder (numeric)"
tab q5_6_2b_transportfunder_new



**q5_6_3b_equipfunder (Equipment purchase)
gen q5_6_3b_equipfunder_new =.
replace q5_6_3b_equipfunder_new = 1 if q5_6_3b_equipfunder == "Myself"
replace q5_6_3b_equipfunder_new = 2 if q5_6_3b_equipfunder == "Friends/relatives"
replace q5_6_3b_equipfunder_new = 3 if q5_6_3b_equipfunder == "Fully ASSET-funded"
replace q5_6_3b_equipfunder_new = 4 if q5_6_3b_equipfunder == "Partially ASSET-funded"
replace q5_6_3b_equipfunder_new = 5 if q5_6_3b_equipfunder == "Other project"
replace q5_6_3b_equipfunder_new = 6 if q5_6_3b_equipfunder == "Waiver"
replace q5_6_3b_equipfunder_new = 88 if q5_6_3b_equipfunder == "Other"

label define FEEFUNDERq5_6_3b 1 "Myself" 2 "Friends/relatives" 3 "Fully ASSET-funded" ///
								4 "Partially ASSET-funded" 5 "Other project" ///
								6 "Waiver" 88 "Other"
label values q5_6_3b_equipfunder_new FEEFUNDERq5_6_3b
label variable q5_6_3b_equipfunder_new "q5_6_3b_equipfunder (numeric)"
tab q5_6_3b_equipfunder_new



**q5_6_4b_stationeryfunder (Stationery)
gen q5_6_4b_stationeryfunder_new =.
replace q5_6_4b_stationeryfunder_new = 1 if q5_6_4b_stationeryfunder == "Myself"
replace q5_6_4b_stationeryfunder_new = 2 if q5_6_4b_stationeryfunder == "Friends/relatives"
replace q5_6_4b_stationeryfunder_new = 3 if q5_6_4b_stationeryfunder == "Fully ASSET-funded"
replace q5_6_4b_stationeryfunder_new = 4 if q5_6_4b_stationeryfunder == "Partially ASSET-funded"
replace q5_6_4b_stationeryfunder_new = 5 if q5_6_4b_stationeryfunder == "Other project"
replace q5_6_4b_stationeryfunder_new = 6 if q5_6_4b_stationeryfunder == "Waiver"
replace q5_6_4b_stationeryfunder_new = 88 if q5_6_4b_stationeryfunder == "Other"

label define FEEFUNDERq5_6_4b 1 "Myself" 2 "Friends/relatives" 3 "Fully ASSET-funded" ///
								4 "Partially ASSET-funded" 5 "Other project" ///
								6 "Waiver" 88 "Other"
label values q5_6_4b_stationeryfunder_new FEEFUNDERq5_6_4b
label variable q5_6_4b_stationeryfunder_new "q5_6_4b_stationeryfunder (numeric)"
tab q5_6_4b_stationeryfunder_new


**q5_6_5c_otherfunder1 [Other (specify)]

gen q5_6_5c_otherfunder1_new =.
replace q5_6_5c_otherfunder1_new = 1 if q5_6_5c_otherfunder1 == "Myself"
replace q5_6_5c_otherfunder1_new = 2 if q5_6_5c_otherfunder1 == "Friends/relatives"
replace q5_6_5c_otherfunder1_new = 3 if q5_6_5c_otherfunder1 == "Fully ASSET-funded"
replace q5_6_5c_otherfunder1_new = 4 if q5_6_5c_otherfunder1 == "Partially ASSET-funded"
replace q5_6_5c_otherfunder1_new = 5 if q5_6_5c_otherfunder1 == "Other project"
replace q5_6_5c_otherfunder1_new = 6 if q5_6_5c_otherfunder1 == "Waiver"
replace q5_6_5c_otherfunder1_new = 88 if q5_6_5c_otherfunder1 == "Other"

label define FEEFUNDERq5_6_5c 1 "Myself" 2 "Friends/relatives" ///
								3 "Fully ASSET-funded" 4 "Partially ASSET-funded" ///
								5 "Other project" 6 "Waiver" 88 "Other"
label values q5_6_5c_otherfunder1_new FEEFUNDERq5_6_5c
label variable q5_6_5c_otherfunder1_new "q5_6_5c_otherfunder1 (numeric)"
tab q5_6_5c_otherfunder1_new
tab q5_6_5c_other_text


**q5_6_6c_otherfunder2 [Other (specify)]

gen q5_6_6c_otherfunder2_new =.
replace q5_6_6c_otherfunder2_new = 1 if q5_6_6c_otherfunder2 == "Myself"
replace q5_6_6c_otherfunder2_new = 2 if q5_6_6c_otherfunder2 == "Friends/relatives"
replace q5_6_6c_otherfunder2_new = 3 if q5_6_6c_otherfunder2 == "Fully ASSET-funded"
replace q5_6_6c_otherfunder2_new = 4 if q5_6_6c_otherfunder2 == "Partially ASSET-funded"
replace q5_6_6c_otherfunder2_new = 5 if q5_6_6c_otherfunder2 == "Other project"
replace q5_6_6c_otherfunder2_new = 6 if q5_6_6c_otherfunder2 == "Waiver"
replace q5_6_6c_otherfunder2_new = 88 if q5_6_6c_otherfunder2 == "Other"

label define FEEFUNDERq5_6_6c 1 "Myself" 2 "Friends/relatives" 3 "Fully ASSET-funded" ///
								4 "Partially ASSET-funded" 5 "Other project" ///
								6 "Waiver" 88 "Other"
label values q5_6_6c_otherfunder2_new FEEFUNDERq5_6_6c
label variable q5_6_6c_otherfunder2_new "q5_6_6c_otherfunder2 (numeric)"
tab q5_6_6c_otherfunder2_new
tab q5_6_6c_other_text



**q6_1_attendrate
gen q6_1_attendrate_new =.
replace q6_1_attendrate_new = 1 if q6_1_attendrate == "Almost always (>90% attendance)"
replace q6_1_attendrate_new = 2 if q6_1_attendrate == "Mostly (75–90%)"
replace q6_1_attendrate_new = 3 if q6_1_attendrate == "Sometimes (50–74%)"
replace q6_1_attendrate_new = 4 if q6_1_attendrate == "Rarely (<50%)"

label define ATTENDRATE 1 "Almost always (>90% attendance)" 2 "Mostly (75–90%)" ///
						3 "Sometimes (50–74%)" 4 "Rarely (<50%)"
label values q6_1_attendrate_new ATTENDRATE
label variable q6_1_attendrate_new "q6_1_attendrate (numeric)"
tab q6_1_attendrate_new
tab q6_1_attendrate_new q2_gender_new, row col
tab q6_1_attendrate_new division_new, row col 
tab q6_1_attendrate_new sector_name_new, row col


**q6_3_1_scheduling

gen q6_3_1_scheduling_new =.
replace q6_3_1_scheduling_new = 1 if q6_3_1_scheduling == "Very poor"
replace q6_3_1_scheduling_new = 2 if q6_3_1_scheduling == "Poor"
replace q6_3_1_scheduling_new = 3 if q6_3_1_scheduling == "Good"
replace q6_3_1_scheduling_new = 4 if q6_3_1_scheduling == "Very good"
replace q6_3_1_scheduling_new = 5 if q6_3_1_scheduling == "Excellent"

label define SCHEDULING 1 "Very poor" 2 "Poor" 3 "Good" 4 "Very good" 5 "Excellent"
label values q6_3_1_scheduling_new SCHEDULING
label variable q6_3_1_scheduling_new "q6_3_1_scheduling (numeric)"
tab q6_3_1_scheduling_new


**q6_3_2_matavail


gen q6_3_2_matavail_new =.
replace q6_3_2_matavail_new = 1 if q6_3_2_matavail == "Very poor"
replace q6_3_2_matavail_new = 2 if q6_3_2_matavail == "Poor"
replace q6_3_2_matavail_new = 3 if q6_3_2_matavail == "Good"
replace q6_3_2_matavail_new = 4 if q6_3_2_matavail == "Very good"
replace q6_3_2_matavail_new = 5 if q6_3_2_matavail == "Excellent"

label define MATAVAIL 1 "Very poor" 2 "Poor" 3 "Good" 4 "Very good" 5 "Excellent"
label values q6_3_2_matavail_new MATAVAIL
label variable q6_3_2_matavail_new "q6_3_2_matavail (numeric)"
tab q6_3_2_matavail_new


**q6_3_3_matquality

gen q6_3_3_matquality_new =.
replace q6_3_3_matquality_new = 1 if q6_3_3_matquality == "Very poor"
replace q6_3_3_matquality_new = 2 if q6_3_3_matquality == "Poor"
replace q6_3_3_matquality_new = 3 if q6_3_3_matquality == "Good"
replace q6_3_3_matquality_new = 4 if q6_3_3_matquality == "Very good"
replace q6_3_3_matquality_new = 5 if q6_3_2_matavail == "Excellent"

label define MATQUALITY 1 "Very poor" 2 "Poor" 3 "Good" 4 "Very good" 5 "Excellent"
label values q6_3_3_matquality_new MATQUALITY
label variable q6_3_3_matquality_new "q6_3_2_matavail (numeric)"
tab q6_3_3_matquality_new


**q6_3_4_overallquality


gen q6_3_4_overallquality_new =.
replace q6_3_4_overallquality_new = 1 if q6_3_4_overallquality == "Very poor"
replace q6_3_4_overallquality_new = 2 if q6_3_4_overallquality == "Poor"
replace q6_3_4_overallquality_new = 3 if q6_3_4_overallquality == "Good"
replace q6_3_4_overallquality_new = 4 if q6_3_4_overallquality == "Very good"
replace q6_3_4_overallquality_new = 5 if q6_3_4_overallquality == "Excellent"

label define OVERALLQUALITY 1 "Very poor" 2 "Poor" 3 "Good" 4 "Very good" 5 "Excellent"
label values q6_3_4_overallquality_new OVERALLQUALITY
label variable q6_3_4_overallquality_new "q6_3_4_overallquality (numeric)"
tab q6_3_4_overallquality_new


**q6_3_5_trainerquality

gen q6_3_5_trainerquality_new =.
replace q6_3_5_trainerquality_new = 1 if q6_3_5_trainerquality == "Very poor"
replace q6_3_5_trainerquality_new = 2 if q6_3_5_trainerquality == "Poor"
replace q6_3_5_trainerquality_new = 3 if q6_3_5_trainerquality == "Good"
replace q6_3_5_trainerquality_new = 4 if q6_3_5_trainerquality == "Very good"
replace q6_3_5_trainerquality_new = 5 if q6_3_5_trainerquality == "Excellent"

label define TRAINERQUALITY 1 "Very poor" 2 "Poor" 3 "Good" 4 "Very good" 5 "Excellent"
label values q6_3_5_trainerquality_new TRAINERQUALITY
label variable q6_3_5_trainerquality_new "q6_3_5_trainerquality (numeric)"
tab q6_3_5_trainerquality_new



**q6_3_6_industryfacility

gen q6_3_6_industryfacility_new =.
replace q6_3_6_industryfacility_new = 1 if q6_3_6_industryfacility == "Very poor"
replace q6_3_6_industryfacility_new = 2 if q6_3_6_industryfacility == "Poor"
replace q6_3_6_industryfacility_new = 3 if q6_3_6_industryfacility == "Good"
replace q6_3_6_industryfacility_new = 4 if q6_3_6_industryfacility == "Very good"
replace q6_3_6_industryfacility_new = 5 if q6_3_6_industryfacility == "Excellent"

label define INDUSTRYFACILITY 1 "Very poor" 2 "Poor" 3 "Good" 4 "Very good" 5 "Excellent"
label values q6_3_6_industryfacility_new INDUSTRYFACILITY
label variable q6_3_6_industryfacility_new "q6_3_5_trainerquality (numeric)"
tab q6_3_6_industryfacility_new


**q7_2_progtype
tab q7_2_progtype
gen q7_2_progtype_new =.
replace q7_2_progtype_new = 1 if q7_2_progtype == "Education program"
replace q7_2_progtype_new = 2 if q7_2_progtype == "Vocational / technical training"
replace q7_2_progtype_new = 88 if q7_2_progtype == "Other"

label define PROGTYPEq7_2 1 "Education program"2 "Vocational / technical training" 88 "Other"
label values q7_2_progtype_new PROGTYPEq7_2
label variable q7_2_progtype_new "q7_2_educ_specify (numeric)"
tab q7_2_progtype_new




**q7_3_aftertraining

gen q7_3_aftertraining_new =.
replace q7_3_aftertraining_new = 1 if q7_3_aftertraining == "Get a job Self-employment"
replace q7_3_aftertraining_new = 2 if q7_3_aftertraining == "Continue to higher level of training in same trade"
replace q7_3_aftertraining_new = 3 if q7_3_aftertraining == "Continue/go back to ongoing education"
replace q7_3_aftertraining_new = 4 if q7_3_aftertraining == "Continue/go back to ongoing job/business"
replace q7_3_aftertraining_new = 5 if q7_3_aftertraining == "Overseas employment"
replace q7_3_aftertraining_new = 88 if q7_3_aftertraining == "Others"

label define AFTERTRAINING 1 "Get a job Self-employment" 2 "Continue to higher level of training in same trade" 3 "Continue/go back to ongoing education" 4 "Continue/go back to ongoing job/business" 5 "Overseas employment" 88 "Others"
label values q7_3_aftertraining_new AFTERTRAINING
label variable q7_3_aftertraining_new  "q7_3_aftertraining (numeric)"
tab q7_3_aftertraining_new 
 

**q7_4_overallsat
gen q7_4_overallsat_new =.
replace q7_4_overallsat_new = 1 if q7_4_overallsat == "Very dissatisfied"
replace q7_4_overallsat_new = 2 if q7_4_overallsat == "Dissatisfied"
replace q7_4_overallsat_new = 3 if q7_4_overallsat == "Neutral"
replace q7_4_overallsat_new = 4 if q7_4_overallsat == "Satisfied"
replace q7_4_overallsat_new = 5 if q7_4_overallsat == "Very satisfied"

label define OVERALLSAT 1 "Very dissatisfied" 2 "Dissatisfied" 3 "Neutral" 4 "Satisfied" 5 "Very satisfied"
label values q7_4_overallsat_new OVERALLSAT
label variable q7_4_overallsat_new "q7_4_overallsat (numeric)"
tab q7_4_overallsat_new




**q7_5_recommend

gen q7_5_recommend_new =.
replace q7_5_recommend_new = 1 if q7_5_recommend == "Yes"
replace q7_5_recommend_new = 2 if q7_5_recommend == "No"
replace q7_5_recommend_new = 3 if q7_5_recommend == "Unsure"

label define RECOMMEND 1 "Yes" 2 "No" 3 "Unsure"
label values q7_5_recommend_new RECOMMEND
label variable q7_5_recommend_new "q7_5_recommend (numeric)"
tab q7_5_recommend_new


**q8a_1_edulevel

gen q8a_1_edulevel_new =.
replace q8a_1_edulevel_new = 1 if  q8a_1_edulevel == "No formal education"
replace q8a_1_edulevel_new = 2 if  q8a_1_edulevel == "Grade 5"
replace q8a_1_edulevel_new = 3 if  q8a_1_edulevel == "Grade 8"
replace q8a_1_edulevel_new = 4 if  q8a_1_edulevel == "SSC or equivalent"
replace q8a_1_edulevel_new = 5 if  q8a_1_edulevel == "HSC or equivalent"
replace q8a_1_edulevel_new = 6 if  q8a_1_edulevel == "Diploma"
replace q8a_1_edulevel_new = 7 if  q8a_1_edulevel == "Bachelor's degree"
replace q8a_1_edulevel_new = 8 if  q8a_1_edulevel == "Master's degree or above"


label define EDULEVEL 1 "No formal education" 2 "Grade 5" 3 "Grade 8" 4 "SSC or equivalent" 5 "HSC or equivalent" 6 "Diploma" 7 "Bachelor's degree" 8 "Master's degree or above"
label values q8a_1_edulevel_new EDULEVEL
label variable q8a_1_edulevel_new "q8a_1_edulevel (numeric)"
tab q8a_1_edulevel_new




**q8a_4_priortraindur
gen q8a_4_priortraindur_new =.
replace q8a_4_priortraindur_new = 1 if  q8a_4_priortraindur == "Less than 3 months"
replace q8a_4_priortraindur_new = 2 if  q8a_4_priortraindur == "3–6 months"
replace q8a_4_priortraindur_new = 3 if  q8a_4_priortraindur == "More than 6 months"


label define PRIORTRAINDUR 1 "Less than 3 months" 2 "3–6 months" 3 "More than 6 months"
label values q8a_4_priortraindur_new PRIORTRAINDUR
label variable q8a_4_priortraindur_new "q8a_4_priortraindur (Numeric)"
tab q8a_4_priortraindur_new




**q8a_6_prejobstatus
gen q8a_6_prejobstatus_new =.
replace q8a_6_prejobstatus_new = 1 if q8a_6_prejobstatus == "Employed (paid job) → Continue to Q8.7"
replace q8a_6_prejobstatus_new = 2 if q8a_6_prejobstatus == "Self-employed / own business    → Continue to Q8.7"
replace q8a_6_prejobstatus_new = 3 if q8a_6_prejobstatus == "Unemployed, → Skip to Q8.11"
replace q8a_6_prejobstatus_new = 4 if q8a_6_prejobstatus == "Student / still in education    → Skip to Q8.12"
replace q8a_6_prejobstatus_new = 5 if q8a_6_prejobstatus == "Part Time job"
replace q8a_6_prejobstatus_new = 88 if q8a_6_prejobstatus == "Other"


label define PREJOBSTATUS 1 "Employed (paid job)" 2 "Self-employed / own business" 3"Unemployed" 4 "Student / still in education" 5 "Part Time job" 88 "Other"

label values q8a_6_prejobstatus_new PREJOBSTATUS
label variable q8a_6_prejobstatus_new "q8a_6_prejobstatus (numeric)"
tab q8a_6_prejobstatus_new 






**q8a_7_worktype

gen q8a_7_worktype_new =.
replace q8a_7_worktype_new = 1 if q8a_7_worktype =="Factory / industrial worker"
replace q8a_7_worktype_new = 2 if q8a_7_worktype =="Office / clerical / administrative"
replace q8a_7_worktype_new = 3 if q8a_7_worktype =="Retail / sales"
replace q8a_7_worktype_new = 4 if q8a_7_worktype =="Agriculture / fisheries"
replace q8a_7_worktype_new = 5 if q8a_7_worktype =="Construction / labor"
replace q8a_7_worktype_new = 6 if q8a_7_worktype =="Own business / trade"
replace q8a_7_worktype_new = 7 if q8a_7_worktype =="Domestic / household services"
replace q8a_7_worktype_new = 88 if q8a_7_worktype =="Other"


label define WORKTYPE 1 "Factory / industrial worker" 2 "Office / clerical / administrative" 3 "Retail / sales" 4 "Agriculture / fisheries" 5 "Construction / labor" 6 "Own business / trade" 7 "Domestic / household services" 88 "Other"
label values q8a_7_worktype_new WORKTYPE
label variable q8a_7_worktype_new "q8a_7_worktype (numeric)"
tab q8a_7_worktype_new



**q8a_8_worksector

gen q8a_8_worksector_new =.
replace q8a_8_worksector_new = 1 if q8a_8_worksector == "Agrofood"
replace q8a_8_worksector_new = 2 if q8a_8_worksector == "Garment, Textile & Leather"
replace q8a_8_worksector_new = 3 if q8a_8_worksector == "ICT"
replace q8a_8_worksector_new = 4 if q8a_8_worksector == "Tourism & Hospitality"
replace q8a_8_worksector_new = 5 if q8a_8_worksector == "Construction & Civil"
replace q8a_8_worksector_new = 6 if q8a_8_worksector == "Light Engineering & Manufacturing"
replace q8a_8_worksector_new = 7 if q8a_8_worksector == "Domestic / household services"
replace q8a_8_worksector_new = 88 if q8a_8_worksector == "Other"


label define WORKSECTOR 1 "Agrofood" 2 "Garment, Textile & Leather" 3 "ICT" 4 "Tourism & Hospitality" 5"Construction & Civil" 6 "Light Engineering & Manufacturing" 7 "Domestic / household services" 88 "Other"
label values q8a_8_worksector_new WORKSECTOR
label variable q8a_8_worksector_new "q8a_8_worksector (Numeric)"
tab q8a_8_worksector_new




**q8a_10_preincome

gen q8a_10_preincome_new =.
replace q8a_10_preincome_new = 1 if q8a_10_preincome == "Below BDT 8,000"

replace q8a_10_preincome_new = 2 if q8a_10_preincome =="BDT 8,001–12,000"
replace q8a_10_preincome_new = 3 if q8a_10_preincome =="BDT 12,001–18,000"
replace q8a_10_preincome_new = 4 if q8a_10_preincome =="BDT 18,001–25,000"
replace q8a_10_preincome_new = 5 if q8a_10_preincome =="Above BDT 25,000"

label define PREINCOME 1 "Below BDT 8,000" 2 "BDT 8,001–12,000" 3 "BDT 12,001–18,000" 4 "BDT 18,001–25,000" 5 "Above BDT 25,000"
label values q8a_10_preincome_new PREINCOME
label variable q8a_10_preincome_new "q8a_10_preincome (numeric)"
tab q8a_10_preincome_new





**q8a_11_unemploydur
gen q8a_11_unemploydur_new =.
replace q8a_11_unemploydur_new = 1 if q8a_11_unemploydur == "Less than 3 months"
replace q8a_11_unemploydur_new = 2 if q8a_11_unemploydur == "3–6 months"
replace q8a_11_unemploydur_new = 3 if q8a_11_unemploydur == "7–12 months"
replace q8a_11_unemploydur_new = 4 if q8a_11_unemploydur == "More than 12 months"


label define UNEMPLOYDUR 1 "Less than 3 months" 2 "3–6 months" 3 "7–12 months" 4 "More than 12 months"
label values q8a_11_unemploydur_new UNEMPLOYDUR
label variable q8a_11_unemploydur_new "q8a_11_unemploydur (Numeric)"
tab q8a_11_unemploydur_new 



**q8a_12_joinreason
gen q8a_12_joinreason_new =.
replace q8a_12_joinreason_new = 1 if q8a_12_joinreason == "Unemployed and needed skills to find work"

replace q8a_12_joinreason_new = 2 if q8a_12_joinreason == "Wanted to improve earnings in current job"
replace q8a_12_joinreason_new = 3 if q8a_12_joinreason == "Wanted to change career / sector"
replace q8a_12_joinreason_new = 4 if q8a_12_joinreason == "Employer / factory sponsored or required it"
replace q8a_12_joinreason_new = 5 if q8a_12_joinreason =="Family encouragement"
replace q8a_12_joinreason_new = 6 if q8a_12_joinreason == "Personal interest in the trade"
replace q8a_12_joinreason_new = 88 if q8a_12_joinreason == "Other"


label define JOINREASON 1 "Unemployed and needed skills to find work" 2 "Wanted to improve earnings in current job" 3 "Wanted to change career / sector" 4 "Employer / factory sponsored or required it" 5 "Family encouragement" 6 "Personal interest in the trade" 88 "Other"
label values q8a_12_joinreason_new JOINREASON
label variable q8a_12_joinreason_new "q8a_12_joinreason (Numeric)"
tab q8a_12_joinreason_new



**q8a_13_primaryearner
gen q8a_13_primaryearner_new =. 
replace q8a_13_primaryearner_new = 1 if q8a_13_primaryearner == "Yes, sole earner"
replace q8a_13_primaryearner_new = 2 if q8a_13_primaryearner == "Yes, main earner (others contribute less)"
replace q8a_13_primaryearner_new = 3 if q8a_13_primaryearner =="No, another person is the main earner"
replace q8a_13_primaryearner_new = 4 if q8a_13_primaryearner == "No, household has multiple equal earners"

label define PRIMARYEARNER 1 "Yes, sole earner" 2 "Yes, main earner (others contribute less)" 3 "No, another person is the main earner" 4 "No, household has multiple equal earners"
label values q8a_13_primaryearner_new PRIMARYEARNER
label variable q8a_13_primaryearner_new "q8a_13_primaryearner (Numeric)"
tab q8a_13_primaryearner_new



**q8c_1_aspiration

gen q8c_1_aspiration_new =.
replace q8c_1_aspiration_new = 1 if q8c_1_aspiration == "On-the-job training"
replace q8c_1_aspiration_new = 2 if q8c_1_aspiration == "Higher NTVQF / qualification level (e.g. Level 4 to Level 5)"
replace q8c_1_aspiration_new = 3 if q8c_1_aspiration == "Further higher education (Diploma / BSc)"
replace q8c_1_aspiration_new = 4 if q8c_1_aspiration == "Short-term specialized courses"
replace q8c_1_aspiration_new = 5 if q8c_1_aspiration == "No further training planned"


label define ASPIRATION 1 "On-the-job training" 2 "Higher NTVQF / qualification level (e.g. Level 4 to Level 5)" 3 "Further higher education (Diploma / BSc)" 4 "Short-term specialized courses" 5 "No further training planned"

label values q8c_1_aspiration_new ASPIRATION
label variable q8c_1_aspiration_new "q8c_1_aspiration (Numeric)"
tab q8c_1_aspiration_new




**q8c_2_furtherreason

gen q8c_2_furtherreason_new =.
replace q8c_2_furtherreason_new = 1 if q8c_2_furtherreason == "To get a promotion or pay rise"
replace q8c_2_furtherreason_new = 2 if q8c_2_furtherreason == "To adapt to new technologies"
replace q8c_2_furtherreason_new = 3 if q8c_2_furtherreason == "To start my own business"
replace q8c_2_furtherreason_new = 4 if q8c_2_furtherreason == "To seek overseas employment"
replace q8c_2_furtherreason_new = 88 if q8c_2_furtherreason == "Other"


label define FURTHERREASON  1 "To get a promotion or pay rise" 2 "To adapt to new technologies" 3 "To start my own business" 4 "To seek overseas employment" 88 "Other"
label values q8c_2_furtherreason_new FURTHERREASON
label variable q8c_2_furtherreason_new "q8c_2_furtherreason (Numeric)"
tab q8c_2_furtherreason_new






**q9a_7_awaitingdur // no observation now



*gen q9a_7_awaitingdur_new =.
*replace q9a_7_awaitingdur_new = 1 if q9a_7_awaitingdur == "2 weeks"
*replace q9a_7_awaitingdur_new = 2 if q9a_7_awaitingdur == "1 month"
*replace q9a_7_awaitingdur_new = 3 if q9a_7_awaitingdur == "2 months"
*replace q9a_7_awaitingdur_new = 4 if q9a_7_awaitingdur == "More than 2 months"
*replace q9a_7_awaitingdur_new = 88 if q9a_7_awaitingdur == "Others"

*label define AWAITINGDUR 1 "2 weeks" 2 "1 month" 3 "2 months" 4 "More than 2 months" 88 "Others"

*label values q9a_7_awaitingdur_new AWAITINGDUR
*label variable q9a_7_awaitingdur_new "q9a_7_awaitingdur (Numeric)"
*tab q9a_7_awaitingdur_new


**q9_1_empstatus // no observation now


*gen q9_1_empstatus_new =.

*replace q9_1_empstatus_new = 1 if q9_1_empstatus ==  "Employed (paid job or self-employed)"
*replace q9_1_empstatus_new = 2 if q9_1_empstatus ==  "Still looking for work"
*replace q9_1_empstatus_new = 3 if q9_1_empstatus == "Not working and not looking"

*label define EMPSTATUS 1 "Employed (paid job or self-employed)" 2 "Still looking for work" 3 "Not working and not looking"
*label values q9_1_empstatus_new EMPSTATUS
*label variables q9_1_empstatus_new "q9_1_empstatus (Numeric)"
*tab q9_1_empstatus_new





**q9b_1_timetojob // no observation now

*gen q9b_1_timetojob_new =.
*replace q9b_1_timetojob_new = 1 if q9b_1_timetojob == "Immediately"
*replace q9b_1_timetojob_new = 2 if q9b_1_timetojob =="1 month"
*replace q9b_1_timetojob_new = 3 if q9b_1_timetojob =="2 months"
*replace q9b_1_timetojob_new = 4 if q9b_1_timetojob =="3 months"
*replace q9b_1_timetojob_new = 5 if q9b_1_timetojob =="More than 3 months"
*replace q9b_1_timetojob_new = 88 if q9b_1_timetojob =="Others"


*label define TIMETOJOB 1 "Immediately" 2 "1 month" 3 "2 months" 4 "3 months" 5 "More than 3 months" 88 "Others"
*label values q9b_1_timetojob_new TIMETOJOB
*label variables q9b_1_timetojob_new "q9b_1_timetojob (Numeric)"
*tab q9b_1_timetojob_new





**q9b_3_whynotsector // no observation now

*gen q9b_3_whynotsector_new =.
*replace q9b_3_whynotsector_new = 1 if q9b_3_whynotsector == "No vacancies in trained sector"
*replace q9b_3_whynotsector_new = 2 if q9b_3_whynotsector == "Better pay / benefits elsewhere"
*replace q9b_3_whynotsector_new = 3 if q9b_3_whynotsector == "Did not find the trade interesting"
*replace q9b_3_whynotsector_new = 4 if q9b_3_whynotsector == "Workplace too far from home"
*replace q9b_3_whynotsector_new = 5 if q9b_3_whynotsector == "Family / personal reasons"
*replace q9b_3_whynotsector_new = 88 if q9b_3_whynotsector == "Other"


*label define WHYNORSECTOR 1 "No vacancies in trained sector" 2 "Better pay / benefits elsewhere" 3 "Did not find the trade interesting" 4 "Workplace too far from home" 5 "Family / personal reasons" 88 "Other"
*label values q9b_3_whynotsector_new WHYNORSECTOR
*label variables q9b_3_whynotsector_new "q9b_3_whynotsector (Numeric)"
*tab q9b_3_whynotsector_new





**q9b_4_emptype // no observation now

*gen q9b_4_emptype_new =.
*replace q9b_4_emptype_new = 1 if q9b_4_emptype == "Full-time"
*replace q9b_4_emptype_new = 2 if q9b_4_emptype == "Part-time"
*replace q9b_4_emptype_new = 3 if q9b_4_emptype == "Self-employed"
*replace q9b_4_emptype_new = 4 if q9b_4_emptype == "Freelance"
*replace q9b_4_emptype_new = 88 if q9b_4_emptype == "Other"

*label define EMPTYPE 1 "Full-time" 2 "Part-time" 3 "Self-employed" 4 "Freelance" 88 "Other"
*label values q9b_4_emptype_new EMPTYPE
*label variables q9b_4_emptype_new "q9b_4_emptype (Numeric)"
*tab q9b_4_emptype_new





**q9b_5_joblocation // no observation now

*gen q9b_5_joblocation_new =.

*replace q9b_5_joblocation_new = 1 if q9b_5_joblocation == "Same factory / workplace where I trained"
*replace q9b_5_joblocation_new = 2 if q9b_5_joblocation == "Another branch / factory of same industry partner"
*replace q9b_5_joblocation_new = 3 if q9b_5_joblocation == "An entirely dierent company or employer"
*replace q9b_5_joblocation_new = 4 if q9b_5_joblocation == "Self-employed (home-based or own shop)"


*label define JOBLOCATION 1 "Same factory / workplace where I trained" 2 "Another branch / factory of same industry partner" 3 "An entirely dierent company or employer" 4 "Self-employed (home-based or own shop)"
*label values q9b_5_joblocation_new JOBLOCATION
*label variables q9b_5_joblocation_new "q9b_5_joblocation (Numeric)"
*tab q9b_5_joblocation_new



**q9b_6_jobsource // no observation now

*gen q9b_6_jobsource_new =.


*replace q9b_6_jobsource_new = 1 if q9b_6_jobsource =="Placement by the training provider / institute"
*replace q9b_6_jobsource_new = 2 if q9b_6_jobsource =="Personal contacts (friends / family)"
*replace q9b_6_jobsource_new = 3 if q9b_6_jobsource =="Job advertisement (online or newspaper)"
*replace q9b_6_jobsource_new = 4 if q9b_6_jobsource =="Self-eort / started own business"
*replace q9b_6_jobsource_new = 88 if q9b_6_jobsource =="Other"


*label define JOBSOURCE 1 "Placement by the training provider / institute" 2 "Personal contacts (friends / family)" 3 "Job advertisement (online or newspaper)" 4 "Self-eort / started own business" 88 "Other"
*label values q9b_6_jobsource_new JOBSOURCE
*label variables q9b_6_jobsource_new "q9b_6_jobsource (Numeric)"
*tab q9b_6_jobsource_new




**q9b_7_currentincome // no observation now

*gen q9b_7_currentincome_new =.


*replace q9b_7_currentincome_new = 1 if q9b_7_currentincome == "No income / was not working"
*replace q9b_7_currentincome_new = 2 if q9b_7_currentincome == "Below BDT 8,000"
*replace q9b_7_currentincome_new = 3 if q9b_7_currentincome == "BDT 8,001–12,000"
*replace q9b_7_currentincome_new = 4 if q9b_7_currentincome == "BDT 12,001–18,000"
*replace q9b_7_currentincome_new = 5 if q9b_7_currentincome == "BDT 18,001–25,000"
*replace q9b_7_currentincome_new = 6 if q9b_7_currentincome == "Above BDT 25,000"



*label define CURRENTINCOME 1 "No income / was not working" 2 "Below BDT 8,000" 3 "BDT 8,001–12,000" 4 "BDT 12,001–18,000" 5 "BDT 18,001–25,000" 6 "Above BDT 25,000"
*label values q9b_7_currentincome_new CURRENTINCOME
*label variables q9b_7_currentincome_new "q9b_7_currentincome (Numeric)"
*tab q9b_7_currentincome_new





**q9b_8_incomechange // no observation now

*gen q9b_8_incomechange_new =.
*replace q9b_8_incomechange_new = 1 if q9b_8_incomechange_new == "Increased signicantly (more than 25%)"
*replace q9b_8_incomechange_new = 2 if q9b_8_incomechange_new == "Increased slightly (up to 25%)"
*replace q9b_8_incomechange_new = 3 if q9b_8_incomechange_new == "No change"
*replace q9b_8_incomechange_new = 4 if q9b_8_incomechange_new == "Decreased"
*replace q9b_8_incomechange_new = 5 if q9b_8_incomechange_new == "Cannot compare — was not working before EBT"


*label define INCOMECHANGE 1 "Increased signicantly (more than 25%)" 2 "Increased slightly (up to 25%)" 3 "No change" 4 "Decreased" 5 "Cannot compare — was not working before EBT"
*label values q9b_8_incomechange_new INCOMECHANGE
*label variables q9b_8_incomechange_new "q9b_8_incomechange (Numeric)"
*tab q9b_8_incomechange_new





**q9b_9_ebtcontributed // no observation now

*gen q9b_9_ebtcontributed_new =.
 
*replace q9b_9_ebtcontributed_new = 1 if q9b_9_ebtcontributed == "Yes, directly"
*replace q9b_9_ebtcontributed_new = 2 if q9b_9_ebtcontributed == "Partially"
*replace q9b_9_ebtcontributed_new = 3 if q9b_9_ebtcontributed == "No"
*replace q9b_9_ebtcontributed_new = 4 if q9b_9_ebtcontributed == "Too early to say"


*label define EBTCONTRIBUTED 1 "Yes, directly" 2 "Partially" 3 "No" 4 "Too early to say"
*label values q9b_9_ebtcontributed_new EBTCONTRIBUTED
*label variables q9b_9_ebtcontributed_new "q9b_9_ebtcontributed (Numeric)"
*tab q9b_9_ebtcontributed_new


**q9d_2_planlook // no observation now

*gen q9d_2_planlook_new =.
*replace q9d_2_planlook_new = 1 if q9d_2_planlook == "Yes"
*replace q9d_2_planlook_new = 2 if q9d_2_planlook == "No"
*replace q9d_2_planlook_new = 3 if q9d_2_planlook == "Unsure"

*label define PLANLOOK 1 "Yes" 2 "No" 3 "Unsure"
*label values q9d_2_planlook_new PLANLOOK
*label variables q9d_2_planlook_new "q9d_2_planlook (Numeric)"
*tab q9d_2_planlook_new




*==============================================================================
* SECTION 4: New category Variables
*==============================================================================


**q7_age
tabstat q7_age, statistics(min max)
recode q7_age ///
    (min/17 = 1 "Below 18 years") ///
    (18/22 = 2 "18 to 22 years") ///
    (23/26 = 3 "23 to 26 years") ///
    (27/31 = 4 "27 to 31 years") ///
    (32/36 = 5 "32 to 36 years") ///
    (37/41 = 6 "37 to 41 years") ///
    (42/max = 7 "42 years and above"), gen(q7_age_cat)
tab q7_age_cat
label variable q7_age_cat "Age category"


**q2a_3_attachdur
tabstat q2a_3_attachdur, statistics(min max)
tab q2a_3_attachdur
recode q2a_3_attachdur ///
		(min/9 = 1 "Below 10 days") ///
		(10/19 = 2 "10 to 19 days") ///
		(20/29 = 3 "20 to 29 days") ///
		(30/max = 4 "30 days and above"), gen (q2a_3_attachdur_cat)
tab q2a_3_attachdur_cat
label variable q2a_3_attachdur_cat "q2a_3_attachdur (Category)"


** q2b_1_1 (Both) [Classes per week]

tabstat q2b_1_1_classwk_org, statistics(min max)
tab q2b_1_1_classwk_org
recode q2b_1_1_classwk_org ///
		(min/3 = 1 "1 to 3 Classes per week") ///
		(4/6 = 2 "4 to 6 classes per week") ///
		(7/10 = 3 "7 to 10 classes per week") ///
		(11/max =4 "11 classes and above per week"), gen (q2b_1_1_classwk_org_cat)
tab q2b_1_1_classwk_org_cat
label variable q2b_1_1_classwk_org_cat "q2b_1_1_classwk_org (Category)"

tabstat q2b_1_1_classwk_resp, statistics(min max)
tab q2b_1_1_classwk_resp
recode q2b_1_1_classwk_resp ///
		(min/3 = 1 "1 to 3 Classes per week") ///
		(4/6 = 2 "4 to 6 classes per week") ///
		(7/10 = 3 "7 to 10 classes per week") ///
		(11/max =4 "11 classes and above per week"), gen (q2b_1_1_classwk_resp_cat)
tab q2b_1_1_classwk_resp_cat
label variable q2b_1_1_classwk_resp "q2b_1_1_classwk_resp (Category)"



** q2b_1_2 (Both) [Total weeks of class held]

tabstat q2b_1_2_weeksheld_org, statistics(min max)
tab q2b_1_2_weeksheld_org
recode q2b_1_2_weeksheld_org ///
		(min/3 = 1 "1 to 3 weeks") ///
		(4/6 = 2 "4 to 6 weeks") ///
		(7/10 = 3 "7 to 10 weeks") ///
		(11/max = 4 "11 weeks and above"), gen (q2b_1_2_weeksheld_org_cat)
tab q2b_1_2_weeksheld_org_cat
label variable q2b_1_2_weeksheld_org_cat "q2b_1_2_weeksheld_org (Category)"




tabstat q2b_1_2_weeksheld_resp, statistics(min max)
tab q2b_1_2_weeksheld_resp
recode q2b_1_2_weeksheld_resp ///
		(min/3 = 1 "1 to 3 weeks") ///
		(4/6 = 2 "4 to 6 weeks") ///
		(7/10 = 3 "7 to 10 weeks") ///
		(11/max = 4 "11 weeks and above"), gen (q2b_1_2_weeksheld_resp_cat)
tab q2b_1_2_weeksheld_resp_cat
label variable q2b_1_2_weeksheld_resp_cat "q2b_1_2_weeksheld_resp (Category)"


**q2b_1_3 (Both) [Total Number of Classes]


tabstat q2b_1_3_totalclass_org, statistics(min max)
tab q2b_1_3_totalclass_org
recode q2b_1_3_totalclass_org ///
		(min/9 = 1 "Below 10 Classes") ///
		(10/19 = 2 "10 to 19 Classes") ///
		(20/29 = 3 "20 to 29 Classes") ///
		(30/39 = 4 "30 to 39 Classes") ///
		(40/49 = 5 "40 to 49 Classes") ///
		(50/59 = 6 "50 to 59 Classes") ///
		(60/69 = 7 "60 to 69 Classes") ///
		(70/max = 8 "70 Classes and above"), gen (q2b_1_3_totalclass_org_cat)
tab q2b_1_3_totalclass_org_cat
label variable q2b_1_3_totalclass_org_cat "q2b_1_3_totalclass_org (Category)"


tabstat q2b_1_3_totalclass_resp, statistics(min max)
tab q2b_1_3_totalclass_resp
recode q2b_1_3_totalclass_resp ///
		(min/9 = 1 "Below 10 Classes") ///
		(10/19 = 2 "10 to 19 Classes") ///
		(20/29 = 3 "20 to 29 Classes") ///
		(30/39 = 4 "30 to 39 Classes") ///
		(40/49 = 5 "40 to 49 Classes") ///
		(50/59 = 6 "50 to 59 Classes") ///
		(60/69 = 7 "60 to 69 Classes") ///
		(70/max = 8 "70 Classes and above"), gen (q2b_1_3_totalclass_resp_cat)
tab q2b_1_3_totalclass_resp_cat
label variable q2b_1_3_totalclass_resp_cat "q2b_1_3_totalclass_resp (Category)"


** q2b_1_4 (Both) [Avg hours per classess] //  no need to do category
tabstat q2b_1_4_hrsperclass_org, statistics(min max)
tab q2b_1_4_hrsperclass_org


tabstat q2b_1_4_hrsperclass_resp, statistics(min max)
tab q2b_1_4_hrsperclass_resp


** q2b_1_5 (BOth) [Total Number of Class Hours]

tabstat q2b_1_5_totalhrs_org, statistics(min max)
tab q2b_1_5_totalhrs_org
recode q2b_1_5_totalhrs_org ///
		(min/49 = 1 "Below 50 hours") ///
		(50/99 = 2 "50 to 99 hours") ///
		(100/149 = 3 "100 to 149 hours" ) ///
		(150/199 = 4 "150 to 199 hours") ///
		(200/249 = 5 "200 to 249 hours") ///
		(250/299 = 6 "250 to 299 hours") ///
		(300/max = 7 "300 hours and above"), gen (q2b_1_5_totalhrs_org_cat)
tab q2b_1_5_totalhrs_org_cat
label variable q2b_1_5_totalhrs_org_cat "q2b_1_5_totalhrs_org (Category)"


tabstat q2b_1_5_totalhrs_resp, statistics(min max)
tab q2b_1_5_totalhrs_org
recode q2b_1_5_totalhrs_resp ///
		(min/49 = 1 "Below 50 hours") ///
		(50/99 = 2 "50 to 99 hours") ///
		(100/149 = 3 "100 to 149 hours" ) ///
		(150/199 = 4 "150 to 199 hours") ///
		(200/249 = 5 "200 to 249 hours") ///
		(250/299 = 6 "250 to 299 hours") ///
		(300/max = 7 "300 hours and above"), gen (q2b_1_5_totalhrs_resp_cat)
tab q2b_1_5_totalhrs_resp_cat
label variable q2b_1_5_totalhrs_resp_cat "q2b_1_5_totalhrs_resp (Category)"


**q5_2_stipendamt // no observation now
tabstat q5_2_stipendamt
tab q5_2_stipendamt


**q5_6_1a_fee /// no need to do category
tabstat q5_6_1a_fee, statistics(min max)
tab q5_6_1a_fee


**q5_6_2a_transport
tabstat q5_6_2a_transport, statistics(min max)
tab q5_6_2a_transport
recode q5_6_2a_transport ///
		(min/999 = 1 "Below 1000 BDT") ///
		(1000/1999 = 2 "1000 to 1999 BDT") ///
		(2000/2999 = 3 "2000 to 2999 BDT") ///
		(3000/3999 = 4 "3000 to 3999 BDT") ///
		(4000/max = 5 "4000 BDT and above"), gen (q5_6_2a_transport_cat)
tab q5_6_2a_transport_cat
label variable q5_6_2a_transport_cat "q5_6_2a_transport (Category)"


**q5_6_3a_equip /// no need to do category
tabstat q5_6_3a_equip, statistics(min max)
tab q5_6_3a_equip


**q5_6_4a_stationery
tabstat q5_6_4a_stationery, statistics(min max)
tab q5_6_4a_stationery
recode q5_6_4a_stationery ///
		(min/499 = 1 "Below 500 BDT") ///
		(500/999 = 2 "500 to 999 BDT") ///
		(1000/1499 = 3 "1000 to 1499 BDT") ///
		(1500/max = 4 "1500 BDT and above"), gen (q5_6_4a_stationery_cat)
tab q5_6_4a_stationery_cat
label variable q5_6_4a_stationery_cat "q5_6_4a_stationery (Category)"
		
	
**q5_6_5b_othertotal1
tabstat q5_6_5b_othertotal1, statistics(min max)
tab q5_6_5b_othertotal1
recode q5_6_5b_othertotal1 ///
		(min/499 = 1 "Below 500 BDT") ///
		(500/999 = 2 "500 to 999 BDT") ///
		(1000/1499 = 3 "1000 to 1499 BDT") ///
		(1500/max = 4 "1500 BDT and above"), gen (q5_6_5b_othertotal1_cat)
tab q5_6_5b_othertotal1_cat
label variable q5_6_5b_othertotal1_cat "q5_6_5b_othertotal1 (Category)"		
		
		
**q5_6_6b_othertotal2 ///  no need to do category
tabstat q5_6_6b_othertotal2, statistics(min max)
tab q5_6_6b_othertotal2


** q8b_5_1 (All three) [Salary/Profit]

tab q8b_5_1a_monthswork12m
recode q8b_5_1a_monthswork12m ///
		(min/0 = 1 "Not worked") ///
		(1/3 = 2 "1 to 3 months") ///
		(4/6 = 3 "4 to 6 months") ///
		(7/9 = 4 "7 to 9 months") ///
		(10/12 = 5 "10 to 12 months"), gen (q8b_5_1a_monthswork12m_cat)
tab q8b_5_1a_monthswork12m_cat
label variable q8b_5_1a_monthswork12m_cat "q8b_5_1a_monthswork12m (Category)"


tab q8b_5_1b_daysmonth12m
recode q8b_5_1b_daysmonth12m /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_5_1b_daysmonth12m_cat)
tab  q8b_5_1b_daysmonth12m_cat
label variable  q8b_5_1b_daysmonth12m_cat " q8b_5_1b_daysmonth12m (Category)"


tab q8b_5_1c_hrsday12m
recode q8b_5_1c_hrsday12m ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_5_1c_hrsday12m_cat)
tab q8b_5_1c_hrsday12m_cat
label variable q8b_5_1c_hrsday12m_cat "q8b_5_1c_hrsday12m (category)"






** q8b_5_2 (All three) [Product for own use]

tab q8b_5_2a_monthsownuse12m
recode q8b_5_2a_monthsownuse12m ///
		(min/0 = 1 "Not worked") ///
		(1/3 = 2 "1 to 3 months") ///
		(4/6 = 3 "4 to 6 months") ///
		(7/9 = 4 "7 to 9 months") ///
		(10/12 = 5 "10 to 12 months"), gen (q8b_5_2a_monthsownuse12m_cat)
tab q8b_5_2a_monthsownuse12m_cat
label variable q8b_5_2a_monthsownuse12m_cat "q8b_5_2a_monthsownuse12m (Category)"


tab q8b_5_2b_daysmonthownuse12m
recode q8b_5_2b_daysmonthownuse12m /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_5_2b_daysmonthownuse12m_cat)
tab  q8b_5_2b_daysmonthownuse12m_cat
label variable  q8b_5_2b_daysmonthownuse12m_cat " q8b_5_2b_daysmonthownuse12m (Category)"


tab q8b_5_2c_hrsdayownuse12m
recode q8b_5_2c_hrsdayownuse12m ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_5_2c_hrsdayownuse12m_cat)
tab q8b_5_2c_hrsdayownuse12m_cat
label variable q8b_5_2c_hrsdayownuse12m_cat "q8b_5_2c_hrsdayownuse12m (category)"





*** q8b_6_1 (All three) [1 month before the training]

tab q8b_6_1a_days_pre
recode q8b_6_1a_days_pre /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_1a_days_pre_cat)
tab  q8b_6_1a_days_pre_cat
label variable  q8b_6_1a_days_pre_cat " q8b_6_1a_days_pre (Category)"

tab q8b_6_1b_hrs_pre
recode q8b_6_1b_hrs_pre ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_1b_hrs_pre_cat)
tab q8b_6_1b_hrs_pre_cat
label variable q8b_6_1b_hrs_pre_cat "q8b_6_1b_hrs_pre (category)"


tab q8b_6_1c_income_pre
recode q8b_6_1c_income_pre ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_1c_income_pre_cat)
tab q8b_6_1c_income_pre_cat
label variable q8b_6_1c_income_pre_cat "q8b_6_1c_income_pre (Category)"




*** q8b_6_2 (All three) [1st month after the training]

tab q8b_6_2a_days_m1
recode q8b_6_2a_days_m1 /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_2a_days_m1_cat)
tab  q8b_6_2a_days_m1_cat
label variable  q8b_6_2a_days_m1_cat " q8b_6_2a_days_m1 (Category)"

tab q8b_6_2b_hrs_m1
recode q8b_6_2b_hrs_m1 ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_2b_hrs_m1_cat)
tab q8b_6_2b_hrs_m1_cat
label variable q8b_6_2b_hrs_m1_cat "q8b_6_2b_hrs_m1 (category)"


tab q8b_6_2c_income_m1
recode q8b_6_2c_income_m1 ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_2c_income_m1_cat)
tab q8b_6_2c_income_m1_cat
label variable q8b_6_2c_income_m1_cat "q8b_6_2c_income_m1 (Category)"




*** q8b_6_3 (All three) [2nd month after the training]

tab q8b_6_3a_days_m2
recode q8b_6_3a_days_m2 /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_3a_days_m2_cat)
tab  q8b_6_3a_days_m2_cat
label variable  q8b_6_3a_days_m2_cat " q8b_6_3a_days_m2 (Category)"

tab q8b_6_3b_hrs_m2
recode q8b_6_3b_hrs_m2 ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_3b_hrs_m2_cat)
tab q8b_6_3b_hrs_m2_cat
label variable q8b_6_3b_hrs_m2_cat "q8b_6_3b_hrs_m2 (category)"


tab q8b_6_3c_income_m2
recode q8b_6_3c_income_m2 ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_3c_income_m2_cat)
tab q8b_6_3c_income_m2_cat
label variable q8b_6_3c_income_m2_cat "q8b_6_3c_income_m2 (Category)"





*** q8b_6_4 (All three) [3rd month after the training]

tab q8b_6_4a_days_m3
recode q8b_6_4a_days_m3 /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_4a_days_m3_cat)
tab  q8b_6_4a_days_m3_cat
label variable  q8b_6_4a_days_m3_cat " q8b_6_4a_days_m3 (Category)"

tab q8b_6_4b_hrs_m3
recode q8b_6_4b_hrs_m3 ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_4b_hrs_m3_cat)
tab q8b_6_4b_hrs_m3_cat
label variable q8b_6_4b_hrs_m3_cat "q8b_6_4b_hrs_m3 (category)"


tab q8b_6_4c_income_m3
recode q8b_6_4c_income_m3 ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_4c_income_m3_cat)
tab q8b_6_4c_income_m3_cat
label variable q8b_6_4c_income_m3_cat "q8b_6_4c_income_m3 (Category)"





*** q8b_6_5 (All three) [4th month after the training]

tab q8b_6_5a_days_m4
recode q8b_6_5a_days_m4 /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_5a_days_m4_cat)
tab  q8b_6_5a_days_m4_cat
label variable  q8b_6_5a_days_m4_cat " q8b_6_5a_days_m4 (Category)"

tab q8b_6_5b_hrs_m4
recode q8b_6_5b_hrs_m4 ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_5b_hrs_m4_cat)
tab q8b_6_5b_hrs_m4_cat
label variable q8b_6_5b_hrs_m4_cat "q8b_6_5b_hrs_m4 (category)"


tab q8b_6_5c_income_m4
recode q8b_6_5c_income_m4 ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_5c_income_m4_cat)
tab q8b_6_5c_income_m4_cat
label variable q8b_6_5c_income_m4_cat "q8b_6_5c_income_m4 (Category)"





*** q8b_6_6 (All three) [5th month after the training]

tab q8b_6_6a_days_m5
recode q8b_6_6a_days_m5 /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_6a_days_m5_cat)
tab  q8b_6_6a_days_m5_cat
label variable  q8b_6_6a_days_m5_cat " q8b_6_6a_days_m5 (Category)"

tab q8b_6_6b_hrs_m5
recode q8b_6_6b_hrs_m5 ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_6b_hrs_m5_cat)
tab q8b_6_6b_hrs_m5_cat
label variable q8b_6_6b_hrs_m5_cat "q8b_6_6b_hrs_m5 (category)"


tab q8b_6_6c_income_m5
recode q8b_6_6c_income_m5 ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_6c_income_m5_cat)
tab q8b_6_6c_income_m5_cat
label variable q8b_6_6c_income_m5_cat "q8b_6_6c_income_m5 (Category)"





*** q8b_6_7 (All three) [6th month after the training]

tab q8b_6_7a_days_m6
recode q8b_6_7a_days_m6 /// 
		(min/0 = 1 "Not worked") ///
		(1/9 = 2 "Below 10 days") ///
		(10/19 = 3 "10 to 19 days") ///
		(20/max = 4 "20 days and above"), gen (q8b_6_7a_days_m6_cat)
tab  q8b_6_7a_days_m6_cat
label variable  q8b_6_7a_days_m6_cat " q8b_6_7a_days_m6 (Category)"

tab q8b_6_7b_hrs_m6
recode q8b_6_7b_hrs_m6 ///
		(min/0 = 1 "Not worked") ///
		(1/7 = 2 "Below 8 hours") ///
		(8 = 3 "8 hours") ///
		(9/max = 4 "More than 8 hours"), gen (q8b_6_7b_hrs_m6_cat)
tab q8b_6_7b_hrs_m6_cat
label variable q8b_6_7b_hrs_m6_cat "q8b_6_7b_hrs_m6 (category)"


tab q8b_6_7c_income_m6
recode q8b_6_7c_income_m6 ///
		(min/0 = 1 "No income / was not working") ///
		(1/8000 = 2 "BDT 8,000 and Below") ///
		(8001/12000 = 3 "BDT 8,001–12,000") ///
		(12001/18000 = 4 "BDT 12,001–18,000") ///
		(18001/25000 = 5 "BDT 18,001–25,000") ///
		(25001/max = 6 "Above BDT 25,000"), gen (q8b_6_7c_income_m6_cat)
tab q8b_6_7c_income_m6_cat
label variable q8b_6_7c_income_m6_cat "q8b_6_7c_income_m6 (Category)"







*==============================================================================
* SECTION 5: MULTI-SELECT LABELING
*==============================================================================
foreach grp in q12_c01_newspaper-q12_c10_other q13_c01_getjob-q13_c07_other ///
               q6_2_c01_work-q6_2_c07_other q9a_6_c01_notscheduled-q9a_6_c04_other ///
               q9c_2_c01_novacancy-q9c_2_c07_other q9d_1_c01_study-q9d_1_c06_other {
    foreach v of varlist `grp' {
        label values `v' YESNO
    }
}
egen q12_nchannels   = rowtotal(q12_c01_newspaper-q12_c10_other)
egen q13_nreasons    = rowtotal(q13_c01_getjob-q13_c07_other)
egen q9c_2_nbarriers = rowtotal(q9c_2_c01_novacancy-q9c_2_c07_other)
label variable q12_nchannels   "Number of channels through which respondent heard about training"
label variable q13_nreasons    "Number of reasons cited for choosing training"
label variable q9c_2_nbarriers "Number of job-search barriers cited"

tab q12_nchannels
tab q13_nreasons
tab q9c_2_nbarriers // no observation now


*==============================================================================
* SECTION 6: TRAINING PROFILE
*==============================================================================





gen attendance_rate = q2b_1_3_totalclass_resp / q2b_1_3_totalclass_org
label variable attendance_rate "Share of offered classes actually attended (0-1)"
tab attendance_rate

gen obs_id = _n
list obs_id q2b_1_3_totalclass_resp q2b_1_3_totalclass_org attendance_rate ///
    if attendance_rate > 1



gen hours_attendance_rate = q2b_1_5_totalhrs_resp / q2b_1_5_totalhrs_org
label variable hours_attendance_rate "Share of offered training hours actually attended (0-1)"
tab hours_attendance_rate






*==============================================================================
* SECTION 7: Demographic table 
*==============================================================================

ssc install asdoc, replace
ssc install mrtab, replace 


local demog_cat q2_gender_new q3_status_new ///
		q6_ethnicity_new q7_age_cat q8_marital_new ///
		division_new q2_1_progtype_new sector_name_new ///
		q2_3_level_new q8a_1_edulevel_new

asdoc tab1 `demog_cat', save(trainee.doc)


asdoc tab1 division_new sector_name_new q2_1_progtype_new, ///
    save(trainee.doc) append 




*==============================================================================
* SECTION 8: Frequency table
*==============================================================================


*==============================================================================
* PART A: DESCRIPTIVE STATISTICS 
*==============================================================================


* --- A1. EBT training profile --------------------------------------------------
local training_freq q2_1_progtype_new ///
		sector_name_new q2_3_level_new ///
		q2_5_1_admitcycle_new q2_5_2_admitmonth_new q2_6_status_new
asdoc tab1 `training_freq', save(trainee.doc) append

* --- A2. Industry attachment (ADTP trainees only) -------------------------------
local attach_freq q2a_1_attach_new ///
			q2a_3_attachdur_cat ///
			q2a_4_logbook_new q2a_5_logfreq_new
asdoc tab1 `attach_freq', save(trainee.doc) append 

* --- A3. Instructional approach --------------------------------------------------
local instruct_freq q2b_1_1_classwk_org_cat q2b_1_1_classwk_resp_cat ///
                    q2b_1_2_weeksheld_org_cat q2b_1_2_weeksheld_resp_cat ///
                    q2b_1_3_totalclass_org_cat q2b_1_3_totalclass_resp_cat ///
                    q2b_1_5_totalhrs_org_cat q2b_1_5_totalhrs_resp_cat ///
                    q2b_2_cblm_new q2b_4_dualinstructor_new q2b_4_specify_new ///
                    q2b_5_wholed_new q2b_6_delivery_new
asdoc tab1 `instruct_freq', save(trainee.doc) append 

* --- A4. Occupational health, safety & PPE ---------------------------------------
local ohs_freq q3_1a_ppe_trainee_new q3_1b_ppe_enum_new q3_2_ppe_enforced_new ///
              q3_3_safetyguard_new q3_4_riskbriefed_new
asdoc tab1 `ohs_freq', save(trainee.doc) append 

* --- A5. Facility & safeguarding -------------------------------------------------
local facility_freq q4_access_new q5_toilet_access_new q4_1_water_new ///
                    q4_2_toilets_new q4_3_complaintinfo_new q4_4_complaintbox_new
asdoc tab1 `facility_freq', save(trainee.doc) append 

* --- A6. Financial support --------------------------------------------------------
** q5_3_stipendmode_new : no observation now, add when there will be observations.

local finance_freq q5_1_stipend_new  q5_6_1b_feefunder_new ///
                   q5_6_2a_transport_cat q5_6_2b_transportfunder_new ///
                   q5_6_3b_equipfunder_new q5_6_4a_stationery_cat ///
                   q5_6_4b_stationeryfunder_new q5_6_5b_othertotal1_cat ///
                   q5_6_5c_otherfunder1_new q5_6_6c_otherfunder2_new
asdoc tab1 `finance_freq', save(trainee.doc) append 



* --- A7. Training satisfaction & feedback ----------------------------------------
local satisfaction_freq q6_1_attendrate_new q6_3_1_scheduling_new q6_3_2_matavail_new ///
                        q6_3_3_matquality_new q6_3_4_overallquality_new ///
                        q6_3_5_trainerquality_new q6_3_6_industryfacility_new ///
                        q7_1_otherenrolled_new q7_2_progtype_new q7_3_aftertraining_new ///
                        q7_4_overallsat_new q7_5_recommend_new
asdoc tab1 `satisfaction_freq', save(trainee.doc) append 

* --- A8. Pre-training baseline status (Section 8A) --------------------------------
local baseline_freq q8a_1_edulevel_new q8a_3_priortraining_new q8a_4_priortraindur_new ///
                    q8a_5_priortrainrelated_new q8a_6_prejobstatus_new q8a_7_worktype_new ///
                    q8a_8_worksector_new q8a_9_workrelated_new q8a_10_preincome_new ///
                    q8a_11_unemploydur_new q8a_12_joinreason_new q8a_13_primaryearner_new
asdoc tab1 `baseline_freq', save(trainee.doc) append 

* --- A9. Concurrent work status (Section 8B, pre-training baseline only) & aspirations (8C) --
local workstatus_freq q8b_1_worked7d_new q8b_2_absentjob7d_new q8b_3_ownuse7d_new ///
                      q8b_4_absentownuse7d_new q8b_5_1a_monthswork12m_cat ///
                      q8b_5_1b_daysmonth12m_cat q8b_5_1c_hrsday12m_cat ///
                      q8b_5_2a_monthsownuse12m_cat q8b_5_2b_daysmonthownuse12m_cat ///
                      q8b_5_2c_hrsdayownuse12m_cat q8b_6_1a_days_pre_cat ///
                      q8b_6_1b_hrs_pre_cat q8b_6_1c_income_pre_cat ///
                      q8c_1_aspiration_new q8c_2_furtherreason_new
asdoc tab1 `workstatus_freq', save(trainee.doc) append




*==============================================================================
* PART B: SUMMARY STATISTICS (mean, sd, min, max) -- CONTINUOUS VARIABLES
*==============================================================================
local continuous_vars q7_age q2a_3_attachdur ///
    q2b_1_1_classwk_org q2b_1_1_classwk_resp ///
    q2b_1_2_weeksheld_org q2b_1_2_weeksheld_resp ///
    q2b_1_3_totalclass_org q2b_1_3_totalclass_resp ///
    q2b_1_4_hrsperclass_org q2b_1_4_hrsperclass_resp ///
    q2b_1_5_totalhrs_org q2b_1_5_totalhrs_resp ///
    q2b_3_ninstructors attendance_rate hours_attendance_rate ///
    q5_2_stipendamt q5_6_1a_fee q5_6_2a_transport q5_6_3a_equip ///
    q5_6_4a_stationery q5_6_5b_othertotal1 q5_6_6b_othertotal2 ///
    q8b_5_1a_monthswork12m q8b_5_1b_daysmonth12m q8b_5_1c_hrsday12m ///
    q8b_5_2a_monthsownuse12m q8b_5_2b_daysmonthownuse12m ///
	q8b_5_2c_hrsdayownuse12m q8b_6_1a_days_pre ///
	q8b_6_1b_hrs_pre q8b_6_1c_income_pre

asdoc sum `continuous_vars', save(trainee.doc) append 



*==============================================================================
* PART C: CROSS-TABULATIONS BY KEY STRATIFIERS 1 (chi2 & ttest)
* Stratifiers: Division, Training Program Type, Sector, Gender, Disability (PWD)
* Applied selectively to thematically relevant outcome variables, not as a
* blanket cross of every stratifier against every variable.
*==============================================================================

* --- C2. Program design / coverage -----------------------------------------------
asdoc tab division_new sector_name_new, row chi2 save(trainee.doc) append
asdoc tab division_new q2_1_progtype_new, row chi2 save(trainee.doc) append
asdoc tab sector_name_new q2_1_progtype_new, row chi2 save(trainee.doc) append
asdoc tab q2_gender_new q2_1_progtype_new, row chi2 save(trainee.doc) append
asdoc tab q2_gender_new sector_name_new, row chi2 save(trainee.doc) append
asdoc tab q3_status_new sector_name_new, row chi2 save(trainee.doc) append


* Age and attendance by gender and by disability status 

asdoc ttest q7_age if inlist(q2_gender_new, 1, 2), ///
    by(q2_gender_new) ///
    save(trainee.doc) append

asdoc ttest q7_age, by(q3_status_new) save(trainee.doc) append 

asdoc ttest attendance_rate if inlist(q2_gender_new, 1, 2), ///
    by(q2_gender_new) ///
    save(trainee.doc) append


asdoc ttest attendance_rate, by(q3_status_new) save(trainee.doc) append 

*==============================================================================
* PART D: CROSS-TABULATIONS BY KEY STRATIFIERS 2 (frequencies and percentages)
* Stratifiers: Division, Training Program Type, Sector, Gender, Disability (PWD)
* Applied selectively to thematically relevant outcome variables, not as a
* blanket cross of every stratifier against every variable.
*==============================================================================
local strat_vars division_new q2_1_progtype_new ///
		sector_name_new q2_gender_new q3_status_new

* --- D1. Demographic composition / targeting equity -----------------------------
local comp_outcomes q2_gender_new ///
			q3_status_new q6_ethnicity_new ///
			q7_age_cat q8a_1_edulevel_new
foreach sv of local strat_vars {
    foreach cv of local comp_outcomes {
        if "`sv'" != "`cv'" {
            asdoc tab `sv' `cv', row col save(trainee.doc) append ///
                title(Table C1: `cv' by `sv')
        }
    }
}

* --- D2. Program design / coverage -----------------------------------------------
asdoc tab division_new sector_name_new, row col save(trainee.doc) append ///
    title(Table C2a: Sector composition by Division)
asdoc tab division_new q2_1_progtype_new, row col save(trainee.doc) append ///
    title(Table C2b: Program type by Division)
asdoc tab sector_name_new q2_1_progtype_new, row col save(trainee.doc) append ///
    title(Table C2c: Program type by Sector)
asdoc tab q2_gender_new q2_1_progtype_new, row col save(trainee.doc) append ///
    title(Table C2d: Program type by Gender)
asdoc tab q2_gender_new sector_name_new, row col save(trainee.doc) append ///
    title(Table C2e: Sector by Gender)
asdoc tab q3_status_new sector_name_new, row col save(trainee.doc) append ///
    title(Table C2f: Sector by Disability status)
asdoc tab q2a_1_attach_new q2_1_progtype_new, row col save(trainee.doc) append ///
    title(Table C2g: Industry attachment uptake by Program type)

* --- D3. Attendance & training-quality perception (continuous outcomes) ----------
* A simple mean satisfaction index across the six 6.3 Likert items, built here
* purely for these summary cross-tabs (not altering existing variables)
capture drop satisfaction_index
egen satisfaction_index = rowmean(q6_3_1_scheduling_new q6_3_2_matavail_new ///
                                  q6_3_3_matquality_new q6_3_4_overallquality_new ///
								  q6_3_5_trainerquality_new q6_3_6_industryfacility_new)
label variable satisfaction_index "Mean of six 6.3 training-satisfaction items (1-5), built for this file"


* Attendance rate by division
asdoc tabstat attendance_rate, ///
    by(division_new) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Attendance Rate by Division)

* Hours attendance rate by Division
asdoc tabstat hours_attendance_rate, ///
    by(division_new) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Hours Attendance Rate by Division)

* Satisfaction index by Division	
asdoc tabstat satisfaction_index, ///
    by(division_new) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Satisfaction Index by Division)	
	
* Attendance rate by Program Type
asdoc tabstat attendance_rate, ///
    by(q2_1_progtype_new) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Attendance Rate by Program Type)
	
* Hours attendance rate by Program Type	
asdoc tabstat hours_attendance_rate, ///
    by(q2_1_progtype_new) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Hours Attendance Rate by Program Type)

	
* Satisfaction index by Program Type
asdoc tabstat satisfaction_index, ///
    by(q2_1_progtype_new) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Satisfaction Index by Program Type)
	

* sector_asdoc (for appearing sector name in doc file)
capture drop sector_asdoc
clonevar sector_asdoc = sector_name_new

label define sector_asdoc_lbl ///
    1 "Agro-Food" ///
    2 "Ceramic" ///
    3 "Construction" ///
    4 "Furniture" ///
    5 "Garment/Textile" ///
    6 "Health" ///
    7 "ICT" ///
    8 "Informal" ///
    9 "Light Engineering" ///
    10 "Tourism/Hospitality", replace

label values sector_asdoc sector_asdoc_lbl	


* Attendance rate by Sector
asdoc tabstat attendance_rate, ///
    by(sector_asdoc) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Attendance Rate by Sector)

* Hours attendance rate by Sector
asdoc tabstat hours_attendance_rate, ///
    by(sector_asdoc) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Hours Attendance Rate by Sector)

* Satisfaction index by Sector
asdoc tabstat satisfaction_index, ///
    by(sector_asdoc) ///
    statistics(mean sd N) ///
    save(trainee.doc) append ///
    title(Table C3: Satisfaction Index by Sector)	
	


	

* --- D4. OHS/PPE and facility compliance (safeguarding equity) -------------------

local strat_vars division_new q2_1_progtype_new ///
		sector_name_new q2_gender_new q3_status_new

local ohs_facility_outcomes q3_1a_ppe_trainee_new ///
	q3_2_ppe_enforced_new q3_3_safetyguard_new ///
    q4_access_new q5_toilet_access_new q4_2_toilets_new q4_4_complaintbox_new
foreach sv of local strat_vars {
    foreach cv of local ohs_facility_outcomes {
        asdoc tab `sv' `cv', row col save(trainee.doc) append ///
            title(Table C4: `cv' by `sv')
    }
}

* --- D5. Financial support and cost-sharing --------------------------------------
local strat_vars division_new q2_1_progtype_new ///
		sector_name_new q2_gender_new q3_status_new

local finance_outcomes q5_1_stipend_new q5_6_1b_feefunder_new
foreach sv of local strat_vars {
    foreach cv of local finance_outcomes {
        asdoc tab `sv' `cv', row col save(trainee.doc) append ///
            title(Table C5: `cv' by `sv')
    }
}

* --- D6. Pre-training baseline / targeting analysis ------------------------------
* Are the most vulnerable / lowest-income groups being reached across
* divisions, sectors, program types, gender and disability status
local strat_vars division_new q2_1_progtype_new /// 
		sector_name_new q2_gender_new q3_status_new

local baseline_outcomes q8a_6_prejobstatus_new ///
		q8a_10_preincome_new q8a_13_primaryearner_new
foreach sv of local strat_vars {
    foreach cv of local baseline_outcomes {
        asdoc tab `sv' `cv', row col save(trainee.doc) append ///
            title(Table C6: `cv' by `sv')
    }
}


*==============================================================================
*MULTIPLE-RESPONSE (MR) FREQUENCY/% TABLES
*============================================
*SCOPE:
    *question sets only:
      *Q12  How did you come to know about this training?
      *Q13  Why did you choose this training?
      *6.2  If attendance was less than 90%, what were the main reasons?
      *5.4  When was the stipend paid?
      *5.5  How did you use your stipend?
    *(Graduate-only multi-selects -- 9A.6, 9C.2, 9D.1 -- are out of scope.)
  *Stratifiers: division_new, q2_1_progtype_new, sector_name_new, q2_gender_new, q3_status_new (disability/PWD status)
  *==============================================================================*/

ssc install mrtab, replace


local strat_vars division_new q2_1_progtype_new sector_name_new q2_gender_new q3_status_new

*==============================================================================
* REUSABLE SUBROUTINE: mr-table for one question set, broken down by every
* category of every stratifier. Tries genuine mrtab first; falls back to an
* equivalent tab1 frequency/% table if asdoc can't capture mrtab's output.
*==============================================================================
capture program drop mrxtab_bygroup
program define mrxtab_bygroup
    syntax varlist, TITLE(string) OUTFILE(string) STRATVARS(string)

    * --- Overall (whole trainee sample) table -------------------------------
    capture noisily asdoc mrtab `varlist', save(trainee.doc) append title(`title' -- Overall)
    if _rc != 0 {
        di as txt "mrtab-via-asdoc not available for [`title' -- Overall]; using tab1 fallback"
        asdoc tab1 `varlist', save(trainee.doc) append title(`title' -- Overall)
    }

    * --- By each stratifier's observed categories ---------------------------
    foreach sv of local stratvars {
        levelsof `sv', local(lvls)
        foreach lv of local lvls {
            local lbl : label (`sv') `lv'
            capture noisily asdoc mrtab `varlist' if `sv'==`lv', ///
                save(trainee.doc) append title(`title' -- `sv' = `lbl')
            if _rc != 0 {
                di as txt "mrtab-via-asdoc not available for [`title' -- `sv' = `lbl']; using tab1 fallback"
                asdoc tab1 `varlist' if `sv'==`lv', ///
                    save(trainee.doc) append title(`title' -- `sv' = `lbl')
            }
        }
    }
end

*==============================================================================
* Q12: How did you come to know about this training?
*==============================================================================
mrxtab_bygroup q12_c01_newspaper q12_c02_govtsite q12_c03_socialmedia q12_c04_workshop ///
    q12_c05_employer q12_c06_wordofmouth q12_c07_cso q12_c08_poster q12_c09_sms q12_c10_other, ///
    title("Table MR1: Q12 How did you come to know about this training") ///
    outfile(trainee.doc) stratvars(`strat_vars')

*==============================================================================
* Q13: Why did you choose this training?
*==============================================================================
mrxtab_bygroup q13_c01_getjob q13_c02_techskills q13_c03_careerprog q13_c04_newtech ///
    q13_c05_family q13_c06_certificate q13_c07_other, ///
    title("Table MR2: Q13 Why did you choose this training") ///
    outfile(trainee.doc) stratvars(`strat_vars')

*==============================================================================
* 6.2: If attendance was less than 90%, what were the main reasons?
*==============================================================================
mrxtab_bygroup q6_2_c01_work q6_2_c02_family q6_2_c03_transport q6_2_c04_illness ///
    q6_2_c05_cancelled q6_2_c06_na q6_2_c07_other, ///
    title("Table MR3: 6.2 Reasons for attendance below 90pct") ///
    outfile(trainee.doc) stratvars(`strat_vars')

*==============================================================================
* 5.4: When was the stipend paid?
*==============================================================================
mrxtab_bygroup q5_4_c01_start q5_4_c02_end q5_4_c03_2wkafter, ///
    title("Table MR4: 5.4 When was the stipend paid") ///
    outfile(trainee.doc) stratvars(`strat_vars')

*==============================================================================
* 5.5: How did you use your stipend?
*==============================================================================

local strat_vars division_new q2_1_progtype_new sector_name_new q2_gender_new q3_status_new

mrxtab_bygroup q5_5_c01_food q5_5_c02_transport q5_5_c03_mobiledata q5_5_c04_household ///
    q5_5_c05_savings q5_5_c06_other, ///
    title("Table MR5: 5.5 How did you use your stipend") ///
    outfile(trainee.doc) stratvars(`strat_vars')


**--------------------------------------------
*			Saving the data as SPSS file
**--------------------------------------------
ssc install savespss

savespss "I:\15. Data Set\Clean\trainee.sav", replace


```


---



## For Institution Survey Data


```stata
cls
clear all

cd "I:\15. Data Set\Clean"

cd "D:\Munem\All Projects\2026\DTE\15. Data Set\Clean"


* --- 0. Load raw data --------------------------------------------------

import excel using "Institutions_Merged", sheet("data") firstrow clear



* 1. New variable names, in EXACT column order ----------------------

rename	NameofEnumera~r		enum_name
rename	NameoftheResp~t		resp_name
rename	Designationof~t		resp_designation
rename	CellNumberoft~t		resp_phone
rename	E					q1_1_inst_name
rename	Division			division
rename	G					q1_2_stp_registered
rename	H					q1_3_stp_reg_year
rename	I	 				q1_4_stp_name
rename	J					q1_5_ind_partner_name
rename	K					q1_6_year_established
rename	L               	q1_7_sector
rename	Other           	q1_7_other
rename	N               	q1_8_avg_workforce_size
rename	O               	q1_9_staff_involved_n
rename	P               	q1_10_know_source
rename	Other2          	q1_10_other
rename	R               	q1_11_grp
rename	S               	q1_11_month
rename	T               	q1_11_year
rename	U               	q1_12_grp
rename	V               	q1_12_month
rename	W               	q1_12_year
rename	X               	q1_13_grp
rename	Y               	q1_13_tstp
rename	Z               	q1_13_adtp
rename	AA              	q1_13_sutp
rename	AB              	q1_13_ihtp
rename	AC              	q1_14_grp
rename	AD              	q1_14_1_tstp
rename	AE              	q1_14_1a_ongoing
rename	AF              	q1_14_1b_completed
rename	AG              	q1_14_1c_assessed
rename	AH              	q1_14_2_adtp
rename	AI              	q1_14_2a_ongoing
rename	AJ              	q1_14_2b_completed
rename	AK              	q1_14_2c_assessed
rename	AL              	q1_14_3_sutp
rename	AM              	q1_14_3a_ongoing
rename	AN              	q1_14_3b_completed
rename	AO              	q1_14_3c_assessed
rename	AP              	q1_14_4_ihtp
rename	AQ              	q1_14_4a_ongoing
rename	AR              	q1_14_4b_completed
rename	AS              	q1_14_4c_assessed
rename	AT              	q1_15_grp
rename	AU              	q1_15_1_tstp
rename	AV              	q1_15_1a_assessed
rename	AW              	q1_15_1b_competent
rename	AX              	q1_15_1c_eligible
rename	AY              	q1_15_1d_employed
rename	AZ              	q1_15_2_adtp
rename	BA              	q1_15_2a_assessed
rename	BB              	q1_15_2b_competent
rename	BC              	q1_15_2c_eligible
rename	BD              	q1_15_2d_employed
rename	BE              	q1_15_3_sutp
rename	BF              	q1_15_3a_assessed
rename	BG              	q1_15_3b_competent
rename	BH              	q1_15_3c_eligible
rename	BI              	q1_15_3d_employed
rename	BJ              	q1_15_4_ihtp
rename	BK              	q1_15_4a_assessed
rename	BL              	q1_15_4b_competent
rename	BM              	q1_15_4c_eligible
rename	BN              	q1_15_4d_employed
rename	BO              	q1_16_grp
rename	First           	q1_16_1_grp
rename	BQ              	q1_16_1a_trade
rename	BR              	q1_16_1b_level
rename	BS              	q1_16_1c_n_trained
rename	Second          	q1_16_2_grp
rename	BU              	q1_16_2a_trade
rename	BV              	q1_16_2b_level
rename	BW              	q1_16_2c_n_trained
rename	Third           	q1_16_3_grp
rename	BY              	q1_16_3a_trade
rename	BZ              	q1_16_3b_level
rename	CA              	q1_16_3c_n_trained
rename	Fourth          	q1_16_4_grp
rename	CC              	q1_16_4a_trade
rename	CD              	q1_16_4b_level
rename	CE              	q1_16_4c_n_trained
rename	Fifth           	q1_16_5_grp
rename	CG              	q1_16_5a_trade
rename	CH              	q1_16_5b_level
rename	CI              	q1_16_5c_n_trained
rename	Sixth           	q1_16_6_grp
rename	CK              	q1_16_6a_trade
rename	CL              	q1_16_6b_level
rename	CM              	q1_16_6c_n_trained
rename	CN              	q1_17_grp
rename	First2          	q1_17_1_grp
rename	CP              	q1_17_1a_trade
rename	CQ              	q1_17_1b_reason
rename	CR              	q1_17_1c_est_needed
rename	Second3         	q1_17_2_grp
rename	CT              	q1_17_2a_trade
rename	CU              	q1_17_2b_reason
rename	CV              	q1_17_2c_est_needed
rename	Third4          	q1_17_3_grp
rename	CX              	q1_17_3a_trade
rename	CY              	q1_17_3b_reason
rename	CZ              	q1_17_3c_est_needed
rename	Fourth5         	q1_17_4_grp
rename	DB              	q1_17_4a_trade
rename	DC              	q1_17_4b_reason
rename	DD              	q1_17_4c_est_needed
rename	DE              	q1_18_skills_gap
rename	DF              	q1_19_skills_gap_detail
rename	DG              	q2_1_grp 
rename	DH              	q2_1a_resp
rename	DI              	q2_1b_obs
rename	DJ              	q2_2_grp
rename	DK              	q2_2a_resp
rename	DL              	q2_2b_obs
rename	DM              	q2_3_grp
rename	DN              	q2_3a_resp
rename	DO              	q2_3b_obs
rename	DP              	q2_4_grp
rename	DQ              	q2_4a_resp
rename	DR              	q2_4b_obs
rename	DateofVisit     	q2_5_date
rename	OfficialTime    	q2_5_time
rename	DU              	q2_5_1_grp
rename	DV              	q2_5_1a_male
rename	DW              	q2_5_1b_female
rename	DX              	q2_5_1c_total
rename	DY              	q2_5_2_grp
rename	DZ              	q2_5_2a_male
rename	EA              	q2_5_2b_female
rename	EB              	q2_5_2c_total
rename	EC              	q2_5_3_grp
rename	ED              	q2_5_3a_ethnic_minority
rename	EE              	q2_5_3b_pwd
rename	EF              	q2_5_4_grp
rename	EG              	q2_5_4a_male
rename	EH              	q2_5_4b_female
rename	EI              	q2_5_4c_total
rename	EJ              	q2_5_5_grp
rename	EK              	q2_5_5a_male
rename	EL              	q2_5_5b_female
rename	EM              	q2_5_5c_total
rename	EN              	q2_5_6_grp
rename	EO              	q2_5_6a_ethnic_minority
rename	EP              	q2_5_6b_pwd
rename	EQ              	q2_5_7_grp
rename	SSC             	q2_5_7_ssc
rename	HSC             	q2_5_7_hsc
rename	Diploma         	q2_5_7_diploma
rename	Bachelororabove	q2_5_7_bachelor_above
rename	EV              	q3_1_grp
rename	EW              	q3_1a_resp
rename	EX              	q3_1b_obs
rename	EY              	q3_2_grp
rename	EZ              	q3_2a_resp
rename	FA              	q3_2b_obs
rename	FB              	q3_3_grp
rename	FC              	q3_3a_resp
rename	FD              	q3_3b_obs
rename	FE             		q3_4_grp
rename	FF              	q3_4a_resp
rename	Other7          	q3_4a_other
rename	FH              	q3_4b_obs
rename	Other8          	q3_4b_other
rename	FJ              	q3_5_grp
rename	FK              	q3_5a_resp
rename	FL              	q3_5b_obs
rename	FM              	q3_6_grp
rename	FN              	q3_6a_resp
rename	FO              	q3_6b_obs
rename	FP              	q3_7_grp
rename	FQ              	q3_7a_resp
rename	FR              	q3_7b_obs
rename	FS              	q3_8_grp
rename	FT              	q3_8a_resp
rename	FU              	q3_8b_obs
rename	FV              	q3_9_grp
rename	FW              	q3_9a_resp
rename	FX              	q3_9b_obs
rename	FY              	q4_1_grp
rename	FZ              	q4_1a_resp
rename	GA              	q4_1b_obs
rename	GB              	q4_2_grp
rename	GC              	q4_2a_resp
rename	GD              	q4_2b_obs
rename	GE              	q4_3_grp
rename	GF              	q4_3a_resp
rename	GG              	q4_3b_obs
rename	GH              	q4_4_grp
rename	GI              	q4_4a_resp
rename	GJ              	q4_4b_obs
rename	GK              	q4_5_grp
rename	GL              	q4_5a_resp
rename	GM              	q4_5b_obs
rename	GN              	q5_1_grp
rename	GO              	q5_1a_resp
rename	GP              	q5_1b_obs
rename	GQ              	q5_2_grp
rename	GR              	q5_2a_resp
rename	GS              	q5_2b_obs
rename	GT              	q5_3_grp
rename	GU              	q5_3a_resp
rename	GV              	q5_3b_obs
rename	GW              	q6_1_fee_received
rename	GX              	q6_2_grp
rename	GY              	q6_2_1_grp
rename	stTrancheNoof~s		q6_2_1_t1_n
rename	stTrancheAmount		q6_2_1_t1_amt
rename	ndTrancheNoof~s		q6_2_1_t2_n
rename	ndTrancheAmount		q6_2_1_t2_amt
rename	rdTrancheNoof~s		q6_2_1_t3_n
rename	rdTrancheAmount		q6_2_1_t3_amt
rename	thTrancheNoof~s		q6_2_1_t4_n
rename	thTrancheAmount		q6_2_1_t4_amt
rename	HH              	q6_2_2_grp
rename	stTrancheNoof~9		q6_2_2_t1_n
rename	stTrancheAmo~10		q6_2_2_t1_amt
rename	ndTrancheNoo~11		q6_2_2_t2_n
rename	ndTrancheAmo~12		q6_2_2_t2_amt
rename	rdTrancheNoo~13		q6_2_2_t3_n
rename	rdTrancheAmo~14		q6_2_2_t3_amt
rename	thTrancheNoo~15		q6_2_2_t4_n
rename	thTrancheAmo~16		q6_2_2_t4_amt
rename	HQ              	q6_2_3_grp
rename	stTrancheNoo~17		q6_2_3_t1_n
rename	stTrancheAmo~18		q6_2_3_t1_amt
rename	ndTrancheNoo~19		q6_2_3_t2_n
rename	ndTrancheAmo~20		q6_2_3_t2_amt
rename	rdTrancheNoo~21		q6_2_3_t3_n
rename	rdTrancheAmo~22		q6_2_3_t3_amt
rename	thTrancheNoo~23		q6_2_3_t4_n
rename	thTrancheAmo~24		q6_2_3_t4_amt
rename	HZ              	q6_2_4_grp
rename	stTrancheNoo~25		q6_2_4_t1_n
rename	stTrancheAmo~26		q6_2_4_t1_amt
rename	ndTrancheNoo~27		q6_2_4_t2_n
rename	ndTrancheAmo~28		q6_2_4_t2_amt
rename	rdTrancheNoo~29		q6_2_4_t3_n
rename	rdTrancheAmo~30		q6_2_4_t3_amt
rename	thTrancheNoo~31		q6_2_4_t4_n
rename	thTrancheAmo~32		q6_2_4_t4_amt
rename	II              	q6_3_grp
rename	General         	q6_3_general
rename	Female          	q6_3_female
rename	PWD             	q6_3_pwd
rename	IM              	q6_3_em
rename	Total           	q6_3_total
rename	IO              	q6_4_grp
rename	IP              	q6_4_doc_delay
rename	IQ              	 q6_4_emp_track
rename	IR              	q6_4_dropout
rename	IS              	q6_4_admin_delay
rename	IT              	q6_4_payment_issue
rename	IU              	q6_4_not_eligible
rename	IV              	q6_4_not_informed
rename	IW              	q6_4_other_flag
rename	Other33         	q6_4_other_text
rename	IY              	q6_5_grp
rename	IZ              	q6_5a_resp
rename	JA              	q6_5b_obs
rename	JB              	q6_6_grp
rename	JC              	q6_6a_resp
rename	JD              	q6_6b_obs
rename	JE              	q6_7_grp
rename	JF              	q6_7a_resp
rename	JG              	q6_7b_obs
rename	JH              	q7_1_industry_contrib_type
rename	JI              	q7_2_cash_amt
rename	JJ              	q7_3_grp
rename	JK              	q7_3_training_space
rename	JL              	q7_3_factory_utilities
rename	JM              	q7_3_machinery_use
rename	JN              	q7_3_trainer_time
rename	JO              	q7_3_coord_salaries
rename	JP              	q7_3_raw_materials
rename	JQ              	q7_3_admin_support
rename	JR              	q7_3_other_flag
rename	Other34         	q7_3_other_text
rename	JT              	q7_4_inkind_value_bdt
rename	JU              	q8_1_other_training_running
rename	JV              	q8_2_grp
rename	JW              	q8_2_tvet
rename	JX              	q8_2_apprenticeship
rename	JY              	q8_2_bizmgmt
rename	JZ              	q8_2_entrepreneurship
rename	KA              	q8_2_finance
rename	KB              	q8_2_digital
rename	KC              	q8_2_softskills
rename	KD              	q8_2_other_flag
rename	Other35         	q8_2_other_text
rename	KF              	q8_3_1a_type
rename	KG              	q8_3_1b_funding
rename	KH              	q8_3_1c_n
rename	KI              	q8_3_2a_type
rename	KJ              	q8_3_2b_funding
rename	KK              	q8_3_2c_n
rename	KL              	q8_3_3a_type
rename	KM              	q8_3_3b_funding
rename	KN              	q8_3_3c_n
rename	KO              	q8_3_4a_type
rename	KP              	q8_3_4b_funding
rename	KQ              	q8_3_4c_n
rename	KR              	q8_3_5a_type
rename	KS              	q8_3_5b_funding
rename	KT              	q8_3_5c_n
rename	KU              	q8_3_6a_type
rename	KV              	q8_3_6b_funding
rename	KW              	q8_3_6c_n
rename	KX              	q8_3_7a_type
rename	KY              	q8_3_7b_funding
rename	KZ              	q8_3_7c_n
rename	LA              	q8_3_8a_type
rename	LB              	q8_3_8b_funding
rename	LC              	q8_3_8c_n
rename	LD              	q8_4_grp
rename	LE              	q8_4_dev_partner
rename	LF              	q8_4_own_revenue
rename	LG              	q8_4_fee_based
rename	LH              	q8_4_govt_grant
rename	LI              	q8_4_csr
rename	LJ              	q8_4_other_flag
rename	OtherDevelopm~s		q8_4_dev_partner_text
rename	Otherspecify    	q8_4_other_text
rename	LM              	q9_1_grp
rename	LN              	q9_1_trainee_engagement
rename	LO              	q9_1_attendance_retention
rename	LP              	q9_1_lack_trainers
rename	LQ              	q9_1_coordination
rename	LR              	q9_1_cblm_quality
rename	LS              	q9_1_equipment_cost
rename	LT              	q9_1_compliance
rename	LU              	q9_1_employment_tracking
rename	LV              	q9_1_payment_delay
rename	LW              	q9_1_gender_inclusion
rename	LX              	q9_1_ohs
rename	LY              	q9_1_other_flag
rename	Otherspecify36  	q9_1_other_text
rename	MA              	q9_2_other_challenge_desc
rename	MB              	q10_1_continue_ebt
rename	MC              	q10_2_grp
rename	MD              	q10_2_industry_partnership
rename	ME              	q10_2_industry_costshare
rename	MF              	q10_2_govt_incentive
rename	MG              	q10_2_cblm_capacity
rename	MH              	q10_2_flex_schedule
rename	MI              	q10_2_curriculum_update
rename	MJ              	q10_2_shared_facility
rename	MK              	q10_2_placement_support
rename	ML              	q10_2_other_flag
rename	Otherspecify37  	q10_2_other_text
rename	MN              	q10_3_sustain_support_needed
rename	MO              	q11_1_willing_increase_contrib
rename	MP              	q11_2_contrib_share_pref
rename	Otherspecify38  	q11_2_other_text
rename	MR              	q11_3_contrib_form
rename	MS              	q11_4_grp
rename	MT              	q11_4_outcomes_demo
rename	MU              	q11_4_govt_incentive
rename	MV              	q11_4_admin_burden_reduce
rename	MW              	q11_4_flex_modality
rename	MX              	q11_4_certification_access
rename	MY              	q11_4_other_flag
rename	Otherspecify39  	q11_4_other_text
rename	NA              	q11_5_cofinance_comments
rename	Pleasecollect~n		gps_grp
rename	_Pleasecollec~o		gps_lat
rename	ND					gps_lon
rename	NE					gps_alt
rename	NF					gps_precision


* --- 4. Variable labels (original question wording, <=80 chars)
label variable enum_name "Name of Enumeraotor"
label variable resp_name "Name of the Respondent"
label variable resp_designation "Designation of the Respondent"
label variable resp_phone "Cell Number of the Respondent"
label variable q1_1_inst_name "1.1 Name of the Institution/Industry"
label variable division "Division"
label variable q1_2_stp_registered "1.2 Is this registered as an Industrial STP under NSDA?"
label variable q1_3_stp_reg_year "1.3 If yes, which year was it registered as an Industrial STP under NSDA?"
label variable q1_4_stp_name "1.4 Name of the STP"
label variable q1_5_ind_partner_name "1.5 Name of the Industry Partner (ISC/Association/Chamber of Commerce)"
label variable q1_6_year_established "1.6 Year of Establishment (this can be different from the registration date)"
label variable q1_7_sector "1.7 Which primary industry sector does this institution belong to?"
label variable q1_7_other "Other"
label variable q1_8_avg_workforce_size "1.8 Average Size: Average size of the factory/enterprise (based on total workfo"
label variable q1_9_staff_involved_n "1.9 How many staff from the industry/factory are directly involved in this trai"
label variable q1_10_know_source "1.10 How did you know about the EBT of ASSET Project?"
label variable q1_10_other "Other2"
label variable q1_11_grp "1.11 When did you apply for conducting the EBT to ASSET Project?"
label variable q1_11_month "1.11.1 Month"
label variable q1_11_year "1.11.2 Year"
label variable q1_12_grp "1.12 When did you sign the contract with ASSET Project for conducting the EBT?"
label variable q1_12_month "1.12.1 Month"
label variable q1_12_year "1.12.2 Year"
label variable q1_13_grp "1.13 Which EBT program types under ASSET are implemented by the IP?"
label variable q1_13_tstp "1.13 Which EBT program types under ASSET are implemented by the IP?/TSTP"
label variable q1_13_adtp "1.13 Which EBT program types under ASSET are implemented by the IP?/ADTP"
label variable q1_13_sutp "1.13 Which EBT program types under ASSET are implemented by the IP?/SUTP"
label variable q1_13_ihtp "1.13 Which EBT program types under ASSET are implemented by the IP?/IHTP"
label variable q1_14_grp "1.14 How many batches of TSTP, ADTP, SUTP, and IHTP has your institute complete"
label variable q1_14_1_tstp "1.14.1 TSTP"
label variable q1_14_1a_ongoing "1.14.1a Ongoing (No.)"
label variable q1_14_1b_completed "1.14.1b Completed (No.)"
label variable q1_14_1c_assessed "1.14.1c Assessed (No.)"
label variable q1_14_2_adtp "1.14.2 ADTP"
label variable q1_14_2a_ongoing "1.14.2a Ongoing (No.)"
label variable q1_14_2b_completed "1.14.2b Completed (No.)"
label variable q1_14_2c_assessed "1.14.2c Assessed (No.)"
label variable q1_14_3_sutp "1.14.3 SUTP"
label variable q1_14_3a_ongoing "1.14.3a Ongoing (No.)"
label variable q1_14_3b_completed "1.14.3b Completed (No.)"
label variable q1_14_3c_assessed "1.14.3c Assessed (No.)"
label variable q1_14_4_ihtp "1.14.4 IHTP"
label variable q1_14_4a_ongoing "1.14.4a Ongoing (No.)"
label variable q1_14_4b_completed "1.14.4b Completed (No.)"
label variable q1_14_4c_assessed "1.14.4c Assessed (No.)"
label variable q1_15_grp "1.15 Please provide details on number of trainees that are already assessed and"
label variable q1_15_1_tstp "1.15.1 TSTP"
label variable q1_15_1a_assessed "1.15.1a Assessed (no. of Trainees)"
label variable q1_15_1b_competent "1.15.1b Competent"
label variable q1_15_1c_eligible "1.15.1c Number of Eligible Trainees for employment (competent for at least 3 mo"
label variable q1_15_1d_employed "1.15.1d Number of employed of those who are competent for at least 3 months"
label variable q1_15_2_adtp "1.15.2 ADTP"
label variable q1_15_2a_assessed "1.15.2a Assessed (no. of Trainees)"
label variable q1_15_2b_competent "1.15.2b Competent"
label variable q1_15_2c_eligible "1.15.2c Number of Eligible Trainees for employment (competent for at least 3 mo"
label variable q1_15_2d_employed "1.15.2d Number of employed of those who are competent for at least 3 months"
label variable q1_15_3_sutp "1.15.3 SUTP"
label variable q1_15_3a_assessed "1.15.3a Assessed (no. of Trainees)"
label variable q1_15_3b_competent "1.15.3b Competent"
label variable q1_15_3c_eligible "1.15.3c Number of Eligible Trainees for employment (competent for at least 3 mo"
label variable q1_15_3d_employed "1.15.3d Number of employed of those who are competent for at least 3 months"
label variable q1_15_4_ihtp "1.15.4 IHTP"
label variable q1_15_4a_assessed "1.15.4a Assessed (no. of Trainees)"
label variable q1_15_4b_competent "1.15.4b Competent"
label variable q1_15_4c_eligible "1.15.4c Number of Eligible Trainees for employment (competent for at least 3 mo"
label variable q1_15_4d_employed "1.15.4d Number of employed of those who are competent for at least 3 months"
label variable q1_16_grp "1.16 List the specific occupations and NTVQF/Level offered"
label variable q1_16_1_grp "First"
label variable q1_16_1a_trade "1.16.1a Trade / Occupation Name"
label variable q1_16_1b_level "1.16.1b Level"
label variable q1_16_1c_n_trained "1.16.1c Number of Trained"
label variable q1_16_2_grp "Second"
label variable q1_16_2a_trade "1.16.2a Trade / Occupation Name"
label variable q1_16_2b_level "1.16.2b Level"
label variable q1_16_2c_n_trained "1.16.2c Number of Trained"
label variable q1_16_3_grp "Third"
label variable q1_16_3a_trade "1.16.3a Trade / Occupation Name"
label variable q1_16_3b_level "1.16.3b Level"
label variable q1_16_3c_n_trained "1.16.3c Number of Trained"
label variable q1_16_4_grp "Fourth"
label variable q1_16_4a_trade "1.16.4a Trade / Occupation Name"
label variable q1_16_4b_level "1.16.4b Level"
label variable q1_16_4c_n_trained "1.16.4c Number of Trained"
label variable q1_16_5_grp "Fifth"
label variable q1_16_5a_trade "1.16.5a Trade / Occupation Name"
label variable q1_16_5b_level "1.16.5b Level"
label variable q1_16_5c_n_trained "1.16.5c Number of Trained"
label variable q1_16_6_grp "Sixth"
label variable q1_16_6a_trade "1.16.6a Trade / Occupation Name"
label variable q1_16_6b_level "1.16.6b Level"
label variable q1_16_6c_n_trained "1.16.6c Number of Trained"
label variable q1_17_grp "1.17 Emerging Trades / Occupations Needed (Next 2–3 Years): What new trades or "
label variable q1_17_1_grp "First"
label variable q1_17_1a_trade "1.17.1a Trade / Occupation Needed"
label variable q1_17_1b_reason "1.17.1b Reason / Demand Driver"
label variable q1_17_1c_est_needed "1.17.1c Est. No. Needed"
label variable q1_17_2_grp "Second"
label variable q1_17_2a_trade "1.17.2a Trade / Occupation Needed"
label variable q1_17_2b_reason "1.17.2b Reason / Demand Driver"
label variable q1_17_2c_est_needed "1.17.2c Est. No. Needed"
label variable q1_17_3_grp "Third"
label variable q1_17_3a_trade "1.17.3a Trade / Occupation Needed"
label variable q1_17_3b_reason "1.17.3b Reason / Demand Driver"
label variable q1_17_3c_est_needed "1.17.3c Est. No. Needed"
label variable q1_17_4_grp "Fourth"
label variable q1_17_4a_trade "1.17.4a Trade / Occupation Needed"
label variable q1_17_4b_reason "1.17.4b Reason / Demand Driver"
label variable q1_17_4c_est_needed "1.17.4c Est. No. Needed"
label variable q1_18_skills_gap "1.18 Are there any specific technical skills gap your sector faces?"
label variable q1_19_skills_gap_detail "1.19 If yes, please provide specific technical skills gap your sector faces."
label variable q2_1_grp "2.1 Is the training venue spacious enough to accommodate all trainees comfortab"
label variable q2_1a_resp "2.1a Institutional Response"
label variable q2_1b_obs "2.1b Physically Observed by Enumerator"
label variable q2_2_grp "2.2 Are manual attendance registers maintained for all trainees?"
label variable q2_2a_resp "2.2a Institutional Response"
label variable q2_2b_obs "2.2b Physically Observed by Enumerator"
label variable q2_3_grp "2.3 Are biometric systems (finger/facial) currently operational? (for trainees)"
label variable q2_3a_resp "2.3a Institutional Response"
label variable q2_3b_obs "2.3b Physically Observed by Enumerator"
label variable q2_4_grp "2.4 Are registration data and milestones uploaded to the TMIS?"
label variable q2_4a_resp "2.4a Institutional Response"
label variable q2_4b_obs "2.4b Physically Observed by Enumerator"
label variable q2_5_date "Date of Visit"
label variable q2_5_time "Official Time"
label variable q2_5_1_grp "2.5.1 Total Participants Enrolled"
label variable q2_5_1a_male "2.5.1a Male"
label variable q2_5_1b_female "2.5.1b Female"
label variable q2_5_1c_total "2.5.1c Total"
label variable q2_5_2_grp "2.5.2 NEET People Enrolled"
label variable q2_5_2a_male "2.5.2a Male"
label variable q2_5_2b_female "2.5.2b Female"
label variable q2_5_2c_total "2.5.2c Total"
label variable q2_5_3_grp "2.5.3 Total Participants Enrolled"
label variable q2_5_3a_ethnic_minority "2.5.3a Ethnic Minority"
label variable q2_5_3b_pwd "2.5.3b Person with Disability"
label variable q2_5_4_grp "2.5.4 Total Participant Present on the day of visit"
label variable q2_5_4a_male "2.5.4a Male"
label variable q2_5_4b_female "2.5.4b Female"
label variable q2_5_4c_total "2.5.4c Total"
label variable q2_5_5_grp "2.5.5 NEET People Present on the day of visit"
label variable q2_5_5a_male "2.5.5a Male"
label variable q2_5_5b_female "2.5.5b Female"
label variable q2_5_5c_total "2.5.5c Total"
label variable q2_5_6_grp "2.5.6 Present on the day of visit"
label variable q2_5_6a_ethnic_minority "2.5.6a Ethnic Minority"
label variable q2_5_6b_pwd "2.5.6b Person with Disability"
label variable q2_5_7_grp "2.5.7 Participants educational qualification (Nos)"
label variable q2_5_7_ssc "<SSC"
label variable q2_5_7_hsc "<HSC"
label variable q2_5_7_diploma "<Diploma"
label variable q2_5_7_bachelor_above "Bachelor or above"
label variable q3_1_grp "3.1 Are there two unique instructors (one Factory, one Industry/Institute)?"
label variable q3_1a_resp "3.1a Institutional Response"
label variable q3_1b_obs "3.1b Physically Observed by Enumerator"
label variable q3_2_grp "3.2 Are instructors certified at CBT&A Level 4 or above?"
label variable q3_2a_resp "3.2a Institutional Response"
label variable q3_2b_obs "3.2b Physically Observed by Enumerator"
label variable q3_3_grp "3.3 Have trainees received the CBLM hard-copy package?"
label variable q3_3a_resp "3.3a Institutional Response"
label variable q3_3b_obs "3.3b Physically Observed by Enumerator"
label variable q3_4_grp "3.4 CBLM Origin: Who developed these materials (NSDA/BTEB/Other)?"
label variable q3_4a_resp "3.4a Institutional Response"
label variable q3_4a_other "Other7"
label variable q3_4b_obs "3.4b Physically Observed by Enumerator"
label variable q3_4b_other "Other8"
label variable q3_5_grp "3.5 Is the Approach a mix of theoretical and practical hands-on?"
label variable q3_5a_resp "3.5a Institutional Response"
label variable q3_5b_obs "3.5b Physically Observed by Enumerator"
label variable q3_6_grp "3.6 Is the training curriculum aligned with the relevant Competency Standards ("
label variable q3_6a_resp "3.6a Institutional Response"
label variable q3_6b_obs "3.6b Physically Observed by Enumerator"
label variable q3_7_grp "3.7 Are training facilities equipped with the necessary tools, equipment, and m"
label variable q3_7a_resp "3.7a Institutional Response"
label variable q3_7b_obs "3.7b Physically Observed by Enumerator"
label variable q3_8_grp "3.8 Do trainees have adequate access to learning resources during training?"
label variable q3_8a_resp "3.8a Institutional Response"
label variable q3_8b_obs "3.8b Physically Observed by Enumerator"
label variable q3_9_grp "3.9 How many MoUs have been done with Industry for Apprenticeships by the STP ?"
label variable q3_9a_resp "3.9a Institutional Response"
label variable q3_9b_obs "3.9b Physically Observed by Enumerator"
label variable q4_1_grp "4.1 Is Personal Protective Equipment (PPE) provided to all trainees?"
label variable q4_1a_resp "4.1a Institutional Response"
label variable q4_1b_obs "4.1b Physically Observed by Enumerator"
label variable q4_2_grp "4.2 Are safety guards and emergency stops installed on machinery?"
label variable q4_2a_resp "4.2a Institutional Response"
label variable q4_2b_obs "4.2b Physically Observed by Enumerator"
label variable q4_3_grp "4.3 Are trainees briefed on occupational hazards?"
label variable q4_3a_resp "4.3a Institutional Response"
label variable q4_3b_obs "4.3b Physically Observed by Enumerator"
label variable q4_4_grp "4.4 Is clean drinking water provided at the venue?"
label variable q4_4a_resp "4.4a Institutional Response"
label variable q4_4b_obs "4.4b Physically Observed by Enumerator"
label variable q4_5_grp "4.5 Are there gender-separated toilets for trainees?"
label variable q4_5a_resp "4.5a Institutional Response"
label variable q4_5b_obs "4.5b Physically Observed by Enumerator"
label variable q5_1_grp "5.1 Is there a functional Grievance Redress Mechanism (GRM)?"
label variable q5_1a_resp "5.1a Institutional Response"
label variable q5_1b_obs "5.1b Physically Observed by Enumerator"
label variable q5_2_grp "5.2 Is the Complaint Box visible and accessible to trainees?"
label variable q5_2a_resp "5.2a Institutional Response"
label variable q5_2b_obs "5.2b Physically Observed by Enumerator"
label variable q5_3_grp "5.3 Is the training venue/washroom accessible for PWDs (ramps, etc.)?"
label variable q5_3a_resp "5.3a Institutional Response"
label variable q5_3b_obs "5.3b Physically Observed by Enumerator"
label variable q6_1_fee_received "6.1 Did your institution receive a per-trainee fee from the ASSET project?"
label variable q6_2_grp "6.2 Please provide the cycle wise amount"
label variable q6_2_1_grp "6.2.1 First Cycle (1st)"
label variable q6_2_1_t1_n "1st Tranche - No. of Trainees"
label variable q6_2_1_t1_amt "1st Tranche - Amount"
label variable q6_2_1_t2_n "2nd Tranche - No. of Trainees"
label variable q6_2_1_t2_amt "2nd Tranche - Amount"
label variable q6_2_1_t3_n "3rd Tranche - No. of Trainees"
label variable q6_2_1_t3_amt "3rd Tranche - Amount"
label variable q6_2_1_t4_n "4th Tranche - No. of Trainees"
label variable q6_2_1_t4_amt "4th Tranche - Amount"
label variable q6_2_2_grp "6.2.2 Second Cycle (2nd)"
label variable q6_2_2_t1_n "1st Tranche - No. of Trainees9"
label variable q6_2_2_t1_amt "1st Tranche - Amount10"
label variable q6_2_2_t2_n "2nd Tranche - No. of Trainees11"
label variable q6_2_2_t2_amt "2nd Tranche - Amount12"
label variable q6_2_2_t3_n "3rd Tranche - No. of Trainees13"
label variable q6_2_2_t3_amt "3rd Tranche - Amount14"
label variable q6_2_2_t4_n "4th Tranche - No. of Trainees15"
label variable q6_2_2_t4_amt "4th Tranche - Amount16"
label variable q6_2_3_grp "6.2.3 Third Cycle (3rd)"
label variable q6_2_3_t1_n "1st Tranche - No. of Trainees17"
label variable q6_2_3_t1_amt "1st Tranche - Amount18"
label variable q6_2_3_t2_n "2nd Tranche - No. of Trainees19"
label variable q6_2_3_t2_amt "2nd Tranche - Amount20"
label variable q6_2_3_t3_n "3rd Tranche - No. of Trainees21"
label variable q6_2_3_t3_amt "3rd Tranche - Amount22"
label variable q6_2_3_t4_n "4th Tranche - No. of Trainees23"
label variable q6_2_3_t4_amt "4th Tranche - Amount24"
label variable q6_2_4_grp "6.2.4 Fourth Cycle (4th)"
label variable q6_2_4_t1_n "1st Tranche - No. of Trainees25"
label variable q6_2_4_t1_amt "1st Tranche - Amount26"
label variable q6_2_4_t2_n "2nd Tranche - No. of Trainees27"
label variable q6_2_4_t2_amt "2nd Tranche - Amount28"
label variable q6_2_4_t3_n "3rd Tranche - No. of Trainees29"
label variable q6_2_4_t3_amt "3rd Tranche - Amount30"
label variable q6_2_4_t4_n "4th Tranche - No. of Trainees31"
label variable q6_2_4_t4_amt "4th Tranche - Amount32"
label variable q6_3_grp "6.3 Provide Participant Wise total per-trainee fee"
label variable q6_3_general "General"
label variable q6_3_female "Female"
label variable q6_3_pwd "PWD"
label variable q6_3_em "EM"
label variable q6_3_total "Total"
label variable q6_4_grp "6.4 If the full payment/all tranches were not received, what were the reasons?"
label variable q6_4_doc_delay "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_emp_track "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_dropout "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_admin_delay "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_payment_issue "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_not_eligible "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_not_informed "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_other_flag "6.4 If the full payment/all tranches were not received, what were the reasons?/"
label variable q6_4_other_text "Other33"
label variable q6_5_grp "6.5 Is there a dedicated Operational Account (OA) for ASSET?"
label variable q6_5a_resp "6.5a Institutional Response"
label variable q6_5b_obs "6.5b Physically Observed by Enumerator"
label variable q6_6_grp "6.6 Are stipends paid via Bank/Mobile Transfer (Nagad/Bkash)?"
label variable q6_6a_resp "6.6a Institutional Response"
label variable q6_6b_obs "6.6b Physically Observed by Enumerator"
label variable q6_7_grp "6.7 Does the SPMU include female members and experts?"
label variable q6_7a_resp "6.7a Institutional Response"
label variable q6_7b_obs "6.7b Physically Observed by Enumerator"
label variable q7_1_industry_contrib_type "7.1 How did your institution fulfil the required 15% industry contribution?"
label variable q7_2_cash_amt "7.2 If financial, total amount contributed (BDT)"
label variable q7_3_grp "7.3 If in-kind, specify type of contribution (Select all that apply)"
label variable q7_3_training_space "7.3 If in-kind, specify type of contribution (Select all that apply)/Training s"
label variable q7_3_factory_utilities "7.3 If in-kind, specify type of contribution (Select all that apply)/Factory ut"
label variable q7_3_machinery_use "7.3 If in-kind, specify type of contribution (Select all that apply)/Machinery/"
label variable q7_3_trainer_time "7.3 If in-kind, specify type of contribution (Select all that apply)/Industry T"
label variable q7_3_coord_salaries "7.3 If in-kind, specify type of contribution (Select all that apply)/Coordinati"
label variable q7_3_raw_materials "7.3 If in-kind, specify type of contribution (Select all that apply)/Raw materi"
label variable q7_3_admin_support "7.3 If in-kind, specify type of contribution (Select all that apply)/Administra"
label variable q7_3_other_flag "7.3 If in-kind, specify type of contribution (Select all that apply)/Other"
label variable q7_3_other_text "Other34"
label variable q7_4_inkind_value_bdt "7.4 Estimated monetary value of in-kind contribution (BDT)"
label variable q8_1_other_training_running "8.1 Does your institution currently run any other training programs besides ASS"
label variable q8_2_grp "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_tvet "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_apprenticeship "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_bizmgmt "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_entrepreneurship "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_finance "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_digital "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_softskills "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_other_flag "8.2 If yes or being planned, please provide details of each program: Type of Tr"
label variable q8_2_other_text "Other35"
label variable q8_3_1a_type "8.3.1a Type of Training"
label variable q8_3_1b_funding "8.3.1b Source of Funding"
label variable q8_3_1c_n "8.3.1c No. of Trainees"
label variable q8_3_2a_type "8.3.2a Type of Training"
label variable q8_3_2b_funding "8.3.2b Source of Funding"
label variable q8_3_2c_n "8.3.2c No. of Trainees"
label variable q8_3_3a_type "8.3.3a Type of Training"
label variable q8_3_3b_funding "8.3.3b Source of Funding"
label variable q8_3_3c_n "8.3.3c No. of Trainees"
label variable q8_3_4a_type "8.3.4a Type of Training"
label variable q8_3_4b_funding "8.3.4b Source of Funding"
label variable q8_3_4c_n "8.3.4c No. of Trainees"
label variable q8_3_5a_type "8.3.5a Type of Training"
label variable q8_3_5b_funding "8.3.5b Source of Funding"
label variable q8_3_5c_n "8.3.5c No. of Trainees"
label variable q8_3_6a_type "8.3.6a Type of Training"
label variable q8_3_6b_funding "8.3.6b Source of Funding"
label variable q8_3_6c_n "8.3.6c No. of Trainees"
label variable q8_3_7a_type "8.3.7a Type of Training"
label variable q8_3_7b_funding "8.3.7b Source of Funding"
label variable q8_3_7c_n "8.3.7c No. of Trainees"
label variable q8_3_8a_type "8.3.88a Type of Training"
label variable q8_3_8b_funding "8.3.88b Source of Funding"
label variable q8_3_8c_n "8.3.88c No. of Trainees"
label variable q8_4_grp "8.4 Sources of Funding for Other Training Programs"
label variable q8_4_dev_partner "8.4 Sources of Funding for Other Training Programs/Other Development Partner(s)"
label variable q8_4_own_revenue "8.4 Sources of Funding for Other Training Programs/Own Revenue / Internal Budget"
label variable q8_4_fee_based "8.4 Sources of Funding for Other Training Programs/Fee-Based (paid by trainees/"
label variable q8_4_govt_grant "8.4 Sources of Funding for Other Training Programs/Government Grant / Subsidy"
label variable q8_4_csr "8.4 Sources of Funding for Other Training Programs/Corporate Social Responsibil"
label variable q8_4_other_flag "8.4 Sources of Funding for Other Training Programs/Other (specify)"
label variable q8_4_dev_partner_text "Other Development Partner(s)"
label variable q8_4_other_text "Other (specify)"
label variable q9_1_grp "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_trainee_engagement "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_attendance_retention "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_lack_trainers "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_coordination "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_cblm_quality "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_equipment_cost "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_compliance "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_employment_tracking "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_payment_delay "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_gender_inclusion "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_ohs "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_other_flag "9.1 What are the main challenges faced specifically in delivering Enterprise-Ba"
label variable q9_1_other_text "Other (specify)36"
label variable q9_2_other_challenge_desc "9.2 Any other challenge not listed above? Please describe:"
label variable q10_1_continue_ebt "10.1 Would your institution be willing to continue EBT after ASSET project fund"
label variable q10_2_grp "10.2 What factors would be most important to sustain EBT?"
label variable q10_2_industry_partnership "10.2 What factors would be most important to sustain EBT?/Continued industry pa"
label variable q10_2_industry_costshare "10.2 What factors would be most important to sustain EBT?/Increased industry co"
label variable q10_2_govt_incentive "10.2 What factors would be most important to sustain EBT?/Government incentives"
label variable q10_2_cblm_capacity "10.2 What factors would be most important to sustain EBT?/Improved CBLM and ins"
label variable q10_2_flex_schedule "10.2 What factors would be most important to sustain EBT?/Flexible training sch"
label variable q10_2_curriculum_update "10.2 What factors would be most important to sustain EBT?/Market-linked curricu"
label variable q10_2_shared_facility "10.2 What factors would be most important to sustain EBT?/Shared training facil"
label variable q10_2_placement_support "10.2 What factors would be most important to sustain EBT?/Improved trainee plac"
label variable q10_2_other_flag "10.2 What factors would be most important to sustain EBT?/Other"
label variable q10_2_other_text "Other (specify)37"
label variable q10_3_sustain_support_needed "10.3 What specific actions or support would your institution need to independen"
label variable q11_1_willing_increase_contrib "11.1 If EBT were to be continued beyond ASSET, would your institution be willin"
label variable q11_2_contrib_share_pref "11.2 If yes, what share of the total EBT cost would your institution be willing"
label variable q11_2_other_text "Other (specify)38"
label variable q11_3_contrib_form "11.3 Would this increased contribution be:"
label variable q11_4_grp "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_outcomes_demo "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_govt_incentive "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_admin_burden_reduce "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_flex_modality "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_certification_access "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_other_flag "11.4 What conditions or incentives would encourage a higher contribution from y"
label variable q11_4_other_text "Other (specify)39"
label variable q11_5_cofinance_comments "11.5 Any additional comments on your institution's capacity or willingness to c"
label variable gps_grp "Please collect your GPS location"
label variable gps_lat "_Please collect your GPS location_latitude"
label variable gps_lon "_Please collect your GPS location_longitude"
label variable gps_alt "_Please collect your GPS location_altitude"
label variable gps_precision "_Please collect your GPS location_precision"





*===============================================================================
* A. GENERIC YES/NO RECODE (Institutional Response vs Physically Observed pairs
*    from Sections 2-6, + a few stand-alone Yes/No items)
*===============================================================================
label define yn01 1 "Yes" 0 "No", replace

local yesno_vars ///
    q2_1a_resp q2_1b_obs q2_2a_resp q2_2b_obs q2_3a_resp q2_3b_obs q2_4a_resp q2_4b_obs ///
    q3_1a_resp q3_1b_obs q3_2a_resp q3_2b_obs q3_3a_resp q3_3b_obs ///
    q3_5a_resp q3_5b_obs q3_6a_resp q3_6b_obs q3_7a_resp q3_7b_obs q3_8a_resp q3_8b_obs ///
    q4_1a_resp q4_1b_obs q4_2a_resp q4_2b_obs q4_3a_resp q4_3b_obs q4_4a_resp q4_4b_obs ///
    q4_5a_resp q4_5b_obs ///
    q5_1a_resp q5_1b_obs q5_2a_resp q5_2b_obs q5_3a_resp q5_3b_obs ///
    q6_5a_resp q6_5b_obs q6_6a_resp q6_6b_obs q6_7a_resp q6_7b_obs ///
    q1_2_stp_registered q1_18_skills_gap q6_1_fee_received

foreach v of local yesno_vars {
    capture confirm string variable `v'
    if _rc == 0 {
        gen byte `v'_n = .
        replace `v'_n = 1 if trim(lower(`v')) == "yes"
        replace `v'_n = 0 if trim(lower(`v')) == "no"
        label values `v'_n yn01
        label variable `v'_n "`v' (numeric)"
    }
    else di as error "`v' is not a string variable - check before recoding"
}

*===============================================================================
* B. AGREEMENT FLAG for every Institutional Response (a) / Enumerator
*    Observation (b) pair -- useful for a data-quality / triangulation
*    cross-tab: 1 = institution and enumerator agree.
*===============================================================================
local pairs q2_1 q2_2 q2_3 q2_4 q3_1 q3_2 q3_3 q3_5 q3_6 q3_7 q3_8 ///
    q4_1 q4_2 q4_3 q4_4 q4_5 q5_1 q5_2 q5_3 q6_5 q6_6 q6_7

foreach base of local pairs {
    capture confirm numeric variable `base'a_resp_n
    capture confirm numeric variable `base'b_obs_n
    if !_rc {
        gen byte `base'_agree = (`base'a_resp_n == `base'b_obs_n) if !missing(`base'a_resp_n, `base'b_obs_n)
        label define agree01 1 "Agree" 0 "Disagree", replace
        label values `base'_agree agree01
        label variable `base'_agree "`base': institution vs enumerator agree"
    }
}

*===============================================================================
* C. MULTI-CATEGORY STRING VARIABLES -> encode into labelled numeric versions
*===============================================================================
encode division,                    gen(division_n)
encode q1_7_sector,                 gen(q1_7_sector_n)
encode q1_10_know_source,           gen(q1_10_know_source_n)
encode q7_1_industry_contrib_type,  gen(q7_1_industry_contrib_type_n)
encode q11_3_contrib_form,          gen(q11_3_contrib_form_n)
encode q11_2_contrib_share_pref,    gen(q11_2_contrib_share_pref_n)

* 3-category items (Yes/No/Not sure or similar) 

encode q8_1_other_training_running,     gen(q8_1_other_training_running_n)
encode q10_1_continue_ebt,              gen(q10_1_continue_ebt_n)
encode q11_1_willing_increase_contrib,  gen(q11_1_willing_increase_contrib_n)

* CBLM origin (NSDA / BTEB / Other) -- both response and observed versions
encode q3_4a_resp, gen(q3_4a_resp_n)
encode q3_4b_obs,  gen(q3_4b_obs_n)

*===============================================================================
* D. NUMERIC-AS-STRING FIELDS -> destring into a companion numeric variable
*    (kept separate rather than replacing, per your instruction). "force"
*    will set any non-numeric text (e.g. "N/A", blank) to missing -- inspect
*    those cases first with:  list q3_9a_resp if real(q3_9a_resp)==.
*===============================================================================
foreach v of varlist q3_9a_resp q3_9b_obs {
    capture confirm string variable `v'
    if _rc == 0 {
        destring `v', gen(`v'_n) force
        label variable `v'_n "`v' (numeric)"
    }
}

*===============================================================================
* E. MULTI-SELECT CHECKBOX ITEMS
*    These (q1_13_*, q7_3_*, q6_4_*, q9_1_*, q10_2_*, q11_4_*, q8_2_*, q8_4_*)
*    were already exported by KoBo as separate 0/1 byte variables -- ONE
*    binary variable per option. No recoding is needed; they are already
*    analysis-ready. 
*===============================================================================

*===============================================================================
* 0. CHECK that count variables are numeric (if any prints an error message,
*    destring it first, e.g. destring varname, gen(varname_num) force)
*===============================================================================
local numchk q1_8_avg_workforce_size q1_9_staff_involved_n q1_6_year_established ///
    q1_3_stp_reg_year q2_5_1a_male q2_5_1b_female q2_5_1c_total q2_5_2c_total ///
    q2_5_3a_ethnic_minority q2_5_3b_pwd q2_5_4c_total q2_5_4b_female ///
    q7_2_cash_amt q7_4_inkind_value_bdt q6_3_total ///
    q1_15_1a_assessed q1_15_1b_competent q1_15_1c_eligible q1_15_1d_employed ///
    q1_15_2a_assessed q1_15_2b_competent q1_15_2c_eligible q1_15_2d_employed ///
    q1_15_3a_assessed q1_15_3b_competent q1_15_3c_eligible q1_15_3d_employed ///
    q1_15_4a_assessed q1_15_4b_competent q1_15_4c_eligible q1_15_4d_employed
foreach v of local numchk {
    capture confirm numeric variable `v'
    if _rc di as error "`v' is NOT numeric - destring before continuing"
}

*===============================================================================
* 1. DERIVED OUTCOME INDICATORS (used here and in 04_crosstabs.do)
*===============================================================================
* --- 1a. Competency and employment rates by programme (Q1.15) ----------------
local i = 1
foreach p in tstp adtp sutp ihtp {
    gen double comp_rate_`p' = q1_15_`i'b_competent / q1_15_`i'a_assessed ///
        if q1_15_`i'a_assessed > 0 & !missing(q1_15_`i'a_assessed, q1_15_`i'b_competent)
    gen double emp_rate_`p'  = q1_15_`i'd_employed / q1_15_`i'c_eligible ///
        if q1_15_`i'c_eligible > 0 & !missing(q1_15_`i'c_eligible, q1_15_`i'd_employed)
    label variable comp_rate_`p' "Competency rate, `=upper("`p'")' (competent / assessed)"
    label variable emp_rate_`p'  "Employment rate, `=upper("`p'")' (employed / eligible)"
    local ++i
}

* --- 1b. Pooled across programmes (per institution) ---------------------------
egen tot_assessed = rowtotal(q1_15_1a_assessed q1_15_2a_assessed q1_15_3a_assessed q1_15_4a_assessed), missing
egen tot_competent = rowtotal(q1_15_1b_competent q1_15_2b_competent q1_15_3b_competent q1_15_4b_competent), missing
egen tot_eligible = rowtotal(q1_15_1c_eligible q1_15_2c_eligible q1_15_3c_eligible q1_15_4c_eligible), missing
egen tot_employed = rowtotal(q1_15_1d_employed q1_15_2d_employed q1_15_3d_employed q1_15_4d_employed), missing
gen double comp_rate_all = tot_competent / tot_assessed if tot_assessed > 0
gen double emp_rate_all  = tot_employed  / tot_eligible if tot_eligible > 0
label variable tot_assessed  "Total trainees assessed (all programmes)"
label variable tot_competent "Total trainees competent (all programmes)"
label variable tot_eligible  "Total trainees eligible for employment (all programmes)"
label variable tot_employed  "Total trainees employed (all programmes)"
label variable comp_rate_all "Competency rate, all programmes (competent / assessed)"
label variable emp_rate_all  "Employment rate, all programmes (employed / eligible)"

* --- 1c. Enrolment / attendance indicators (Q2.5, day of visit) ---------------
gen double female_share_enrol = q2_5_1b_female / q2_5_1c_total if q2_5_1c_total > 0
gen double neet_share_enrol   = q2_5_2c_total  / q2_5_1c_total if q2_5_1c_total > 0
gen double em_share_enrol     = q2_5_3a_ethnic_minority / q2_5_1c_total if q2_5_1c_total > 0
gen double pwd_share_enrol    = q2_5_3b_pwd / q2_5_1c_total if q2_5_1c_total > 0
gen double attend_rate        = q2_5_4c_total / q2_5_1c_total if q2_5_1c_total > 0
gen double female_share_present = q2_5_4b_female / q2_5_4c_total if q2_5_4c_total > 0
label variable female_share_enrol   "Female share of enrolled trainees"
label variable neet_share_enrol     "NEET share of enrolled trainees"
label variable em_share_enrol       "Ethnic-minority share of enrolled trainees"
label variable pwd_share_enrol      "PWD share of enrolled trainees"
label variable attend_rate          "Attendance rate on visit day (present / enrolled)"
label variable female_share_present "Female share of trainees present on visit day"

* --- 1d. Sanity flags: rates above 100% point to data-entry problems ---------
foreach v in comp_rate_all emp_rate_all attend_rate {
    quietly count if `v' > 1 & !missing(`v')
    di as result "`v' > 100% in " r(N) " institution(s) - check in Excel"
}

*===============================================================================
* 2. CATEGORICAL PROFILE VARIABLES -> frequencies and percentages
*===============================================================================
local first = 1
local cat_vars division_n q1_7_sector_n q1_2_stp_registered_n q1_10_know_source_n ///
    q1_18_skills_gap_n q6_1_fee_received_n q7_1_industry_contrib_type_n ///
    q8_1_other_training_running_n q10_1_continue_ebt_n ///
    q11_1_willing_increase_contrib_n q11_2_contrib_share_pref_n q11_3_contrib_form_n ///
    q3_4a_resp_n q3_4b_obs_n

foreach v of local cat_vars {
    local lab : variable label `v'
    if `first' == 1 {
        asdoc tab `v', missing save(institution.doc) title(Frequency table: `v' - `lab')
        local first = 0
    }
    else asdoc tab `v', missing save(institution.doc) append title(Frequency table: `v' - `lab')
}

*===============================================================================
* 3. CONTINUOUS / COUNT VARIABLES -> N, mean, SD, median, min, max
*===============================================================================
local cont_vars q1_8_avg_workforce_size q1_9_staff_involved_n q1_6_year_established ///
    q1_3_stp_reg_year q7_2_cash_amt q7_4_inkind_value_bdt q6_3_total ///
    q2_5_1a_male q2_5_1b_female q2_5_1c_total q2_5_4c_total ///
    female_share_enrol neet_share_enrol em_share_enrol pwd_share_enrol ///
    attend_rate female_share_present ///
    tot_assessed tot_competent tot_eligible tot_employed comp_rate_all emp_rate_all

asdoc tabstat `cont_vars', stat(N mean sd p50 min max) columns(statistics) ///
    save(institution.doc) append dec(2) title(Summary statistics: continuous and derived indicators)

*===============================================================================
* 4. PROGRAMME-LEVEL OUTCOMES (TSTP / ADTP / SUTP / IHTP)
*===============================================================================
* 4a. Batches ongoing / completed / assessed (Q1.14)
asdoc tabstat q1_14_1a_ongoing q1_14_1b_completed q1_14_1c_assessed ///
    q1_14_2a_ongoing q1_14_2b_completed q1_14_2c_assessed ///
    q1_14_3a_ongoing q1_14_3b_completed q1_14_3c_assessed ///
    q1_14_4a_ongoing q1_14_4b_completed q1_14_4c_assessed, ///
    stat(N sum mean sd p50) columns(statistics) save(institution.doc) append dec(2) ///
    title(Batches ongoing, completed and assessed by programme (Q1.14))

* 4b. Trainees assessed / competent / eligible / employed (Q1.15)
asdoc tabstat q1_15_1a_assessed q1_15_1b_competent q1_15_1c_eligible q1_15_1d_employed ///
    q1_15_2a_assessed q1_15_2b_competent q1_15_2c_eligible q1_15_2d_employed ///
    q1_15_3a_assessed q1_15_3b_competent q1_15_3c_eligible q1_15_3d_employed ///
    q1_15_4a_assessed q1_15_4b_competent q1_15_4c_eligible q1_15_4d_employed, ///
    stat(N sum mean sd p50) columns(statistics) save(institution.doc) append dec(2) ///
    title(Trainees assessed, competent, eligible and employed by programme (Q1.15))

* 4c. Competency and employment rates by programme
asdoc tabstat comp_rate_tstp comp_rate_adtp comp_rate_sutp comp_rate_ihtp ///
    emp_rate_tstp emp_rate_adtp emp_rate_sutp emp_rate_ihtp, ///
    stat(N mean sd p50) columns(statistics) save(institution.doc) append dec(3) ///
    title(Competency and employment rates by programme (institution-level))

*===============================================================================
* 5. MULTI-SELECT (CHECKBOX) ITEMS -> % of institutions selecting each option
*    (each variable is 0/1, so the MEAN is the proportion; x100 = percent)
*===============================================================================
asdoc tabstat q1_13_tstp q1_13_adtp q1_13_sutp q1_13_ihtp, stat(N sum mean) ///
    columns(statistics) save(institution.doc) append dec(3) title(EBT programme types implemented (Q1.13))

asdoc tabstat q6_4_doc_delay q6_4_emp_track q6_4_dropout q6_4_admin_delay ///
    q6_4_payment_issue q6_4_not_eligible q6_4_not_informed q6_4_other_flag, ///
    stat(N sum mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(Reasons full payment was not received (Q6.4))

asdoc tabstat q7_3_training_space q7_3_factory_utilities q7_3_machinery_use ///
    q7_3_trainer_time q7_3_coord_salaries q7_3_raw_materials q7_3_admin_support q7_3_other_flag, ///
    stat(N sum mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(In-kind contribution types (Q7.3))

asdoc tabstat q8_2_tvet q8_2_apprenticeship q8_2_bizmgmt q8_2_entrepreneurship ///
    q8_2_finance q8_2_digital q8_2_softskills q8_2_other_flag, ///
    stat(N sum mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(Other training programmes run (Q8.2))

asdoc tabstat q8_4_dev_partner q8_4_own_revenue q8_4_fee_based q8_4_govt_grant ///
    q8_4_csr q8_4_other_flag, stat(N sum mean) columns(statistics) ///
    save(institution.doc) append dec(3) title(Funding sources for other training (Q8.4))

asdoc tabstat q9_1_trainee_engagement q9_1_attendance_retention q9_1_lack_trainers ///
    q9_1_coordination q9_1_cblm_quality q9_1_equipment_cost q9_1_compliance ///
    q9_1_employment_tracking q9_1_payment_delay q9_1_gender_inclusion q9_1_ohs ///
    q9_1_other_flag, stat(N sum mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(Main challenges in EBT delivery (Q9.1))

asdoc tabstat q10_2_industry_partnership q10_2_industry_costshare q10_2_govt_incentive ///
    q10_2_cblm_capacity q10_2_flex_schedule q10_2_curriculum_update ///
    q10_2_shared_facility q10_2_placement_support q10_2_other_flag, ///
    stat(N sum mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(Factors important to sustain EBT (Q10.2))

asdoc tabstat q11_4_outcomes_demo q11_4_govt_incentive q11_4_admin_burden_reduce ///
    q11_4_flex_modality q11_4_certification_access q11_4_other_flag, ///
    stat(N sum mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(Incentives for higher industry contribution (Q11.4))

*===============================================================================
* 6. COMPLIANCE ITEMS (Sections 2-6): share answering "Yes", by source
*    (institutional response vs enumerator observation), mean of 0/1 variables
*===============================================================================
asdoc tabstat q2_1a_resp_n q2_1b_obs_n q2_2a_resp_n q2_2b_obs_n q2_3a_resp_n q2_3b_obs_n ///
    q2_4a_resp_n q2_4b_obs_n q3_1a_resp_n q3_1b_obs_n q3_2a_resp_n q3_2b_obs_n ///
    q3_3a_resp_n q3_3b_obs_n q3_5a_resp_n q3_5b_obs_n q3_6a_resp_n q3_6b_obs_n ///
    q3_7a_resp_n q3_7b_obs_n q3_8a_resp_n q3_8b_obs_n, stat(N mean) ///
    columns(statistics) save(institution.doc) append dec(3) ///
    title(Share answering Yes: Sections 2 and 3 (response vs observed))

asdoc tabstat q4_1a_resp_n q4_1b_obs_n q4_2a_resp_n q4_2b_obs_n q4_3a_resp_n q4_3b_obs_n ///
    q4_4a_resp_n q4_4b_obs_n q4_5a_resp_n q4_5b_obs_n q5_1a_resp_n q5_1b_obs_n ///
    q5_2a_resp_n q5_2b_obs_n q5_3a_resp_n q5_3b_obs_n q6_5a_resp_n q6_5b_obs_n ///
    q6_6a_resp_n q6_6b_obs_n q6_7a_resp_n q6_7b_obs_n, stat(N mean) ///
    columns(statistics) save(institution.doc) append dec(3) ///
    title(Share answering Yes: Sections 4 to 6 (response vs observed))


	
	
	
	
	
*******************************************************************************
* 04_crosstabs.do   (Stata 17)
* Cross-tabulations planned from the questionnaire logic. Every table is
* APPENDED to the same Word file: institution.doc  (created by 03_descriptives.do)
*
* Two-way tables are produced twice: once with ROW % and once with COLUMN %.
* Cell counts are small (about 104 institutions), so read percentages with care.
* No chi-square/Fisher tests are included yet (add ", chi2" or "exact" later).
*


*-------------------------------------------------------------------------------
* Helper: one program that writes a two-way table with row % and again with col %
*-------------------------------------------------------------------------------
capture program drop xt2
program define xt2
    args rowvar colvar ttl
    asdoc tab `rowvar' `colvar', row save(institution.doc) append title(`ttl' - ROW %)
    asdoc tab `rowvar' `colvar', col save(institution.doc) append title(`ttl' - COLUMN %)
end

*===============================================================================
* BLOCK 1. Institutional profile and programme attitudes (categorical x categorical)
*===============================================================================
* 1. Sector x willingness to continue EBT after ASSET funding ends (Q1.7 x Q10.1)
xt2 q1_7_sector_n q10_1_continue_ebt_n "Sector by willingness to continue EBT"

* 2. Division x willingness to continue EBT
xt2 division_n q10_1_continue_ebt_n "Division by willingness to continue EBT"

* 3. Sector x how the 15% industry contribution was fulfilled (Q1.7 x Q7.1)
xt2 q1_7_sector_n q7_1_industry_contrib_type_n "Sector by type of industry contribution"

* 4. Division x per-trainee fee received from ASSET (Q6.1)
xt2 division_n q6_1_fee_received_n "Division by per-trainee fee received"

* 5. STP registration x instructor certification, response and observed (Q1.2 x Q3.2)
xt2 q1_2_stp_registered_n q3_2a_resp_n "STP registration by instructors certified (institution response)"
xt2 q1_2_stp_registered_n q3_2b_obs_n  "STP registration by instructors certified (enumerator observed)"

* 6. Willing to continue x willing to increase contribution (Q10.1 x Q11.1)
xt2 q10_1_continue_ebt_n q11_1_willing_increase_contrib_n "Continue EBT by willingness to increase contribution"

* 7. Sector x reported technical skills gap (Q1.7 x Q1.18)
xt2 q1_7_sector_n q1_18_skills_gap_n "Sector by reported technical skills gap"

* 8. Industry contribution type x willingness to increase contribution (Q7.1 x Q11.1)
xt2 q7_1_industry_contrib_type_n q11_1_willing_increase_contrib_n "Current contribution type by willingness to increase"

* 9. How institution heard about EBT x willingness to continue (Q1.10 x Q10.1)
xt2 q1_10_know_source_n q10_1_continue_ebt_n "Source of EBT information by willingness to continue"

* 10. CBLM origin: institution response x enumerator observed (Q3.4a x Q3.4b)
asdoc tab q3_4a_resp_n q3_4b_obs_n, row save(institution.doc) append title(CBLM origin: response vs observed - ROW %)
asdoc tab q3_4a_resp_n q3_4b_obs_n, col save(institution.doc) append title(CBLM origin: response vs observed - COLUMN %)

*===============================================================================
* BLOCK 2. Categorical group x continuous / derived indicators (means)
*===============================================================================
* Workforce size, fee level and in-kind value
foreach g in q1_7_sector_n division_n {
    local glab : variable label `g'
    asdoc tabstat q1_8_avg_workforce_size q1_9_staff_involved_n q6_3_total ///
        q7_2_cash_amt q7_4_inkind_value_bdt, by(`g') stat(N mean sd p50) ///
        save(institution.doc) append dec(2) title(Size, fees and contributions by `g')
}

* Outcome rates by sector, division and STP registration
foreach g in q1_7_sector_n division_n q1_2_stp_registered_n {
    asdoc tabstat comp_rate_all emp_rate_all attend_rate female_share_enrol ///
        neet_share_enrol em_share_enrol pwd_share_enrol, by(`g') ///
        stat(N mean sd p50) save(institution.doc) append dec(3) ///
        title(Outcome and inclusion indicators by `g')
}

* Programme-level rates by sector and division
foreach g in q1_7_sector_n division_n {
    asdoc tabstat comp_rate_tstp comp_rate_adtp comp_rate_sutp comp_rate_ihtp ///
        emp_rate_tstp emp_rate_adtp emp_rate_sutp emp_rate_ihtp, by(`g') ///
        stat(N mean) save(institution.doc) append dec(3) ///
        title(Competency and employment rates by programme and `g')
}

* Trainees assessed / employed totals by sector
asdoc tabstat tot_assessed tot_competent tot_eligible tot_employed, ///
    by(q1_7_sector_n) stat(N sum mean) save(`doc') append dec(1) ///
    title(Total trainees assessed, competent, eligible and employed by sector)

*===============================================================================
* BLOCK 3. Data-quality triangulation: institution response vs enumerator observation
*===============================================================================
* 3a. Response x observed cross-tab for every checklist pair (counts and row %)
local pairs q2_1 q2_2 q2_3 q2_4 q3_1 q3_2 q3_3 q3_5 q3_6 q3_7 q3_8 ///
    q4_1 q4_2 q4_3 q4_4 q4_5 q5_1 q5_2 q5_3 q6_5 q6_6 q6_7
foreach b of local pairs {
    capture confirm numeric variable `b'a_resp_n
    if !_rc {
        asdoc tab `b'a_resp_n `b'b_obs_n, row save(institution.doc) append ///
            title(Response vs observed for `b' - ROW %)
    }
}

* 3b. Agreement rate per item (mean of the 0/1 *_agree flags = share who agree)
asdoc tabstat q2_1_agree q2_2_agree q2_3_agree q2_4_agree q3_1_agree q3_2_agree ///
    q3_3_agree q3_5_agree q3_6_agree q3_7_agree q3_8_agree q4_1_agree q4_2_agree ///
    q4_3_agree q4_4_agree q4_5_agree q5_1_agree q5_2_agree q5_3_agree q6_5_agree ///
    q6_6_agree q6_7_agree, stat(N mean) columns(statistics) save(institution.doc) append dec(3) ///
    title(Agreement rate between institution response and enumerator observation)

* 3c. Overall agreement score per institution, then by sector and division
*egen n_items_checked = rownonmiss(q2_1_agree q2_2_agree q2_3_agree q2_4_agree q3_1_agree ///
    q3_2_agree q3_3_agree q3_5_agree q3_6_agree q3_7_agree q3_8_agree q4_1_agree q4_2_agree ///
    q4_3_agree q4_4_agree q4_5_agree q5_1_agree q5_2_agree q5_3_agree q6_5_agree q6_6_agree q6_7_agree)
egen n_items_agree = rowtotal(q2_1_agree q2_2_agree q2_3_agree q2_4_agree q3_1_agree ///
    q3_2_agree q3_3_agree q3_5_agree q3_6_agree q3_7_agree q3_8_agree q4_1_agree q4_2_agree ///
    q4_3_agree q4_4_agree q4_5_agree q5_1_agree q5_2_agree q5_3_agree q6_5_agree q6_6_agree q6_7_agree)
*gen double pct_agree = n_items_agree / n_items_checked if n_items_checked > 0
label variable pct_agree "Share of checklist items where institution and enumerator agree"

*asdoc tabstat pct_agree, stat(N mean sd p50 min max) save(institution.doc) append dec(3) ///
    title(Overall institution-enumerator agreement score)
foreach g in q1_7_sector_n division_n {
    asdoc tabstat pct_agree, by(`g') stat(N mean sd p50) save(institution.doc) append dec(3) ///
        title(Agreement score by `g')
}	
	
	
	
*===============================================================================
* BLOCK 4. Multi-select items by group
* Mean of 0/1 variables = share selecting the option
*===============================================================================

*-------------------------------------------------------------------------------
* Step 1. Create a clean sector variable for asdoc
*-------------------------------------------------------------------------------

capture drop sector_asdoc

clonevar sector_asdoc = q1_7_sector_n

capture label drop sector_asdoc_lbl
label define sector_asdoc_lbl ///
    1 "Agro-Food" ///
    2 "Ceramics" ///
    3 "Construction" ///
    4 "Furniture" ///
    5 "Garment/Textile" ///
    6 "Health" ///
    7 "ICT" ///
    8 "Light Engineering" ///
    9 "Others" ///
    10 "Tourism/Hospitality"

label values sector_asdoc sector_asdoc_lbl


*-------------------------------------------------------------------------------
* Step 2. Define grouping variables
*-------------------------------------------------------------------------------

local strat_vars sector_asdoc division_n


*-------------------------------------------------------------------------------
* Step 3. Challenges in EBT delivery
*-------------------------------------------------------------------------------

asdoc tabstat ///
    q9_1_trainee_engagement ///
    q9_1_attendance_retention ///
    q9_1_lack_trainers ///
    q9_1_coordination ///
    q9_1_cblm_quality ///
    q9_1_equipment_cost ///
    q9_1_compliance ///
    q9_1_employment_tracking ///
    q9_1_payment_delay ///
    q9_1_gender_inclusion ///
    q9_1_ohs, ///
    by(sector_asdoc) ///
    stat(mean) ///
    save(institution.doc) append dec(3) ///
    title(Challenges in EBT delivery by Sector)


asdoc tabstat ///
    q9_1_trainee_engagement ///
    q9_1_attendance_retention ///
    q9_1_lack_trainers ///
    q9_1_coordination ///
    q9_1_cblm_quality ///
    q9_1_equipment_cost ///
    q9_1_compliance ///
    q9_1_employment_tracking ///
    q9_1_payment_delay ///
    q9_1_gender_inclusion ///
    q9_1_ohs, ///
    by(division_n) ///
    stat(mean) ///
    save(institution.doc) append dec(3) ///
    title(Challenges in EBT delivery by Division)


*-------------------------------------------------------------------------------
* Step 4. Sustainability factors
*-------------------------------------------------------------------------------

asdoc tabstat ///
    q10_2_industry_partnership ///
    q10_2_industry_costshare ///
    q10_2_govt_incentive ///
    q10_2_cblm_capacity ///
    q10_2_flex_schedule ///
    q10_2_curriculum_update ///
    q10_2_shared_facility ///
    q10_2_placement_support, ///
    by(sector_asdoc) ///
    stat(mean) ///
    save(institution.doc) append dec(3) ///
    title(Sustainability factors by Sector)


asdoc tabstat ///
    q10_2_industry_partnership ///
    q10_2_industry_costshare ///
    q10_2_govt_incentive ///
    q10_2_cblm_capacity ///
    q10_2_flex_schedule ///
    q10_2_curriculum_update ///
    q10_2_shared_facility ///
    q10_2_placement_support, ///
    by(division_n) ///
    stat(mean) ///
    save(institution.doc) append dec(3) ///
    title(Sustainability factors by Division)


*-------------------------------------------------------------------------------
* Step 5. Incentives for higher contribution
*-------------------------------------------------------------------------------

asdoc tabstat ///
    q11_4_outcomes_demo ///
    q11_4_govt_incentive ///
    q11_4_admin_burden_reduce ///
    q11_4_flex_modality ///
    q11_4_certification_access, ///
    by(sector_asdoc) ///
    stat(mean) ///
    save(institution.doc) append dec(3) ///
    title(Incentives for higher contribution by Sector)


asdoc tabstat ///
    q11_4_outcomes_demo ///
    q11_4_govt_incentive ///
    q11_4_admin_burden_reduce ///
    q11_4_flex_modality ///
    q11_4_certification_access, ///
    by(division_n) ///
    stat(mean) ///
    save(institution.doc) append dec(3) ///
    title(Incentives for higher contribution by Division)	
	
	





*===============================================================================
* A. OVERALL MULTIPLE-RESPONSE TABLES
*===============================================================================


ssc install mrtab, replace
	
* Q9.1: Main challenges in EBT delivery
asdoc mrtab q9_1_trainee_engagement q9_1_attendance_retention ///
    q9_1_lack_trainers q9_1_coordination q9_1_cblm_quality ///
    q9_1_equipment_cost q9_1_compliance q9_1_employment_tracking ///
    q9_1_payment_delay q9_1_gender_inclusion q9_1_ohs q9_1_other_flag, ///
    save(institution.doc) append ///
    title(Main challenges in EBT delivery)	
	

	
* Q1.13: Programme types offered
asdoc mrtab q1_13_tstp q1_13_adtp q1_13_sutp q1_13_ihtp, ///
    save(institution.doc) append ///
    title(Programme types offered by the institution)


* Q6.4: Reasons for the reported issue
asdoc mrtab q6_4_doc_delay q6_4_emp_track q6_4_dropout ///
    q6_4_admin_delay q6_4_payment_issue q6_4_not_eligible ///
    q6_4_not_informed q6_4_other_flag, ///
    save(institution.doc) append ///
    title(Reasons for the reported issue)


* Q7.3: Types of in-kind contributions
asdoc mrtab q7_3_training_space q7_3_factory_utilities ///
    q7_3_machinery_use q7_3_trainer_time q7_3_coord_salaries ///
    q7_3_raw_materials q7_3_admin_support q7_3_other_flag, ///
    save(institution.doc) append ///
    title(Types of in-kind contributions provided)


* Q8.2: Other types of training provided
asdoc mrtab q8_2_tvet q8_2_apprenticeship q8_2_bizmgmt ///
    q8_2_entrepreneurship q8_2_finance q8_2_digital ///
    q8_2_softskills q8_2_other_flag, ///
    save(institution.doc) append ///
    title(Other types of training provided)


* Q8.4: Sources of funding
asdoc mrtab q8_4_dev_partner q8_4_own_revenue q8_4_fee_based ///
    q8_4_govt_grant q8_4_csr q8_4_other_flag, ///
    save(institution.doc) append ///
    title(Sources of funding for training activities)


* Q9.1: Main challenges in EBT delivery
asdoc mrtab q9_1_trainee_engagement q9_1_attendance_retention ///
    q9_1_lack_trainers q9_1_coordination q9_1_cblm_quality ///
    q9_1_equipment_cost q9_1_compliance q9_1_employment_tracking ///
    q9_1_payment_delay q9_1_gender_inclusion q9_1_ohs q9_1_other_flag, ///
    save(institution.doc) append ///
    title(Main challenges in EBT delivery)


* Q10.2: Measures to sustain EBT
asdoc mrtab q10_2_industry_partnership q10_2_industry_costshare ///
    q10_2_govt_incentive q10_2_cblm_capacity q10_2_flex_schedule ///
    q10_2_curriculum_update q10_2_shared_facility ///
    q10_2_placement_support q10_2_other_flag, ///
    save(institution.doc) append ///
    title(Measures to sustain and strengthen EBT)


* Q11.4: Incentives/support needed
asdoc mrtab q11_4_outcomes_demo q11_4_govt_incentive ///
    q11_4_admin_burden_reduce q11_4_flex_modality ///
    q11_4_certification_access q11_4_other_flag, ///
    save(institution.doc) append ///
    title(Incentives or support needed to encourage EBT)	
	

*===============================================================================
* END
*===============================================================================

di as result "Done. Multiple-response tables appended to institution.doc"	
	

**--------------------------------------------
*			Saving the data as SPSS file
**--------------------------------------------
ssc install savespss

savespss "I:\15. Data Set\Clean\institution.sav", replace
	
```



