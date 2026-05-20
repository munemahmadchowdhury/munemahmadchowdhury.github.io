---
title: "Does sustainable industrialization reduce climate vulnerability?"
excerpt: "**Methodology**: Panel Unit Root Tests, Cross-Sectional Dependence Test, Test for Slope Homogeneity (Pesaran-Yamagata Test), Kao & Pedroni test for cointegration, Cross-sectional time-series FGLS regression (Base Model), Oster Bound Test for omitted variables bias, Driscoll-Kraay standard errors Regression, Prais–Winsten regression- heteroskedastic panels corrected standard errors, and Method of Moments Quantile Regression (MMQR).<br/><img src='/images/portfolio_image5.png'>"
collection: portfolio
lang: "STATA"
date: 2025-09-21
---




```stata
cls
clear all

/// 1. To install necessary packages

ssc install xtcd2, replace
ssc install xttest3, replace
ssc install xtpmg, replace
ssc install xtcsd, replace
ssc install levinlin, replace
ssc install multipurt, replace
ssc install xtgcause, replace
ssc install xtcd, replace
ssc install pescadf, replace
ssc install xtdcce2, replace
ssc install moremata, replace
ssc install xtscc, replace 
ssc install xthst, replace 
net install xtdcce2, from("https://janditzen.github.io/xtdcce2/") replace
ssc install mmqreg
ssc install qregplot
ssc install psacalc



/// 2. To set working directory and load data
cd "I:\All\Ongoing Article Works\Brice\Data Analysis"
import excel "Dataset_20250826.xlsx", sheet("Sheet8") firstrow clear



/// 3. To encode country and set as panel data
encode country, gen(country_code)
xtset country_code year

**** Order the variables nicely
order country country_code year



/// 4. To Label variables
label variable ind "SDG 9 Industry Index score"
label variable vul "Vulnerability Score"
label variable co2 "Carbon dioxide (CO2) emissions (total) excluding LULUCF (Mt CO2e)"
label variable co2cap "Carbon dioxide (CO2) emissions excluding LULUCF per capita (t CO2e/capita)"
label variable cpi10 "Consumer price index (2010 = 100)"
label variable dcgdp "Domestic credit to private sector (% of GDP)"
label variable dcpsgdp "Domestic credit to private sector by banks (% of GDP)"
label variable elcconcap "Electric power consumption (kWh per capita)"
label variable elcfossil "Electricity production from oil, gas and coal sources (% of total)"
label variable elcrenp "Electricity production from renewable sources, excluding hydroelectric (% of total)"
label variable elcren "Electricity production from renewable sources, excluding hydroelectric (kWh)"
label variable expgdp "Exports of goods and services (% of GDP)"
label variable exp "Exports of goods and services (current US$)"
label variable fdigdp "Foreign direct investment, net inflows (% of GDP)"
label variable fdi "Foreign direct investment, net inflows (BoP, current US$)"
label variable fossil "Fossil fuel energy consumption (% of total)"
label variable gdp "GDP (current US$)"
label variable gdpcapp "GDP per capita growth (annual %)"
label variable gdpcap "GDP per capita (current US$)"
label variable ge "Government Effectiveness: Estimate"
label variable gcfgdp "Gross capital formation (% of GDP)"
label variable gcf "Gross capital formation (current US$)"
label variable gdsgdp "Gross domestic savings (% of GDP)"
label variable gds "Gross domestic savings (current US$)"
label variable gfcf "Gross fixed capital formation (current US$)"
label variable gs "Gross savings (current US$)"
label variable hci "Human capital index (HCI) (scale 0-1)"
label variable ictp "ICT goods exports (% of total goods exports)"
label variable ict "ICT service exports (BoP, current US$)"
label variable impgdp "Imports of goods and services (% of GDP)"
label variable imp "Imports of goods and services (current US$)"
label variable intpop "Individuals using the Internet (% of population)"
label variable valuep "Industry (including construction), value added (annual % growth)"
label variable value "Industry (including construction), value added (current US$)"
label variable inf "Inflation, consumer prices (annual %)"
label variable infdef "Inflation, GDP deflator (annual %)"
label variable invict "Investment in ICT with private participation (current US$)"
label variable lint "Lending interest rate (%)"
label variable remgdp "Personal remittances, received (% of GDP)"
label variable rem "Personal remittances, received (current US$)"
label variable popg "Population growth (annual %)"
label variable pop "Population, total"
label variable rint "Real interest rate (%)"
label variable renelc "Renewable electricity output (% of total electricity output)"
label variable renw "Renewable energy consumption (% of total final energy consumption)"
label variable res "Research and development expenditure (% of GDP)"
label variable gre "Total greenhouse gas emissions excluding LULUCF (Mt CO2e)"
label variable grecap "Total greenhouse gas emissions excluding LULUCF per capita (t CO2e/capita)"
label variable to "Trade (% of GDP)"
label variable urb "Urban population"
label variable urbp "Urban population (% of total population)"
label variable popden "Population density (people per sq. km of land area)"
	
label variable fd "Financial Development Index"
label variable fi "Financial Institutions"
label variable fm  "Financial Markets"
label variable gdpg	"GDP growth (annual %)"



/// 5. To check Missing values
misstable sum vul ind gdpg fd ge to popg urbp popden intpop



/// 6. To impute missing value by using linear interpolation

local vars vul ind gdpg fd ge to popg urbp popden intpop

foreach var of local vars {
    bysort country_code (year): ipolate `var' year, gen(`var'_imp) epolate
    replace `var' = `var'_imp if missing(`var')
    drop `var'_imp
}


**** Recheck for confirmation about imputation
sum vul ind gdpg fd ge to popg urbp popden intpop



/// 7. To sub-set the data set for analyzing the data by income group


* Create a new variable for income groups
gen income_group = .


* High-income economies (GNI per capita > $13,935)
replace income_group = 1 if country == "Australia"
replace income_group = 1 if country == "Austria"
replace income_group = 1 if country == "Bahrain"
replace income_group = 1 if country == "Belgium"
replace income_group = 1 if country == "Brunei Darussalam"
replace income_group = 1 if country == "Bulgaria"
replace income_group = 1 if country == "Canada"
replace income_group = 1 if country == "Chile"
replace income_group = 1 if country == "Costa Rica"
replace income_group = 1 if country == "Croatia"
replace income_group = 1 if country == "Cyprus"
replace income_group = 1 if country == "Czechia"
replace income_group = 1 if country == "Denmark"
replace income_group = 1 if country == "Estonia"
replace income_group = 1 if country == "Finland"
replace income_group = 1 if country == "France"
replace income_group = 1 if country == "Germany"
replace income_group = 1 if country == "Greece"
replace income_group = 1 if country == "Hungary"
replace income_group = 1 if country == "Iceland"
replace income_group = 1 if country == "Ireland"
replace income_group = 1 if country == "Israel"
replace income_group = 1 if country == "Italy"
replace income_group = 1 if country == "Japan"
replace income_group = 1 if country == "Kuwait"
replace income_group = 1 if country == "Latvia"
replace income_group = 1 if country == "Lithuania"
replace income_group = 1 if country == "Luxembourg"
replace income_group = 1 if country == "Malta"
replace income_group = 1 if country == "Netherlands"
replace income_group = 1 if country == "New Zealand"
replace income_group = 1 if country == "Norway"
replace income_group = 1 if country == "Oman"
replace income_group = 1 if country == "Panama"
replace income_group = 1 if country == "Poland"
replace income_group = 1 if country == "Portugal"
replace income_group = 1 if country == "Romania"
replace income_group = 1 if country == "Russian Federation"
replace income_group = 1 if country == "Saudi Arabia"
replace income_group = 1 if country == "Singapore"
replace income_group = 1 if country == "Slovak Republic"
replace income_group = 1 if country == "Slovenia"
replace income_group = 1 if country == "Spain"
replace income_group = 1 if country == "Sweden"
replace income_group = 1 if country == "Switzerland"
replace income_group = 1 if country == "United Kingdom"
replace income_group = 1 if country == "United States"
replace income_group = 1 if country == "Uruguay"


* Upper-middle-income economies ($4,496 to $13,935)
replace income_group = 2 if country == "Albania"
replace income_group = 2 if country == "Algeria"
replace income_group = 2 if country == "Argentina"
replace income_group = 2 if country == "Armenia"
replace income_group = 2 if country == "Azerbaijan"
replace income_group = 2 if country == "Belarus"
replace income_group = 2 if country == "Bosnia and Herzegovina"
replace income_group = 2 if country == "Botswana"
replace income_group = 2 if country == "Brazil"
replace income_group = 2 if country == "China"
replace income_group = 2 if country == "Colombia"
replace income_group = 2 if country == "Ecuador"
replace income_group = 2 if country == "El Salvador"
replace income_group = 2 if country == "Gabon"
replace income_group = 2 if country == "Georgia"
replace income_group = 2 if country == "Guatemala"
replace income_group = 2 if country == "Indonesia"
replace income_group = 2 if country == "Iran, Islamic Rep."
replace income_group = 2 if country == "Iraq"
replace income_group = 2 if country == "Jamaica"
replace income_group = 2 if country == "Kazakhstan"
replace income_group = 2 if country == "Malaysia"
replace income_group = 2 if country == "Mauritius"
replace income_group = 2 if country == "Mexico"
replace income_group = 2 if country == "Mongolia"
replace income_group = 2 if country == "Paraguay"
replace income_group = 2 if country == "Peru"
replace income_group = 2 if country == "South Africa"
replace income_group = 2 if country == "Thailand"
replace income_group = 2 if country == "Turkiye"


* Lower-middle-income economies ($1,136 to $4,495)
replace income_group = 3 if country == "Angola"
replace income_group = 3 if country == "Bangladesh"
replace income_group = 3 if country == "Bolivia"
replace income_group = 3 if country == "Cameroon"
replace income_group = 3 if country == "Cote d'Ivoire"
replace income_group = 3 if country == "Egypt, Arab Rep."
replace income_group = 3 if country == "Ghana"
replace income_group = 3 if country == "Haiti"
replace income_group = 3 if country == "Honduras"
replace income_group = 3 if country == "India"
replace income_group = 3 if country == "Jordan"
replace income_group = 3 if country == "Kenya"
replace income_group = 3 if country == "Kyrgyz Republic"
replace income_group = 3 if country == "Lao PDR"
replace income_group = 3 if country == "Lebanon"
replace income_group = 3 if country == "Morocco"
replace income_group = 3 if country == "Namibia"
replace income_group = 3 if country == "Nepal"
replace income_group = 3 if country == "Nicaragua"
replace income_group = 3 if country == "Pakistan"
replace income_group = 3 if country == "Philippines"
replace income_group = 3 if country == "Senegal"
replace income_group = 3 if country == "Sri Lanka"
replace income_group = 3 if country == "Tajikistan"
replace income_group = 3 if country == "Tanzania"
replace income_group = 3 if country == "Tunisia"
replace income_group = 3 if country == "Uzbekistan"
replace income_group = 3 if country == "Viet Nam"
replace income_group = 3 if country == "Zimbabwe"



* To apply value labels
label define income_labels 1 "High income" 2 "Upper middle income" 3 "Lower middle income"
label values income_group income_labels

* To verify the assignment
tab income_group, missing



/// 8. To sub-set the data set for analyzing the data by developed and developing

gen dev_status = .

replace dev_status = 1 if income_group == 1
replace dev_status = 2 if income_group == 2
replace dev_status = 2 if income_group == 3

label define dev_labels 1 "Developed" 2 "Developing"
label values dev_status dev_labels

sum dev_status 


/// 9. Descriptive Statistics
sum vul ind gdpg fd ge to popg

/// 10. Correlations
pwcorr vul ind gdpg fd ge to popg, star(0.05)


/// 11. Unit root Tests

**** Fisher-type ADF test for levels and first differences
local vars vul ind gdpg fd ge to popg 

foreach var of local vars {
    * Test at level
    di "===== Fisher ADF test for `var' (Level) ====="
    xtunitroot fisher `var', dfuller lags(1)  
    
    * Create first difference and test
    gen d_`var' = D.`var'
    di "===== Fisher ADF test for d_`var' (First Difference) ====="
    xtunitroot fisher d_`var', dfuller lags(1)  
}


**** Levin-Lin-Chu test with explicit display
foreach var in vul ind gdpg fd ge to popg {
    di "===== LLC test for `var' (Level) ====="
    xtunitroot llc `var', lags(1) 
    di "t-statistic: " r(tstat) ", p-value: " r(p)
    
    di "===== LLC test for d_`var' (First Difference) ====="
    xtunitroot llc d_`var', lags(1) 
    di "t-statistic: " r(tstat) ", p-value: " r(p)
}

**** CIPS test with explicit display
foreach var vul ind gdpg fd ge to popg {
    di "===== CIPS test for `var' (Level) ====="
    xtcips `var', maxlags(1) bglags(1)
    di "CIPS statistic: " r(cips) ", p-value: " r(p)
    
    di "===== CIPS test for d_`var' (First Difference) ====="
    xtcips d_`var', maxlags(1) bglags(1)
    di "CIPS statistic: " r(cips) ", p-value: " r(p)
}


**** Multipurt test (if installed successfully)
multipurt vul ind gdpg fd ge to popg, lags(1)
multipurt d_vul d_ind d_gdpg d_fd d_ge d_to d_popg, lags(1)



/// 12. Run slope homogeneity test (Delta and Adj. Delta tests)
xthst vul ind gdpg fd ge to popg 

/// 13. Cross Sectional Dependence 
xtreg vul ind gdpg fd ge to popg, fe
estimates store fe_model
xtcd2	
xtreg vul ind gdpg fd ge to popg if dev_status == 1, fe
estimates store fe_model
xtcd2
xtreg vul ind gdpg fd ge to popg if dev_status == 2, fe
estimates store fe_model
xtcd2

xtcd vul ind gdpg fd ge to popg


/// 14. Pnael cointegration tests

**** Kao
xtcointtest kao vul ind gdpg fd ge to popg
xtcointtest kao vul ind gdpg fd ge to popg if dev_status == 1
xtcointtest kao vul ind gdpg fd ge to popg if dev_status == 2
**** Pedroni
xtcointtest pedroni vul ind gdpg fd ge to popg
xtcointtest pedroni vul ind gdpg fd ge to popg if dev_status == 1
xtcointtest pedroni vul ind gdpg fd ge to popg if dev_status == 2




/// 15. MMQR for the 10th, 25th, 50th, 75th, and 90th quantiles

mmqreg  vul ind gdpg fd ge to popg, quantile(0.1 0.25 0.5 0.75 0.9) 

qregplot ind gdpg fd ge to popg, ///
    q(10(10)90) ///
    title("MMQR Coefficient Estimates Across Quantiles (Full Sample)") ///
    graphregion(color(white))



mmqreg  vul ind gdpg fd ge to popg if dev_status == 1, quantile(0.1 0.25 0.5 0.75 0.9) 
	

mmqreg  vul ind gdpg fd ge to popg if dev_status == 2, quantile(0.1 0.25 0.5 0.75 0.9) 
	
	
	


/// 16. Two-Way Fixed Effects (Country and Year FE) with Driscoll-Kraay Standard Errors

xtscc vul ind gdpg fd ge to popg, fe
xtscc vul ind gdpg fi ge to popg if dev_status == 1, fe
xtscc vul ind gdpg fi ge to popg if dev_status == 2, fe


/// 17. Model with PCSE, assuming heteroskedasticity and AR(1) autocorrelation

xtpcse vul ind gdpg fd ge to popg, het corr(ar1)
xtpcse vul ind gdpg fd ge to popg if dev_status == 1, het corr(ar1)
xtpcse vul ind gdpg fd ge to popg if dev_status == 2, het corr(ar1)


/// 18. Cross-sectional time-series FGLS regression
xtgls vul ind gdpg fd ge to popg, igls 
xtgls vul ind gdpg fd ge to popg if dev_status == 1, igls
xtgls vul ind gdpg fd ge to popg if dev_status == 2, igls


/// 19. Dumitrescu-Hurlin Panel Causality Test 

xtgcause vul ind, lags(1)
xtgcause vul gdpg, lags(1)
xtgcause vul fd, lags(1)
xtgcause vul ge, lags(1)
xtgcause vul to, lags(1)
xtgcause vul popg, lags(1)

xtgcause ind vul, lags(1)
xtgcause gdpg vul, lags(1)
xtgcause fd vul, lags(1)
xtgcause ge vul, lags(1)
xtgcause to vul, lags(1)
xtgcause popg vul, lags(1)



/// 20. Robustness to omitted variables bias (Oster bounds test)
xtreg vul ind i.year, fe
est store uncontrolled

xtreg vul ind ge fd gdpg to popg i.year, fe
est store controlled

psacalc delta ind, mcontrol(ge fd gdpg to popg) rmax(1)



/// 21. Check coefficient stability-Random-coefficients regression 
xtrc vul ind gdpg fi ge, betas



```
