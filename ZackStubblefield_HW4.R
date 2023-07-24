```{r}
#import dataset 
library(readr)
glassdoortest1 <- read_csv("glassdoortest1.csv")
View(glassdoortest1)
```

```{r}
#packages 
pacman::p_load(openxlsx, 
               tidyverse, 
               devtools,
               conflicted,
               lubridate,
               sentimentr)

#Loading libraries
library(openxlsx)
library(tidyverse)
library(devtools)
library(conflicted)
library(lubridate)
library(sentimentr)


devtools::install_github("mguideng/gdscrapeR")

library(gdscrapeR)
```
```{r}
conflict_prefer("filter", "dplyr")
```

```{r}
#renaming data 
data<-glassdoortest1
```

```{r}
#column names
data<-glassdoortest1
colnames(data)
```

```{r}
data<-data %>%
  rename(ID="...1")
```

```{r}
#Looking at Pros first 
comments_pros<- data %>%
    select(c(ID, pros)) %>%
    filter(!is.na(pros)) %>%
    rename('comments' = 'pros')
```

```{r eval=FALSE}
#convert to lowercase 
comments_pros <- comments_pros %>%
    mutate(comments = tolower(comments))

#received "error in inchar(x): invalid multibyte string, element 1
```
```{r}
install.packages("stringi")
library(stringi)
```
```{r}
comments_pros <- comments_pros %>%
    mutate(comments = stri_trans_tolower(comments))

#Using our friend, ChatGPT, I discovered "stringi" package and including "stri_trans" as part of the code to work around the "error in nchar(x)"
```
Removing Line Breaks from positive comments data frame 
```{r}
library(stringi)

comments_pros$comments <- stri_trans_general(comments_pros$comments, "Latin-ASCII")
 #I received an error a waring: unable to translate 'wage is good but that<92>s it not a good work environment for most people. 3 day work week is nice though.' to a wide stringError in gsub("[\r\n]", "", comments_pros$comments) : input string 74 is invalid' and the "stri_trans_general..latin-ascii was a work around. 

comments_pros <- comments_pros %>%
    select(ID, comments) %>%
    na.omit()
```

Benefits-PTO
```{r}
benefits_PTO <- c(
              '^.*vacation.*$', '(?=.*vaca)(?=.*(?:flexible, unlimited, good, great))',
              '\\bpto\\b'
                     )

pto_pattern<-paste(benefits_PTO, collapse="|")

pto_comments <- as.data.frame(comments_pros[grep(pto_pattern, comments_pros$comments, value = FALSE, perl = TRUE),])

TEST <- comments_pros %>%
    mutate(pto = ifelse(comments %in% pto_comments$comments, "Y",
                             "N")) 

```

Benefits-Insurance 
```{r}
benefits_insurance <- c('(?=.*insur)(?=.*(?:medic|dental|life|vision|supplement|disabl))',
                        '\\b(?:insurance\\W+(?:\\w+\\W+){0,1}?premium)\\b',
                        '\\binsurance\\b'
                        )

benefits_insurance_pattern <- paste(benefits_insurance, collapse = "|")

benefits_insurance_comments <- as.data.frame(comments_pros[grep(benefits_insurance_pattern, comments_pros$comments, value = FALSE, perl = TRUE),]) # This takes the pattern you just created and searches over the entire column of "comments" in the Comments_df

TEST <- TEST %>%
    mutate(benefits_insurance = ifelse(comments %in% benefits_insurance_comments$comments, "Y",
                             "N")) #This creates a new object, TEST, from Comments_df and if any of the comments in the "comments" column match (%in%) the comments exactly, they get a "Y". If not they get a "N" in the new "benefits" column

```

Benefits-Healthcare 
```{r}
benefits_healthcare<- c('\\brx\\b', #this will only get the word "rx" and nothing else
              '^.*medic.*$', #this will get medic, medicine, medical, etc.
              '(?=.*bene)(?=.*(?:health))', #This will get benefits, beneficial, benefit, etc. but only if it occurs with health, healthy, healthcare, in the same comment
              '(?=.*coverage)(?=.*(?:medic|deduct|prescrip|insur|drug|health|dependent))', #This will get coverage, overages, etc. as long as some form of medic, deduct, prescription, etc. occur in the same comment
                    '\\b(?:health\\W+(?:\\w+\\W+){0,1}?care)\\b', #this will only get health care or healthcare (e.g. health and care must occur within one word)
                    '\\bhealthcare\\b', #this will only get the word "healthcare". If there is a space between them, it won't pick it up.
              '\\bhealth\\s?care\\b', #this will get the word "healthcare" or "health care" as the \\s? indicates zero or one whitespace character.
                    '\\b(?:medical\\W+(?:\\w+\\W+){0,3}?benefits|benefits\\W+(?:\\w+\\W+){0,3}?medical)\\b', '^.*benefits.*$')

benefits_healthcarepattern <- paste(benefits_healthcare, collapse = "|") #This puts everything from what you put into `benefits` together into a pattern to search for.

benefits_healthcarecomments <- as.data.frame(comments_pros[grep(benefits_healthcarepattern, comments_pros$comments, value = FALSE, perl = TRUE),]) # This takes the pattern you just created and searches over the entire column of "comments" in the Comments_pros

TEST <- TEST %>%
    mutate(benefits_healthcare = ifelse(comments %in% benefits_healthcarecomments$comments, "Y",
                             "N"))
```

Retirement
```{r}
retirement<-c('\\bretirement\\b','\\b401k\\b')

retirement_pattern<-paste(retirement,collapse = "|")

retirement_comments<-as.data.frame(comments_pros[grep(retirement_pattern, comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% 
  mutate(retirement=ifelse(comments %in% retirement_comments$comments, "Y", "N"))
```

Compensation
```{r}
compensation<-c('\\bcompensation\\b','\\bpay\\b', '\\bsalary\\b', '^.*rate.*$', '^.*hourly.*$', '(?=.*starting)(?=.*(?:pay|salary))')

compensation_pattern<-paste(compensation,collapse="|")

compensation_comments<-as.data.frame(comments_pros[grep(compensation_pattern, comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% 
  mutate(compensation=ifelse(comments %in% compensation_comments$comments, "Y", "N"))
```

Schedule 
```{r}
schedule <-c('(?=.*schedule)(?=.*(?:flexible))','(?=.*time)(?=.*(?:flex))', '^.*hours.*$')

schedule_pattern<-paste(schedule, collapse="|")

schedule_comments<-as.data.frame(comments_pros[grep(schedule_pattern, comments_pros$comments, value=FALSE, perl = TRUE),])

TEST<-TEST %>% mutate(schedule=ifelse(comments %in% schedule_comments$comments, "Y", "N"))
```

Development 
```{r}
development<-c('^.*development.*$', '^.*learn.*$','^.*train.*$', '^.*grow.*$', '^.advance.*$', '(?=.*career)(?=.*(?:growth|advancement))')

development_pattern<-paste(development, collapse = "|")

development_comments<-as.data.frame(comments_pros[grep(development_pattern,comments_pros$comments, value=FALSE, perl = TRUE),])

TEST<-TEST %>% mutate(development=ifelse(comments %in%development_comments$comments,"Y", "N"))
                                        
```

Worklife Balance 
```{r}
worklifebalance <- c('(?=.*worklife)(?=.*(?:balance))', '\\bwork\\s?life\\b')

wlb_patterns<-paste(worklifebalance, collapse="|")

wlb_comments<-as.data.frame(comments_pros[grep(wlb_patterns,comments_pros$comments, value=FALSE, perl = TRUE),])

TEST<-TEST %>% mutate(worklifebalance=ifelse(comments%in%wlb_comments$comments, "Y", "N"))
```

Culture 
```{r}
culture<-c('\\bculture\\b', '\\bsafety\\b', '\\benvironment\\b', '\\b(?:environment\\W+(?:\\w+\\W+){0,1}?good)\\b', '\\batmosphere\\b', '^.*autonomy.*$', '^.*freedom.*$') #trying to weed out "bad work environment 

culture_patterns<-paste(culture,collapse = "|")

culture_comments<-as.data.frame(comments_pros[grep(culture_patterns,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST%>% mutate(culture=ifelse(comments%in%culture_comments$comments, "Y","N"))
```
Diversity 
```{r}
diversity<-c('\\b(?:diverse\\W+(?:\\w+\\W+){0,3}?people)\\b')

diversity_pattern<-paste(diversity, collapse="|")

diversity_comments<-as.data.frame(comments_pros[grep(diversity_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% mutate(diversity=ifelse(comments%in%diversity_comments$comments, "Y","N"))
```
Leadership, Supervisors, Management, etc. 
```{r}
leadership<-c('\\bleadership\\b', '\\bsupervisor\\b', '\\bmanagement\\b','^.*vp.*$',  '^.*manage.*$',  '^.*leader.*$' )

leadership_pattern<-paste(leadership, collapse = "|")

leadership_comments<-as.data.frame(comments_pros[grep(leadership_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% mutate(leadership=ifelse(comments%in%leadership_comments$comments, "Y","N"))
```

The People 
```{r}
#teams, people, coworkers, etx. 
people<-c('^.*people.*$', '^.*team.*$','^.*coworker.*$', '^.*co-worker.*$', '(?=.*friendly)(?=.*(?:people))', '\\bsupport\\b')

people_pattern<-paste(people, collapse="|")

people_comments<-as.data.frame(comments_pros[grep(people_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST%>% mutate(people=ifelse(comments%in%people_comments$comments, "Y","N"))

                      
```

The Work 
```{r}
#looking for things like project, work, etc. When I did my search for diversity I discovered that diveristy is commonly discussed in regard to the tyep of work so I will include that here. 
thework<-c('\\bprojects\\b','(?=.*challenge.)(?=.*(?:work))','^.diverse.*$')

thework_pattern<-paste(thework, collapse="|")

thework_comments<-as.data.frame(comments_pros[grep(thework_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% mutate(thework=ifelse(comments%in%thework_comments$comments, "Y","N"))
```

Travel
```{r}
travel<-c('\\btravel\\b')

travel_pattern<-paste(travel, collapse="|")

travel_comments<-as.data.frame(comments_pros[grep(travel_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% mutate(travel=ifelse(comments%in%travel_comments$comments, "Y","N"))
```

Engagement 
```{r}
#what are employees saying about being engaged? 
engagement<-c('^.*engage.*', '^.*interesting.*$')

engagement_pattern<-paste(engagement, collapse="|")

engagement_comments<-as.data.frame(comments_pros[grep(engagement_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% mutate(engagement=ifelse(comments%in%engagement_comments$comments, "Y","N"))
```
Company
```{r}
#what did employees say about the product, technology, global scale, etc.
company<-c('\\bproduct\\b', '^.*tech.*$', '^.*fortune.*$', '^.*global.*$', '\\bcompany\\b', '^.*brand.*$', '^.*industry.*$')

company_pattern<-paste(company, collapse="|")

company_comments<-as.data.frame(comments_pros[grep(company_pattern,comments_pros$comments, value=FALSE, perl=TRUE),])

TEST<-TEST %>% mutate(company=ifelse(comments%in%company_comments$comments, "Y", "N"))
```

Communication
```{r}
#What did employees say about communication? 
communication<-c('\\bcommunication\\b', '(?=.*communication.)(?=.*(?:leader))','^.*email.*$', '^.*talk.*$', '^.*meetings.*$')

communication_pattern<-paste(communication, collapse="|")

communication_comments<-as.data.frame(comments_pros[grep(communication_pattern,comments_pros$comments, value=FALSE, perl = TRUE),])


TEST<-TEST %>% mutate(communication=ifelse(comments%in%communication_comments$comments, "Y", "N"))
```

Location & Work From Home 
```{r}
#What did employees say about office location and work from home and facilities? 
location<-c('\\blocation\\b', '\\boffice\\b', '\\b(?:work\\W+(?:\\w+\\W+){0,2}?home)\\b', '^.facility.*$')

location_pattern<-paste(location, collapse = "|")

location_comments<-as.data.frame(comments_pros[grep(location_pattern,comments_pros$comments, value=FALSE, perl = TRUE),])

TEST<-TEST %>% mutate(location=ifelse(comments%in%location_comments$comments, "Y", "N"))
```
Teams
```{r}
#What did employees say about teams, teamwork, etc? 
teams<-c('\\team\\b','\\b(?:team\\W+(?:\\w+\\W+){0,1}?work)\\b', '^.coworker.*$', '\\teams\\b' )

teams_pattern<-paste(teams, collapse="|")

teams_comments<-as.data.frame(comments_pros[grep(teams_pattern,comments_pros$comments, value=FALSE, perl = TRUE),])

TEST<-TEST %>% mutate(teams=ifelse(comments%in%teams_comments$comments, "Y", "N"))

```
Other
```{r}
#other column for those who don't fit in the other categories 

TEST <- TEST %>%
    mutate(Other = apply(TEST, 1, function(y){ ifelse("Y" %in% y, "N", "Y")}))
```

Pros Excel
```{r}
INTRO <- c("UGA",

         "Data Source: Glassdoor",

         "Data As Of: Q3 2020",

         "Prepared on: 7/023/2023",

         "Prepared by: Zack")

wb <- openxlsx::createWorkbook() 


#Comment Report

addWorksheet(wb, "Pros Comments") 

writeData(wb, "Pros Comments", INTRO) 


#Create style

style1 <- createStyle(fontColour = "#823017", textDecoration = "Bold") #Choose your custom font color (https://www.rgbtohex.net/) and make it bold. Call it style1

 

addStyle(wb, style = style1, rows= 1:5, cols = 1, sheet = "Pros Comments") 

writeData(wb, "Pros Comments", TEST, startRow = 8) 

hs1 <- createStyle(textDecoration = "Bold") 

addStyle(wb, style = hs1, rows = 8, cols = 1:50, sheet = "Pros Comments") 

#Freeze Panes

#Also check here: https://stackoverflow.com/questions/37677326/applying-style-to-all-sheets-of-a-workbook-using-openxlsx-package-in-r

freezePane(wb, "Pros Comments", firstActiveRow = 9) 

#Add filter

addFilter(wb, "Pros Comments", row = 8, cols = 1:50) 

# Load necessary libraries if not already loaded
library(openxlsx)  # Assuming you are using openxlsx for saving the workbook
library(lubridate) # Assuming you are using lubridate for date manipulation

# Define the directory path
directory_path <- "~/Desktop/Advanced Analytics/Homework 4/"

# Check if the directory exists, and if not, create it
if (!file.exists(directory_path)) {
  dir.create(directory_path, recursive = TRUE)
}

# Create the filename with the current month and year
June_2023 <- paste0(directory_path, format(floor_date(Sys.Date()-months(1), "month"), "%B_%Y"), ".xlsx")

# Write the workbook to the specified file
saveWorkbook(wb, June_2023, overwrite = TRUE)


```

Negative Comments 
```{r}
#negative comments 
comments_cons<- data %>%
    select(c(ID, cons)) %>%
    filter(!is.na(cons)) %>%
    rename('comments' = 'cons')
```
Lowercase 
```{r}
comments_cons <- comments_cons %>%
    mutate(comments = stri_trans_tolower(comments))
```
Line Break
```{r}
comments_cons$comments <- stri_trans_general(comments_cons$comments, "Latin-ASCII")
 #I received an error a waring: unable to translate 'wage is good but that<92>s it not a good work environment for most people. 3 day work week is nice though.' to a wide stringError in gsub("[\r\n]", "", comments_pros$comments) : input string 74 is invalid' and the "stri_trans_general..latin-ascii was a work around. 

comments_cons <- comments_cons %>%
    select(ID, comments) %>%
    na.omit()
```

PTO
#will add a 1 to all by objects in the environment to indicate cons
```{r}
benefits_PTO1 <- c(
              '^.*vacation.*$', '(?=.*vaca)(?=.*(?:flexible, unlimited))',
              '\\bpto\\b'
                     )
#removed good and great
pto_pattern1<-paste(benefits_PTO1, collapse="|")

pto_comments1 <- as.data.frame(comments_cons[grep(pto_pattern1, comments_cons$comments, value = FALSE, perl = TRUE),])

TEST1 <- comments_cons %>%
    mutate(pto = ifelse(comments %in% pto_comments1$comments, "Y",
                             "N")) 
```
Benefits Insurance
```{r}
benefits_insurance1 <- c('(?=.*insur)(?=.*(?:medic|dental|life|vision|supplement|disabl))',
                        '\\b(?:insurance\\W+(?:\\w+\\W+){0,1}?premium)\\b',
                        '\\binsurance\\b'
                        )

benefits_insurance_pattern1 <- paste(benefits_insurance1, collapse = "|")

benefits_insurance_comments1 <- as.data.frame(comments_cons[grep(benefits_insurance_pattern1, comments_cons$comments, value = FALSE, perl = TRUE),]) # This takes the pattern you just created and searches over the entire column of "comments" in the Comments_df

TEST1 <- TEST1 %>%
    mutate(benefits_insurance = ifelse(comments %in% benefits_insurance_comments1$comments, "Y",
                             "N")) #This creates a new object, TEST, from Comments_df and if any of the comments in the "comments" column match (%in%) the comments exactly, they get a "Y". If not they get a "N" in the new "benefits" column
```
Benefits Healthcare
```{r}
benefits_healthcare1<- c('\\brx\\b', #this will only get the word "rx" and nothing else
              '^.*medic.*$', #this will get medic, medicine, medical, etc.
              '(?=.*bene)(?=.*(?:health))', #This will get benefits, beneficial, benefit, etc. but only if it occurs with health, healthy, healthcare, in the same comment
              '(?=.*coverage)(?=.*(?:medic|deduct|prescrip|insur|drug|health|dependent))', #This will get coverage, overages, etc. as long as some form of medic, deduct, prescription, etc. occur in the same comment
                    '\\b(?:health\\W+(?:\\w+\\W+){0,1}?care)\\b', #this will only get health care or healthcare (e.g. health and care must occur within one word)
                    '\\bhealthcare\\b', #this will only get the word "healthcare". If there is a space between them, it won't pick it up.
              '\\bhealth\\s?care\\b', #this will get the word "healthcare" or "health care" as the \\s? indicates zero or one whitespace character.
                    '\\b(?:medical\\W+(?:\\w+\\W+){0,3}?benefits|benefits\\W+(?:\\w+\\W+){0,3}?medical)\\b', '^.*benefits.*$')

benefits_healthcarepattern1 <- paste(benefits_healthcare1, collapse = "|") #This puts everything from what you put into `benefits` together into a pattern to search for.

benefits_healthcarecomments1 <- as.data.frame(comments_cons[grep(benefits_healthcarepattern1, comments_cons$comments, value = FALSE, perl = TRUE),]) # This takes the pattern you just created and searches over the entire column of "comments" in the Comments_pros

TEST1 <- TEST1 %>%
    mutate(benefits_healthcare = ifelse(comments %in% benefits_healthcarecomments1$comments, "Y",
                             "N"))
```
Retirement
```{r}
retirement1<-c('\\bretirement\\b','\\b401k\\b')

retirement_pattern1<-paste(retirement1,collapse = "|")

retirement_comments1<-as.data.frame(comments_cons[grep(retirement_pattern1, comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% 
  mutate(retirement=ifelse(comments %in% retirement_comments1$comments, "Y", "N"))
```

Compensation 
```{r}
compensation1<-c('\\bcompensation\\b','\\bpay\\b', '\\bsalary\\b', '^.*rate.*$', '^.*hourly.*$', '(?=.*starting)(?=.*(?:pay|salary))')

compensation_pattern1<-paste(compensation1,collapse="|")

compensation_comments1<-as.data.frame(comments_cons[grep(compensation_pattern1, comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% 
  mutate(compensation=ifelse(comments %in% compensation_comments1$comments, "Y", "N"))
```

Schedule 
```{r}
schedule1 <-c('(?=.*schedule)(?=.*(?:flexible))','(?=.*time)(?=.*(?:flex))', '^.*hours.*$')

schedule_pattern1<-paste(schedule1, collapse="|")

schedule_comments1<-as.data.frame(comments_cons[grep(schedule_pattern1, comments_cons$comments, value=FALSE, perl = TRUE),])

TEST1<-TEST1 %>% mutate(schedule=ifelse(comments %in% schedule_comments1$comments, "Y", "N"))
```
Development 
```{r}
#what did employees say about development, advancement, learning, training?
development1<-c('^.*development.*$', '^.*learn.*$','^.*train.*$', '^.*grow.*$', '^.advance.*$', '(?=.*career)(?=.*(?:growth|advancement))')

development_pattern1<-paste(development1, collapse = "|")

development_comments1<-as.data.frame(comments_cons[grep(development_pattern1,comments_cons$comments, value=FALSE, perl = TRUE),])

TEST1<-TEST1 %>% mutate(development=ifelse(comments %in%development_comments1$comments,"Y", "N"))
```
Work Life Balance 
```{r}
#What are employees saying specifically about work life balance 
worklifebalance1 <- c('(?=.*worklife)(?=.*(?:balance))', '\\bwork\\s?life\\b')

wlb_patterns1<-paste(worklifebalance1, collapse="|")

wlb_comments1<-as.data.frame(comments_cons[grep(wlb_patterns,comments_cons$comments, value=FALSE, perl = TRUE),])

TEST1<-TEST1 %>% mutate(worklifebalance=ifelse(comments%in%wlb_comments1$comments, "Y", "N"))
```

Culture 
```{r}
culture1<-c('\\bculture\\b', '\\bsafety\\b', '\\benvironment\\b', '\\batmosphere\\b', '^.*autonomy.*$', '^.*freedom.*$') 

culture_patterns1<-paste(culture1,collapse = "|")

culture_comments1<-as.data.frame(comments_cons[grep(culture_patterns1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1%>% mutate(culture=ifelse(comments%in%culture_comments1$comments, "Y","N"))
```
Diversity 
```{r}
diversity1<-c('\\b(?:diverse\\W+(?:\\w+\\W+){0,3}?people)\\b')

diversity_pattern1<-paste(diversity1, collapse="|")

diversity_comments1<-as.data.frame(comments_cons[grep(diversity_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% mutate(diversity=ifelse(comments%in%diversity_comments1$comments, "Y","N"))
```
Leadership, Management, Superivsors, etc. 
```{r}
leadership1<-c('\\bleadership\\b', '\\bsupervisor\\b', '\\bmanagement\\b','^.*vp.*$',  '^.*manage.*$',  '^.*leader.*$', '\\bboss\\b')

leadership_pattern1<-paste(leadership1, collapse = "|")

leadership_comments1<-as.data.frame(comments_cons[grep(leadership_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% mutate(leadership=ifelse(comments%in%leadership_comments1$comments, "Y","N"))
```
The People 
```{r}
#teams, people, coworkers, etx. 
people1<-c('^.*people.*$', '^.*team.*$','^.*coworker.*$', '^.*co-worker.*$', '(?=.*friendly)(?=.*(?:people))', '\\bsupport\\b')

people_pattern1<-paste(people1, collapse="|")

people_comments1<-as.data.frame(comments_cons[grep(people_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1%>% mutate(people=ifelse(comments%in%people_comments1$comments, "Y","N"))
```
The Work 
```{r}
thework1<-c('\\bprojects\\b','(?=.*challenge.)(?=.*(?:work))','^.diverse.*$','\\bhours\\b')
#added hours in the cons 
thework_pattern1<-paste(thework1, collapse="|")

thework_comments1<-as.data.frame(comments_cons[grep(thework_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% mutate(thework=ifelse(comments%in%thework_comments1$comments, "Y","N"))
```

Travel 
```{r}
travel1<-c('\\btravel\\b', '\\bcommute\\b') #added commute in cons

travel_pattern1<-paste(travel1, collapse="|")

travel_comments1<-as.data.frame(comments_cons[grep(travel_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1%>% mutate(travel=ifelse(comments%in%travel_comments1$comments, "Y","N"))
```

Engagement 
```{r}
engagement1<-c('^.*engage.*', '^.*interesting.*$')

engagement_pattern1<-paste(engagement1, collapse="|")

engagement_comments1<-as.data.frame(comments_cons[grep(engagement_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% mutate(engagement=ifelse(comments%in%engagement_comments1$comments, "Y","N"))
```

The Company
```{r}
#what did employees say about the product, technology, global scale, etc.
company1<-c('\\bproduct\\b', '^.*tech.*$', '^.*fortune.*$', '^.*global.*$', '\\bcompany\\b', '^.*brand.*$', '^.*industry.*$')

company_pattern1<-paste(company1, collapse="|")

company_comments1<-as.data.frame(comments_cons[grep(company_pattern1,comments_cons$comments, value=FALSE, perl=TRUE),])

TEST1<-TEST1 %>% mutate(company=ifelse(comments%in%company_comments1$comments, "Y", "N"))
```

Communication
```{r}
#What did employees say about communication? 
communication1<-c('\\bcommunication\\b', '(?=.*communication.)(?=.*(?:leader))','^.*email.*$', '^.*talk.*$', '^.*meetings.*$')

communication_pattern1<-paste(communication1, collapse="|")

communication_comments1<-as.data.frame(comments_cons[grep(communication_pattern1,comments_cons$comments, value=FALSE, perl = TRUE),])


TEST1<-TEST1 %>% mutate(communication=ifelse(comments%in%communication_comments1$comments, "Y", "N"))
```
Location 
```{r}
#What did employees say about office location and work from home and facilities? 
location1<-c('\\blocation\\b', '\\boffice\\b', '\\b(?:work\\W+(?:\\w+\\W+){0,2}?home)\\b', '^.facility.*$', '\\bremote\\b')

location_pattern1<-paste(location1, collapse = "|")

location_comments1<-as.data.frame(comments_cons[grep(location_pattern1,comments_cons$comments, value=FALSE, perl = TRUE),])

TEST1<-TEST1 %>% mutate(location=ifelse(comments%in%location_comments1$comments, "Y", "N"))
```
Teams 
```{r}
#What did employees say about teams, teamwork, etc? 
teams1<-c('\\team\\b','\\b(?:team\\W+(?:\\w+\\W+){0,1}?work)\\b', '^.coworker.*$', '\\teams\\b' )

teams_pattern1<-paste(teams1, collapse="|")

teams_comments1<-as.data.frame(comments_cons[grep(teams_pattern1,comments_cons$comments, value=FALSE, perl = TRUE),])

TEST1<-TEST1 %>% mutate(teams=ifelse(comments%in%teams_comments1$comments, "Y", "N"))
```

Other 
```{r}
#other column for those who don't fit in the other categories 

TEST1 <- TEST1 %>%
    mutate(other = apply(TEST1, 1, function(y){ ifelse("Y" %in% y, "N", "Y")}))
```
Cons Excel 
```{r}
INTRO <- c("UGA",

         "Data Source: Glassdoor",

         "Data As Of: Q3 2020",

         "Prepared on: 7/023/2023",

         "Prepared by: Zack")

wb <- openxlsx::createWorkbook() #Create a work book


#Comment Report

addWorksheet(wb, "Cons Comments")

writeData(wb, "Cons Comments", INTRO) 


#Create style

style1 <- createStyle(fontColour = "#545433", textDecoration = "Bold") #Choose your custom font color (https://www.rgbtohex.net/) and make it bold. Call it style1

 

addStyle(wb, style = style1, rows= 1:5, cols = 1, sheet = "Cons Comments") 
writeData(wb, "Cons Comments", TEST, startRow = 8) 

hs1 <- createStyle(textDecoration = "Bold") #

addStyle(wb, style = hs1, rows = 8, cols = 1:50, sheet = "Cons Comments") #

#Freeze Panes

#Also check here: https://stackoverflow.com/questions/37677326/applying-style-to-all-sheets-of-a-workbook-using-openxlsx-package-in-r

freezePane(wb, "Cons Comments", firstActiveRow = 9) 

#Add filter

addFilter(wb, "Cons Comments", row = 8, cols = 1:50) 

# Load necessary libraries if not already loaded
library(openxlsx)  # Assuming you are using openxlsx for saving the workbook
library(lubridate) # Assuming you are using lubridate for date manipulation

# Define the directory path
directory_path <- "~/Desktop/Advanced Analytics/Homework 4/"

# Check if the directory exists, and if not, create it
if (!file.exists(directory_path)) {
  dir.create(directory_path, recursive = TRUE)
}

# Create the filename with the current month and year
June_20231 <- paste0(directory_path, format(floor_date(Sys.Date()-months(1), "month"), "%B_%Y"), ".xlsx")

# Write the workbook to the specified file
saveWorkbook(wb, June_2023, overwrite = TRUE)
```

