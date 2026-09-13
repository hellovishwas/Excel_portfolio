# Indian Job Market Analytics Dashboard

An interactive **Indian Job Market Analytics Dashboard built in Microsoft Excel** using a large Indian job-listings dataset.

The project takes job-market data through a complete analytics workflow — from data preparation and feature engineering to Excel Data Model design, Power Pivot, PivotTables, PivotCharts, slicers, and VBA-based interaction — and presents the results through a single interactive dashboard.

![Indian Job Market Analytics Dashboard](Job_analysis_dashboard.png)

---

## Overview

The purpose of this project was to turn a large Indian job-market dataset into a practical analytical dashboard that can be used to explore:

- Job demand
- Hiring activity by company
- Geographic distribution of opportunities
- Work-mode distribution
- Experience requirements
- Career/job categories
- Salary patterns
- In-demand skills

The final workbook contains separate sheets for the prepared data, PivotTables, Pivot Charts, and the final dashboard.

The dashboard is designed to be interactive: users can apply slicers to filter the analysis, while a custom VBA-powered location search box makes it easier to work with the large number of location values.

---

## Objectives

The project was developed with the following objectives:

- Prepare a large real-world job-market dataset for analysis.
- Clean and standardize job-level fields.
- Create useful salary and experience features.
- Standardize work-mode and location information.
- Process the multi-valued skills field into a structure suitable for analysis.
- Build an Excel Data Model that can handle job-level and skill-level data.
- Avoid incorrect job counts caused by repeated job IDs.
- Build PivotTables and PivotCharts around meaningful business questions.
- Create an interactive dashboard using slicers.
- Add a VBA-based location search interaction.
- Present the analysis in a professional, portfolio-ready Excel dashboard.

---

## Dataset

The project is based on an Indian job listings dataset containing approximately **97,000+ job records**.

During preparation, the dataset was cleaned and transformed for analysis. The job-level data contained approximately **97,682 records** at the cleaning/preparation stage.

The final Excel analysis table is at a **skill-expanded grain**. Because a single job can contain multiple skills, the final table contains repeated `jobId` values and approximately **747,000+ rows**.

This distinction is important:

> **~97K jobs does not mean ~747K jobs.**  
> The larger row count comes from representing multiple skills associated with the same job.

The final analytical table used in the workbook contains the cleaned fields required for the dashboard.

---

## Data Structure

The final workbook contains four main sheets:

| Sheet | Purpose |
|---|---|
| `DATA(Query)` | Main prepared/analysis table |
| `Pivot_Tables` | Supporting PivotTable calculations |
| `Pivot Charts` | Supporting chart objects and chart data |
| `Dashboard` | Final interactive dashboard |

The main data table contains fields such as:

- `title`
- `jobId`
- `companyName`
- `ReviewsCount`
- `AggregateRating`
- `minimumExperience`
- `maximumExperience`
- `minSalaryLPA`
- `maxSalaryLPA`
- `midSalaryLPA`
- `salaryDisclosed`
- `experienceCategory`
- `workMode`
- `location`
- `JobCategory`
- `tagsAndSkillsClean`

---

# Data Preparation & Cleaning

The raw data required several preparation steps before it could be used reliably for dashboard analysis.

## Duplicate Job Records

Duplicate job records were handled so that the same job would not unnecessarily inflate job-level analysis.

This is especially important because the dashboard uses job-level KPIs such as total jobs and company-level job counts.

---

## Salary Preparation

Salary information was processed into separate analytical fields, including:

- Minimum salary
- Maximum salary
- Mid salary
- Salary disclosure status

A derived `midSalaryLPA` field was used for salary analysis.

Salary outlier handling was also considered during preparation so that extreme values would not automatically distort salary-based analysis.

---

## Experience Categorization

Raw experience information was transformed into analytical categories.

The dashboard uses categories including:

- Fresher
- Early Career
- Mid Career
- Experienced
- Senior
- Unknown

This makes it easier to compare salary and job demand across broad experience levels.

---

## Work Mode Standardization

Work-mode values were standardized for dashboard analysis.

The final dashboard compares:

- On-site
- Hybrid
- Remote

---

## Location Cleaning

Location values were cleaned and standardized so that job opportunities could be grouped and compared consistently.

Because the dataset contains a large number of locations, the dashboard also includes a dedicated **Location slicer with a VBA-powered search box**.

---

# Skills Data Processing

The skills/tags field was one of the more important data-preparation challenges.

A single job can contain multiple skills, for example:

```text
Python | SQL | Excel | Power BI
```

If this entire string were treated as one value, it would not be possible to reliably determine which individual skills occur most frequently.

The skills information was therefore transformed so that individual skill values could be analyzed separately.

This produced a skill-expanded analytical table in which one job can correspond to multiple skill records.

The processed skills data contains hundreds of thousands of skill occurrences and a very large number of raw skill values before normalization/analysis.

---

# Data Model & Power Pivot

The project uses Excel's **Data Model and Power Pivot** to structure the analysis.

The most important modeling consideration was the difference between:

### Job Level

One row conceptually represents a job.

Used for:

- Total jobs
- Company hiring activity
- Location distribution
- Work-mode distribution
- Experience categories
- Job categories
- Salary analysis

### Skill Level

One job can have multiple associated skills.

For example:

```text
Job A
 ├── Python
 ├── SQL
 ├── Excel
 └── Power BI
```

Therefore, the skill-expanded table naturally contains repeated job IDs.

The model and analytical logic were designed with this difference in grain in mind instead of treating every skill row as a separate job.

---

## Avoiding Inflated Job Counts

This was one of the key analytical issues in the project.

If a job has four skills, the skill-expanded data can contain four rows for that job.

A simple row count would therefore count:

```text
1 Job × 4 Skills = 4 Rows
```

rather than:

```text
1 Job = 1 Job
```

For this reason, the dashboard's job-level analysis uses appropriate **distinct job counting logic** where required.

This is why the displayed dashboard can show approximately **96,628 jobs** even though the final analysis table contains hundreds of thousands of rows.

The same principle is important when analyzing companies, locations, career categories, and other job-level dimensions.

---

# Dashboard Design

The final dashboard is presented as a **Job Market Intelligence / Analytics Suite**.

It combines:

- KPI cards
- Analytical charts
- PivotCharts
- Interactive slicers
- A VBA-powered location search interaction

The dashboard uses a consistent visual layout so that users can quickly move from high-level KPIs to detailed job-market analysis.

---

# KPI Cards

The dashboard contains four primary KPI cards.

### Total Jobs

Shows the distinct number of jobs represented in the current dashboard context.

In the displayed dashboard state:

**96,628**

### Total Companies

Shows the number of companies represented in the current dashboard context.

In the displayed dashboard state:

**18,381**

### Average Salary

Shows the average salary metric used by the dashboard.

In the displayed dashboard state:

**5.79 LPA**

### Salary Disclosure

Shows the salary disclosure metric represented in the dashboard.

In the displayed dashboard state:

**100.0%**

These KPI values are dynamic and can change when dashboard filters are applied.

---

# Interactive Slicers

The dashboard includes five main slicers:

1. **Work Mode**
2. **Salary Disclosure**
3. **Experience Category**
4. **Location**
5. **Career Category**

These filters allow users to move from a general market view to a more specific segment.

For example, selecting an experience category can change the relevant job and salary analysis across the dashboard.

---

# VBA & Macro Functionality

The workbook is saved as an **`.xlsm` macro-enabled Excel file** because VBA is used for dashboard interaction.

## Location Search Box

A custom ActiveX text box named:

```text
txtLocationSearch
```

is placed with the Location slicer.

The associated VBA event responds when the search text changes and applies matching location values to the Location slicer.

Conceptually:

```text
User enters location text
        ↓
VBA Change Event
        ↓
Search Location Slicer Items
        ↓
Matching locations remain visible/selected
```

This makes the Location filter more practical because a location slicer can contain many values that would otherwise require manual scrolling.

The workbook therefore combines standard Excel slicers with a small amount of VBA to improve dashboard usability.

---

# Key Visualizations

## 1. Work Mode Distribution

### Business Question

**What type of working arrangement dominates the job listings?**

The visualization compares On-site, Hybrid, and Remote opportunities.

In the displayed dashboard state, **On-site jobs form the largest share**, with Hybrid and Remote representing smaller portions.

This provides a quick overview of the working-model distribution in the analyzed dataset.

---

## 2. Geographic Distribution of Job Opportunities

### Business Question

**Which locations have the largest concentration of job opportunities?**

The location analysis ranks locations according to the number of jobs represented in the dashboard.

In the displayed state, **Bengaluru** has the highest job volume among the locations shown, followed by Hyderabad and Pune.

This helps identify major geographic job hubs within the analyzed data.

---

## 3. Top Companies Hiring in the Market

### Business Question

**Which companies have the highest number of job listings in the dataset?**

The visualization ranks companies by job volume.

In the displayed dashboard state, **Accenture** is the leading company in the shown ranking, followed by companies including IDESLABS, Wipro, and others.

This provides a view of relative hiring activity represented in the dataset.

---

## 4. Salary Progression Across Experience Levels

### Business Question

**How does average salary change across experience levels?**

The visualization compares average salary across the dashboard's experience categories.

The displayed values show an upward progression:

| Experience Category | Average Salary |
|---|---:|
| Fresher | ~2.85 LPA |
| Early Career | ~3.19 LPA |
| Mid Career | ~6.54 LPA |
| Experienced | ~12.83 LPA |
| Senior | ~21.99 LPA |

This makes the relationship between experience level and salary easy to interpret.

---

## 5. Job Opportunities Across Career Categories

### Business Question

**Which career categories have the highest job demand?**

The dashboard compares job volume across career categories.

In the displayed state, some of the larger categories include:

- Sales & Business Development
- IT & Software
- Data & AI
- Banking & Insurance
- Finance & Accounting

This provides a high-level view of where job listings are concentrated by career category.

---

## 6. Most In-Demand Skills

### Business Question

**Which individual skills appear most frequently across the job listings?**

This visualization is made possible by the skills transformation described earlier.

The dashboard highlights skills such as:

- Sales
- Python
- Project Management
- Customer Service
- SAP
- Management
- CSS
- Java
- SQL
- Business Development

The purpose of this analysis is to move beyond job titles and identify the skills appearing across the job listings.

---

# Key Insights From the Dashboard

The final dashboard provides several descriptive observations.

### Work Mode

The displayed dashboard shows a strong concentration of **On-site** opportunities compared with Hybrid and Remote roles.

### Geographic Concentration

Major Indian employment hubs account for substantial job volume. Bengaluru is the highest among the locations visible in the displayed dashboard state.

### Career Categories

Sales & Business Development, IT & Software, and Data & AI are among the larger career categories shown in the dashboard.

### Salary and Experience

Average salary increases substantially across the experience categories shown, with Senior roles having the highest displayed average salary.

### Companies

The company ranking highlights organizations with comparatively high job-listing volume in the analyzed dataset, with Accenture appearing at the top of the displayed ranking.

### Skills

The Top Skills analysis demonstrates demand across both technical and business-oriented skills, including Python, SQL, Java, SAP, Sales, Management, and Business Development.

These are **descriptive observations from the analyzed dataset**, not claims about the entire Indian job market.

---

# Project Workflow

The complete workflow can be summarized as:

```text
Raw Job Dataset
        ↓
Data Cleaning & Standardization
        ↓
Duplicate Handling
        ↓
Salary Preparation
        ↓
Experience Categorization
        ↓
Work Mode Standardization
        ↓
Location Cleaning
        ↓
Skills Transformation
        ↓
Job Category Preparation
        ↓
Excel Table
        ↓
Excel Data Model / Power Pivot
        ↓
Analytical Relationships & Counting Logic
        ↓
PivotTables
        ↓
PivotCharts
        ↓
Slicers
        ↓
VBA Location Search
        ↓
Interactive Dashboard
```

The key principle was to build the dashboard **on top of prepared and modeled data**, rather than directly charting raw data.

---

# Challenges & Solutions

## Challenge 1 — Large Dataset

### Problem

The project started with a large Indian job-market dataset containing tens of thousands of job records.

### Solution

The data was cleaned, standardized, and structured before being used for dashboard analysis.

---

## Challenge 2 — Duplicate Job Records

### Problem

Duplicate jobs could inflate job counts and distort company or location analysis.

### Solution

Duplicate handling and distinct job-counting logic were used for job-level analysis.

---

## Challenge 3 — Multi-Valued Skills

### Problem

One job can contain many skills in a single field.

### Solution

The skills information was transformed into individual skill records so that skills could be counted and ranked separately.

---

## Challenge 4 — Different Data Granularity

### Problem

The job-level data and skill-level data do not have the same grain.

One job can generate multiple skill records.

### Solution

The analysis was designed with the difference between job-level and skill-level records in mind.

Job-level metrics use job-level counting logic, while the Top Skills analysis operates on the skill-expanded data.

---

## Challenge 5 — Large Location List

### Problem

A large number of location values makes a standard slicer harder to navigate.

### Solution

A custom ActiveX search box was added to the dashboard and connected to VBA.

The search box filters the Location slicer based on the user's typed text.

---

## Challenge 6 — Salary Data Quality

### Problem

Salary fields require careful handling because values can be missing, undisclosed, inconsistent, or affected by extreme values.

### Solution

Minimum and maximum salary fields were separated, a mid-salary field was derived, salary disclosure was identified, and salary outliers were handled during preparation.

---

# Tools & Technologies

| Tool / Technology | Usage |
|---|---|
| **Microsoft Excel** | Main platform for analysis and dashboard development |
| **Power Query** | Data preparation and transformation |
| **Excel Data Model** | Structured analytical model |
| **Power Pivot** | Data modeling and analytical calculations |
| **PivotTables** | Aggregation and analysis |
| **PivotCharts** | Dashboard visualizations |
| **Slicers** | Interactive dashboard filtering |
| **VBA / Macro** | Location search interaction |
| **ActiveX TextBox** | Search input for the Location slicer |

The final dashboard itself was developed in **Microsoft Excel**. VBA was used as a supporting layer for the Location search functionality.

---

# What I Learned

This project provided practical experience in:

- Cleaning and preparing a large real-world dataset.
- Understanding the importance of data grain.
- Handling duplicate job records.
- Creating derived analytical fields.
- Preparing salary data for analysis.
- Categorizing experience levels.
- Standardizing work modes and locations.
- Transforming a multi-valued skills column.
- Designing job-level versus skill-level analysis.
- Working with Excel Data Model and Power Pivot.
- Understanding how incorrect relationships can lead to inflated counts.
- Building PivotTables and PivotCharts for business questions.
- Designing interactive dashboards using slicers.
- Using VBA to improve Excel dashboard usability.
- Combining analytical modeling with visual storytelling.

---

# Future Improvements

Potential future improvements include:

- Adding more time-based analysis if the posting-date data is expanded into a dedicated trend view.
- Expanding salary analysis across additional dimensions.
- Further normalizing skill names to improve skill grouping.
- Adding deeper company-level analysis.
- Adding more geographic drill-down capabilities.
- Introducing additional Power Pivot measures.
- Adding more dashboard-level interactions and drill-down features.
- Improving automated refresh and deployment workflows.

---

# Conclusion

The **Indian Job Market Analytics Dashboard** is an end-to-end Excel analytics project that combines data preparation, feature engineering, data modeling, analytical logic, visualization, and dashboard interaction.

The project demonstrates that building an effective Excel dashboard is not only about creating charts.

The workflow required decisions around:

**Data Quality → Data Granularity → Feature Engineering → Skills Transformation → Data Modeling → Counting Logic → Pivot Analysis → Interactivity → Dashboard Design**

The final result is an interactive **Job Market Intelligence** dashboard that allows users to explore job demand, companies, locations, work modes, experience levels, career categories, salaries, and in-demand skills from a large Indian job listings dataset.

---

## Project at a Glance

**Project:** Indian Job Market Analytics Dashboard  
**Domain:** Data Analytics / Business Intelligence  
**Primary Tool:** Microsoft Excel  
**Dataset:** Indian Job Listings  
**Scale:** ~97K+ job records  
**Final Skill-Expanded Analysis Table:** ~747K+ rows  
**Dashboard:** Interactive Excel Dashboard  
**Automation / Interaction:** VBA + ActiveX Location Search  
**Core Excel Features:** Power Query, Power Pivot, Data Model, PivotTables, PivotCharts, Slicers

---

## Dashboard Preview

![Job Market Intelligence Dashboard](Job_analysis_dashboard.png)

---

## Repository Contents

A typical repository structure for this project can be:

```text
Indian-Job-Market-Analytics/
│
├── README.md
├── Vishwas_Sharma - Copy.xlsm
└── Job_analysis_dashboard.png
```

---

**Indian Job Market Analytics | Excel Data Analytics Portfolio Project**
