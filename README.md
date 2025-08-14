# Call Center Sales and Marketing Performance

## Overview
This project creates an interactive Excel-based dashboard for analyzing call center performance. It processes 200,000 call records from a CSV file, providing insights into call outcomes, agent performance, product discussions, customer demographics, and trends. The dashboard uses PivotTables, charts, slicers, and VBA macros for dynamic visualizations and filtering.

## Features
- **Key Metrics Overview**: Displays total calls, success/failure/abandoned rates, average agent ratings, and call durations.
- **Agent Performance**: Filter by agent to view success rates, ratings, call volumes, and breakdowns by product/outcome.
- **Trends and Visuals**: Monthly trends via slicers, product distribution charts, state-based maps/charts (toggleable via buttons), and reasons for abandoned calls.
- **Customer Insights**: Analysis by age, gender, state, income bracket, and repeat status.
- **Interactive Elements**: VBA macros to toggle map/chart views and slicer visibility; supports filtering by month, agent, product, etc.
- **Data-Driven Recommendations**: Highlights actionable insights, e.g., optimizing staffing to reduce 22.32% abandonment rate or targeting high-value states like New York.

## Technologies Used
- **Microsoft Excel**: PivotTables, charts (bar, pie, line), slicers, conditional formatting, formulas (e.g., AVERAGEIF, SUMIFS).
- **VBA Macros**: For toggling visibility of shapes (e.g., maps, charts, buttons) and slicers.
- **Data Processing**: Imported CSV data into Excel sheets for analysis.

## Data Source
- **File**: Practice Data.csv and Call Center Sales and Marketing.xlsx (Data sheet).
- **Columns**: Call_ID, Date, Agent details, Rating, Product_Discussed, Call_Duration_Minutes, Call_Outcome, Customer_Age, Gender, State, Income_Bracket, Time_of_Day, Follow_Up_Call_Required, Repeat_Customer, Reason_Call_Abandoned.
- **Sample Size**: 200,000 records from 2024.

## Visual Example
![Dashboard]()

## Key Findings and Recommendations

- **Finding**: Loans are the most discussed product, accounting for approximately 30% of successful calls across all agents, with Internet Packages following at 25%. This indicates strong customer interest in financial and connectivity products.
- **Recommendation**: Develop specialized training programs for agents focusing on loan and Internet Package upselling techniques. Prioritize high-income customers in states like New York and California, which contribute significantly to call volume, to increase conversion rates by 10-15% through tailored pitches.


- **Finding**: 22.32% of calls (44,644) are abandoned, with "long wait time" being the primary reason (over 80% of abandoned cases). Afternoon hours show the highest abandonment rates, correlating with peak call volumes.
Recommendation: Optimize staffing schedules by increasing agent availability during afternoon hours (12 PM–4 PM) to reduce wait times. Implement queue management systems and predictive analytics to balance call loads, potentially decreasing abandonment rates by 15-20% and improving customer satisfaction.


- **Finding**: Ava Sandoval handles 20% of total calls (40,000) with a success rate of 19.995% and an average rating of 4.25, while agents like Ray Mason have lower ratings (3.3 average) and success rates (below 15%).
- **Recommendation**: Establish a mentoring program pairing high-performing agents like Ava Sandoval with lower-rated agents to share best practices in call handling and customer engagement. Conduct monthly performance reviews using dashboard data to set improvement targets, aiming to elevate team-wide success rates by 5-10%.


- **Finding**: New York drives 13-15% of total call volume (e.g., 2,682 calls for Monique Lawrence), with high-income customers being the most frequent callers. California and North Carolina also rank high in call volume.
- **Recommendation**: Launch geo-targeted marketing campaigns focusing on high-income and repeat customers in New York, California, and North Carolina. Offer promotions on Loans and Internet Packages to these demographics, leveraging their high engagement to boost revenue by 8-12%.


- **Finding**: Electronics products have the highest failure rate (e.g., 1,303 failures for Ava Sandoval) due to technical issues, while Insurance and Travel Packages show lower failure rates but fewer discussions.
- **Recommendation**: Enhance technical support training for agents handling Electronics inquiries to address common issues, aiming to convert 20% of failures into successes. Increase marketing efforts for Insurance and Travel Packages to high-income customers to diversify product discussions and drive additional sales.


- **Finding**: Call durations for failed calls average 15-19 minutes, significantly longer than successful calls (10-25 minutes), particularly for technical issue-related failures.
- **Recommendation**: Implement scripted protocols for common technical issues, tracked via Excel workflows, to streamline troubleshooting and reduce average call duration for failures by 5-7 minutes. This could increase daily call handling capacity by 10%, improving overall efficiency.


- **Finding**: Repeat customers account for 30% of calls in high-volume states, with a higher likelihood of successful outcomes (25% higher success rate than non-repeat customers).
- **Recommendation**: Create a loyalty program targeting repeat customers, offering incentives like discounts on Internet Packages or expedited loan processing. Use dashboard filters to identify and prioritize these customers, potentially increasing retention rates and revenue by 10%.


