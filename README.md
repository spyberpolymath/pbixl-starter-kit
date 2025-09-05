# Power BI & Excel Data Analytics Project

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Power BI](https://img.shields.io/badge/Power%20BI-Compatible-yellow)](https://powerbi.microsoft.com/)
[![Excel](https://img.shields.io/badge/Excel-2019%2B-green)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen)](https://github.com/spyberpolymath/pbixl-starter-kit)

## 📊 Overview

A comprehensive data analytics learning project that demonstrates professional business intelligence workflows using **Microsoft Power BI** and **Microsoft Excel**. This project provides hands-on experience with real-world datasets spanning multiple business domains including Sales, Marketing, Finance, HR, and Operations.

### 🎯 Key Features

- **Multi-domain dataset** with realistic business scenarios
- **End-to-end analytics workflow** from data cleaning to visualization
- **Integrated Excel and Power BI solutions** showcasing best practices
- **Production-ready dashboards** with interactive filtering and KPIs
- **Comprehensive documentation** and step-by-step tutorials

## 🚀 Quick Start

### Prerequisites

- Microsoft Excel 2019+ or Microsoft 365
- Power BI Desktop (free download)
- Basic understanding of data analysis concepts

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/spyberpolymath/pbixl-starter-kit.git
   cd pbixl-starter-kit
   ```

2. **Open the dataset**
   ```bash
   # Open the main Excel file
   start bi_excel_power_dataset.xlsx
   ```

3. **Launch Power BI Desktop** and import the dataset to begin building dashboards

```

## 📈 Dataset Overview

### Business Domains Covered

---------------------------------------------------------------------------------------------------------------------------------------------------------------
| Domain              | Tables                                        | Key Metrics                           | Use Cases                                     |
|---------------------|-----------------------------------------------|---------------------------------------|-----------------------------------------------|
| **Sales**           | `fact_sales`, `dim_product`, `dim_region`     | Revenue, Quantity, Profit Margin      | Sales performance tracking, regional analysis |
| **Marketing**       | `marketing_campaigns`, `customer_acquisition` | Campaign ROI, Conversion rates        | Marketing effectiveness analysis              |
| **Finance**         | `budget_actuals`, `financial_metrics`         | Budget variance, P&L analysis         | Financial reporting and forecasting           |
| **Human Resources** | `employee_data`, `performance_metrics`        | Headcount, Attrition rates            | Workforce analytics                           |
| **Operations**      | `order_processing`, `delivery_metrics`        | Processing time, Delivery performance | Operational efficiency tracking               |
---------------------------------------------------------------------------------------------------------------------------------------------------------------
### Data Model

The dataset follows a **star schema** design with:
- **Fact Tables**: Transaction-level data (sales, orders, campaigns)
- **Dimension Tables**: Reference data (products, regions, dates, employees)
- **Relationships**: Properly defined foreign key relationships for optimal performance

## 🛠️ Learning Path

### Phase 1: Excel Mastery
Learn essential Excel skills for data analytics:

#### Core Functions & Formulas
```excel
# Lookup Functions
=VLOOKUP(A2, ProductTable, 2, FALSE)
=INDEX(MATCH()) combinations

# Conditional Logic
=IF(B2>1000, "High", "Low")
=IFERROR(C2/D2, 0)

# Aggregation Functions
=SUMIF(Region, "North", Sales)
=COUNTIFS(Date, ">="&DATE(2023,1,1), Status, "Completed")
```

#### Advanced Techniques
- **PivotTables**: Dynamic data summarization
- **PivotCharts**: Automated visualization
- **Conditional Formatting**: Visual data insights
- **Data Validation**: Input controls and quality
- **Slicers & Filters**: Interactive reporting

#### Excel Dashboard Example
![Excel Dashboard Preview](documentation/images/excel_dashboard_preview.png)

### Phase 2: Power BI Development

#### Data Preparation
1. **Connect to Data Sources**
   - Excel files, CSV files, databases
   - Data refresh configuration

2. **Power Query Transformations**
   ```powerquery
   // Remove duplicates and clean data
   = Table.Distinct(Source)
   = Table.TransformColumnTypes(CleanedData, {{"Date", type date}, {"Amount", Currency.Type}})
   ```

3. **Data Modeling**
   - Create relationships between tables
   - Define calculated columns and measures

#### DAX Calculations
```dax
// Key Performance Indicators
Total Sales = SUM(fact_sales[SalesAmount])
Profit Margin = DIVIDE([Total Profit], [Total Sales])
YoY Growth = 
    VAR CurrentYear = [Total Sales]
    VAR PreviousYear = CALCULATE([Total Sales], SAMEPERIODLASTYEAR(dim_date[Date]))
    RETURN DIVIDE(CurrentYear - PreviousYear, PreviousYear)

// Time Intelligence
Sales YTD = TOTALYTD([Total Sales], dim_date[Date])
Sales QTD = TOTALQTD([Total Sales], dim_date[Date])
```

#### Visualization Best Practices
- **KPI Cards**: Key metrics at-a-glance
- **Bar/Column Charts**: Categorical comparisons
- **Line Charts**: Trends over time
- **Maps**: Geographic insights
- **Tables/Matrix**: Detailed data views
- **Slicers**: Interactive filtering

### Phase 3: Integration Mastery

#### Excel → Power BI Integration
```powerquery
// Connect Excel workbook as data source
Source = Excel.Workbook(File.Contents("bi_excel_power_dataset.xlsx"), null, true)
```

#### Power BI → Excel Integration
- **Analyze in Excel**: Direct connection to Power BI datasets
- **Export Data**: Move insights back to Excel for detailed analysis
- **Embedding**: Power BI visuals in Excel reports

## 📊 Sample Dashboards

### Sales Performance Dashboard
- **Revenue Trends**: Monthly/quarterly sales analysis
- **Regional Performance**: Geographic breakdown of sales
- **Product Analysis**: Top-performing products and categories
- **Customer Segmentation**: High-value customer identification

### Financial Analysis Report
- **Budget vs Actuals**: Variance analysis across departments
- **Profit & Loss**: Income statement visualization
- **Cash Flow**: Monthly cash flow projections
- **Cost Center Analysis**: Department-wise expense tracking

### HR Analytics Dashboard
- **Workforce Metrics**: Headcount trends and demographics
- **Performance Management**: Employee performance tracking
- **Attrition Analysis**: Turnover patterns and retention insights
- **Recruitment Funnel**: Hiring process effectiveness

## 🔧 Advanced Features

### Power BI Service Integration
```bash
# Publish to Power BI Service
pbix-cli publish --workspace "Analytics Workspace" --file "sales_dashboard.pbix"
```

### Automated Data Refresh
- **Scheduled Refresh**: Daily/weekly data updates
- **Incremental Refresh**: Efficient large dataset handling
- **Real-time Streaming**: Live data connections

### Security & Governance
- **Row-Level Security (RLS)**: Data access controls
- **Sensitivity Labels**: Data classification
- **Audit Logs**: Usage tracking and compliance

## 📚 Documentation

-----------------------------------------------------------------------------------------------------------------
| Resource                                                | Description                          | Level        |
|---------------------------------------------------------|--------------------------------------|--------------|
| [Excel Tutorial](documentation/excel_tutorial.md)       | Step-by-step Excel analytics guide   | Beginner     |
| [Power BI Tutorial](documentation/powerbi_tutorial.md)  | Complete Power BI dashboard creation | Intermediate |
| [Integration Guide](documentation/integration_guide.md) | Excel + Power BI best practices      | Advanced     |
| [DAX Reference](documentation/dax_reference.md)         | Common DAX formulas and patterns     | All Levels   |
-----------------------------------------------------------------------------------------------------------------

### Development Setup
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📋 Use Cases & Applications

### Business Intelligence
- **Executive Dashboards**: C-level reporting and KPIs
- **Department Analytics**: Function-specific insights
- **Operational Reporting**: Daily/weekly operational metrics

### Industry Applications
- **Retail**: Inventory management, sales forecasting
- **Healthcare**: Patient analytics, operational efficiency
- **Education**: Student performance, enrollment tracking
- **Manufacturing**: Production metrics, quality control
- **Financial Services**: Risk analysis, regulatory reporting

## 🎯 Learning Outcomes

Upon completion of this project, you will have:

✅ **Technical Skills**
- Advanced Excel analytics and dashboard creation
- Power BI development and DAX programming
- Data modeling and relationship management
- Interactive visualization design

✅ **Business Skills**
- Business intelligence best practices
- KPI development and monitoring
- Data storytelling and presentation
- Cross-functional analytics collaboration

✅ **Portfolio Assets**
- Professional dashboards for job interviews
- Real-world analytics project experience
- GitHub repository showcasing technical skills
- Documentation demonstrating communication abilities

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/spyberpolymath/pbixl-starter-kit/issues)
- **Discussions**: [GitHub Discussions](https://github.com/spyberpolymath/pbixl-starter-kit/discussions)
- **Email**: amananiloffical@gmail.com

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Power BI team for excellent documentation
- Excel community for best practices and tutorials
- Open source data analytics community
- Contributors and beta testers

---

**⭐ If this project helped you learn data analytics, please give it a star!**