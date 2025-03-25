# Sales-Analysis-Excel-PowerQuery-PivotTables

In this project, I utilized Excel and Power Query to perform an in-depth sales analysis. The goal was to analyze raw sales data, identify key performance metrics, and deliver actionable insights through a dynamic and interactive dashboard.

## Project Workflow

### 1. Data Import & Preparation
- Downloaded AdventureWorks dataset (".bak file") from Microsoft Learn
- Imported the ".bak file" into SQL Server using SSMS
- Connected Excel to SQL Server to import the dataset into Power Query

### 2. Data Transformation in Power Query
Performed comprehensive data cleaning and transformation:
- Removed unnecessary columns
- Handled NULL values effectively
- Merged related columns for better analysis
- Joined tables using "Merge Queries" functionality
- Established proper PK-FK relationships between tables

**Power Query Transformation Result:**  
![Power_Query](https://github.com/Mohammed1999sstack/Sales-Analysis--Excel-PowerQuery-PivotTables/blob/main/Screen%20Shoots/Power%20Query.png?raw=true)

### 3. Data Modeling
Created a STAR SCHEMA dimensional model with:

**FACT Table:**
- **Sales table**: Contains all sales transactions with:
  - SalesOrderID, OrderDate, DueDate, ShipDate
  - CustomerID, SalesPersonID
  - TerritoryID (FK), ProductID (FK)
  - OrderQty, UnitPrice, TotalDue

**DIMENSION Tables:**
1. **SalesTerritory**: TerritoryID + Name
2. **Products**: Product details including:
   - ProductID, SubcategoryID, CategoryID
   - Product Name, Subcategory Name, Category Name

**Data Model Visualization:**  
![Data_Model](https://github.com/Mohammed1999sstack/Sales-Analysis--Excel-PowerQuery-PivotTables/blob/main/Screen%20Shoots/Data%20Modeling.png?raw=true)

### 4. Analysis with PivotTables
Loaded the transformed data into Excel and created multiple PivotTables for analysis:

**PivotTables Screenshot:**  
![Pivot_Tables](https://github.com/Mohammed1999sstack/Sales-Analysis--Excel-PowerQuery-PivotTables/blob/main/Screen%20Shoots/Pivot%20Tables.png?raw=true)

### 5. Dashboard Creation
Developed an interactive dashboard featuring:
- Multiple chart types for comprehensive visualization
- User-friendly interface with slicers and filters
- Key performance metrics and trends

**Final Dashboard:**  
![Dashboard](https://github.com/Mohammed1999sstack/Sales-Analysis--Excel-PowerQuery-PivotTables/blob/main/Screen%20Shoots/Dashboard.png?raw=true)

## Data Source
AdventureWorks2012 dataset from [Microsoft Learn](https://learn.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver16&tabs=ssms)

## Key Skills Applied
- SQL Server data management
- Advanced Power Query transformations
- Dimensional modeling (Star Schema)
- PivotTable analysis techniques
- Interactive dashboard design
- Business intelligence reporting
