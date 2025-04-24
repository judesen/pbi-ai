# PBI TOMWrapper

A .NET library for reading and writing to local Power BI Desktop models using the Tabular Object Model (TOM).

## Overview

This project provides a wrapper around the Microsoft Tabular Object Model (TOM) API specifically designed for working with local Power BI Desktop models. The wrapper simplifies common operations such as:

- Connecting to running Power BI Desktop instances
- Reading model metadata (tables, columns, measures, relationships)
- Creating, updating, and deleting measures
- Adding calculated columns
- Creating relationships between tables
- Refreshing data models

This project is inspired by the [Tabular Editor](https://github.com/TabularEditor/TabularEditor) project.

## Requirements

- .NET 6.0 or later
- Power BI Desktop (latest version recommended)
- Microsoft.AnalysisServices.NetCore.retail.amd64 package
- System.Management package

## Usage

### Finding and Connecting to Power BI Desktop Instances

```csharp
// Create a new wrapper instance
using (var wrapper = new PBIModelWrapper())
{
    // Find all running Power BI Desktop instances
    var instances = wrapper.GetRunningPBIInstances();
    
    if (instances.Count > 0)
    {
        // Connect to the first instance
        wrapper.ConnectToInstance(instances[0]);
        
        // Now you can work with the model
        // ...
    }
}
```

### Creating a Measure

```csharp
// Create a new measure in a table
var measure = wrapper.CreateMeasure(
    "Sales",                  // Table name
    "Total Sales",            // Measure name
    "SUM(Sales[Amount])",     // DAX expression
    "$#,##0.00"               // Format string (optional)
);
```

### Updating a Measure

```csharp
// Update an existing measure
var updatedMeasure = wrapper.UpdateMeasure(
    "Sales",                          // Table name
    "Total Sales",                    // Measure name
    "SUM(Sales[Amount]) + 100",       // New DAX expression
    "$#,##0.00"                       // New format string (optional)
);
```

### Creating a Calculated Column

```csharp
// Create a new calculated column
var column = wrapper.CreateCalculatedColumn(
    "Customers",                      // Table name
    "Full Name",                      // Column name
    "Customers[FirstName] & \" \" & Customers[LastName]", // DAX expression
    DataType.String                   // Data type
);
```

### Working with Tables and Relationships

```csharp
// Get all tables in the model
var tables = wrapper.GetTables();

// Get relationships in the model
var relationships = wrapper.GetRelationships();

// Create a new relationship
var relationship = wrapper.CreateRelationship(
    "Sales",        // From table (many side)
    "CustomerID",   // From column
    "Customers",    // To table (one side)
    "ID"            // To column
);
```

## Demo Console Application

The included console application demonstrates how to use the wrapper to connect to Power BI Desktop and perform various operations on the data model.

## Limitations

- The wrapper can only connect to running instances of Power BI Desktop
- Certain operations may be restricted by Power BI Desktop (such as creating/deleting tables)
- When Power BI Desktop file has a Live Connection, you cannot make structural changes to the model
- This is a simplified implementation for educational purposes

## License

This project is licensed under the MIT License - see the LICENSE file for details. 