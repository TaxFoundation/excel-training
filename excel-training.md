# Tax Foundation Excel Bootcamp

This training course is designed for incoming interns, but can serve as a handy refresher for any person looking to up their Microsoft Excel game.

## Table of Contents

* What's the Goal of This Guide?
* Organizing Messy Data
  * Text-to-columns to Separate Conjoined Values
  * Number Formats to Enhance Presentation and Prevent Weird Errors
  * Advanced Sorting to Arrange Things How You Want Them
  * Transposition to Flip Everything Around
* Using Basic Formulas
  * Add, Subtract, Multiply, Divide
* Using Advanced Formulas
  * Using VLOOKUP to Match Data from Different Tables
  * Using IF to Conditionally Generate Data
* [PivotTables to Summarize and Organize Data](#pivot-tables)
  * [Create Your First PivotTable](#first-pivot-table)
  * [Choosing What Your PivotTable Displays](#choose-pivot-values)
  * [Filtering by Criteria](#filter-pivot-table)
* [File > Save As](#file-save-as)
  * [XLSX](#xlsx)
  * [CSV](#csv)

## What's the Goal of This Guide?

People come to the Tax Foundation with all sorts of backgrounds and varying levels of experience. Not everyone is an Excel whiz. This guide is meant to go over the most commonly used Excel features for managing the datasets we work with.

### There is a Wrong Way and a Right Way to Excel

![An inexperienced Excel user.](/images/excel-wrong.gif)

There are many ways to complete tasks in Excel, but some are better than others. Many tasks can be done through monotonous repetition. However, choosing this path wastes your time and crushes your soul. It's much better to invest some time upfront to learn about Excel features that will speed up your work and automate repetitive tasks.

![An experienced Excel user.](/images/excel-right.gif)

## <a id="pivot-tables"></a> PivotTables to Summarize and Organize Data

PivotTables are a very cool Excel feature that allows you to create summary tables of your data. For example, let's say you have a dataset that lists tax collections at the county level and you want to know the sum of collections at the state level. You could devise a series of formulas to get this information, but it's much easier to quickly build a PivotTable. Let's dig into it!

### <a id="first-pivot-table"></a>Create Your First PivotTable

1. Select all of the data you intend to summarize.
2. Go to `Insert > Tables > PivotTable`
3. By default, PivotTables are created on new worksheets. The default settings are usually fine, so go ahead and click `OK` to create your PivotTable.

![Creating a PivotTable](/images/create-pivot-table.gif)

### <a id="choose-pivot-values"></a>Choosing What Your PivotTable Displays

A blank PivotTable isn't much help. You need to select which fields you want to summarize and how. In our example of summarizing count tax collections by state, we want to choose state for our rows and the taxes as our values. We can easily drag our chosen fields into the sections we want them using the PivotTable pane.

By default, the number in Values will be a count summary. We don't need to know how many counties are in each state! For our example, we select the Value and choose `Value Field Settings...` to change from `Count` to `Sum`. We also change the number format to `Currency`, since we know we're working with dollar values.

There are many ways to summarize and format data values, and the correct one will vary from project to project. Don't be afraid to play around with it and see what you can do!

![Choosing What the PivotTable Displays](/images/choose-pivot-values.gif)

### <a id="filter-pivot-table"></a>Filtering by Criteria

Sometimes you'll want to filter the PivotTable results. You can easily do this by dragging-and-dropping the field you want to filter by into `FILTERS` in the PivotTable pane. This will add the field to a list of filters above the PivotTable, where you can fine-tune the criteria to filter by.

![Filter by Criteria](/images/filter-pivot-table.gif)

## <a id="file-save-as"></a>File > Save As

Truthfully, you should've been saving your work as you went along! There's no telling when Excel might crash, destroying all of your hard work.

But this section isn't just a reminder to save! Now that you've wrangled your data, you need to make sure it's in a ready-to-use file format. You'll typically be working with two formats: **Excel** and **CSV**. Here's what you need to know about each.

### <a id="xlsx"></a>XLSX

The default file used by Excel is the **XLSX** file. This filetype is best for preserving constructs that help you clean your data, such as formulas and PivotTables. **Using XLSX is best while the data you're cleaning is a work-in-progress.**

### <a id="csv"></a>CSV

In almost all cases, you're saving the final file as a **comma separated values** file, or **CSV**. You will save the final CSV files separate from the WIP Excel files you used previously.

A CSV is exactly what the name implies: values with commas in between, and line-breaks between rows. **This is the preferred format for final individual data tables** because it strips out all of Excel's magic and leaves an easy-to-use, platform-agnostic dataset. No formulas, no PivotTables, no number formats, no font formats, not even separate worksheets. CSV is clean and simple.

#### Example CSV File

```
id,value,percent
1,347.05,0.14
2,937.56,0.32
```

#### Saving As CSV

In Excel, go to `File > Save As` and choose `CSV (Comma Delimitted)` as the file type.

Because saving as CSV means losing all of the special Excel magic in the file, Excel will warn you about saving CSVs every single time. Be patient, and tell it, yes, you really, truly do want to save as CSV.
