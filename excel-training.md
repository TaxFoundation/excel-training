# Tax Foundation Excel Bootcamp

This training course is designed for incoming interns, but can serve as a handy refresher for any person looking to up their Microsoft Excel game.

## Table of Contents

* [What's the Goal of This Guide?](#what-s-the-goal-of-this-guide)
* [Organizing Messy Data](#organizing-messy-data)
  * [Text-to-columns to Separate Conjoined Values](#text-to-columns-to-separate-conjoined-values)
  * [Number Formats to Enhance Presentation and Prevent Weird Errors](#number-formats-to-enhance-presentation-and-prevent-weird-errors)
  * [Advanced Sorting to Arrange Things How You Want Them](#advanced-sorting-to-arrange-things-how-you-want)
  * [Transposition to Flip Everything Around](#transposition-to-flip-everything-around)
* [Using Basic Formulas](#using-basic-formulas)
  * [Add, Subtract, Multiply, Divide](#add-subtract-multiply-divide)
* [Using Advanced Formulas](#using-advanced-formulas)
  * [Using VLOOKUP to Match Data from Different Tables](#using-vlookup-to-match-data-from-different-tables)
  * [Using IF to Conditionally Generate Data](#using-if-statements-to-conditionally-generate-data)
* [PivotTables to Summarize and Organize Data](#pivottables-to-summarize-and-organize-data)
  * [Create Your First PivotTable](#create-your-first-pivottable)
  * [Choosing What Your PivotTable Displays](#choosing-what-your-pivottable-displays)
  * [Filtering by Criteria](#filtering-by-criteria)
* [File > Save As](#file-save-as)
  * [XLSX](#xlsx)
  * [CSV](#csv)

## What's the Goal of This Guide?

People come to the Tax Foundation with all sorts of backgrounds and varying levels of experience. Not everyone is an Excel whiz. This guide is meant to go over the most commonly used Excel features for managing the datasets we work with.

### There is a Wrong Way and a Right Way to Excel

![An inexperienced Excel user.](/images/excel-wrong.gif)

There are many ways to complete tasks in Excel, but some are better than others. Many tasks can be done through monotonous repetition. However, choosing this path wastes your time and crushes your soul. It's much better to invest some time upfront to learn about Excel features that will speed up your work and automate repetitive tasks.

![An experienced Excel user.](/images/excel-right.gif)

## Using Advanced Formulas

Now that you've got the easy bits under your belt, let's make some *really* interesting formulas! Excel formulas allow for a lot of programmatic logic to make [data munging](https://en.wikipedia.org/wiki/Data_wrangling) much easier. Let's look at some of the most useful formulas.

### Using VLOOKUP to Match Data from Different Tables

**VLOOKUP** is a useful function for finding a value from one table based on a value in another table.

For example, let's say you have a table of state-level data where the states are identified by [FIPS code](https://www.census.gov/geo/reference/ansi_statetables.html). You want the data to be identified by the full state name instead. Luckily, you've got a table that matches FIPS codes to full state names. Should you use that table as a reference, manually replacing FIPS codes in your data set with state names? No, never! We'll use VLOOKUP to to match the full state name to the FIPS code.

#### Example Data

![On the left, state-level data by FIPS code. On the right, state names by FIPS code.](/images/vlookup-data.png)

#### Writing Your VLOOKUP

The VLOOKUP formula has four parts:

1. A reference to the cell whose value you're looking to match.
2. A reference to the table you're matching that cell against.
3. The column number you want to retrieve from that table.
4. Whether or not a close, but not exact, match is acceptable. (Hint: This should always be set to FALSE for our purposes.)

The final formula in our example might look like this:

![VLOOKUP formula that gives us full state names for FIPS codes.](/images/vlookup-example.png)

Some important things to note:

* The first value is a relative cell reference to `A2`. When you copy this formula down through the column, it will automatically update to look for `A3`, `A4`, etc.
* The second value is an absolute table reference. The `$` before the letter and number tell Excel to keep these values exactly the same when this formula is copied. This way, cells `C2:C51` will always be looking for values in the precise location of the reference table, `Sheet2!$A$2:$B$52`. If you don't make this reference absolute, you'll end up looking for the value in `C51` in a table of mostly empty cells!
* Column numbers for part three begin at 1 and count up.
* The values we're searching through in the second table are sorted in ascending order, and the value we're looking for is in the first column. This is necessary with VLOOKUP.
* The data we want to modify does not include the District of Columbia, but the FIPS reference table does. It's OK if something in the reference table doesn't have a match in the table where we're using VLOOKUP. The opposite situation--a value in our data that doesn't exist in the reference table--will throw and error.

![VLOOKUP Demonstration](/images/vlookup-demo.gif)

## Using IF Statements to Conditionally Generate Data

## PivotTables to Summarize and Organize Data

PivotTables are a very cool Excel feature that allows you to create summary tables of your data. For example, let's say you have a dataset that lists tax collections at the county level and you want to know the sum of collections at the state level. You could devise a series of formulas to get this information, but it's much easier to quickly build a PivotTable. Let's dig into it!

### Create Your First PivotTable

1. Select all of the data you intend to summarize.
2. Go to `Insert > Tables > PivotTable`
3. By default, PivotTables are created on new worksheets. The default settings are usually fine, so go ahead and click `OK` to create your PivotTable.

![Creating a PivotTable](/images/create-pivot-table.gif)

### Choosing What Your PivotTable Displays

A blank PivotTable isn't much help. You need to select which fields you want to summarize and how. In our example of summarizing count tax collections by state, we want to choose state for our rows and the taxes as our values. We can easily drag our chosen fields into the sections we want them using the PivotTable pane.

By default, the number in Values will be a count summary. We don't need to know how many counties are in each state! For our example, we select the Value and choose `Value Field Settings...` to change from `Count` to `Sum`. We also change the number format to `Currency`, since we know we're working with dollar values.

There are many ways to summarize and format data values, and the correct one will vary from project to project. Don't be afraid to play around with it and see what you can do!

![Choosing What the PivotTable Displays](/images/choose-pivot-values.gif)

### Filtering by Criteria

Sometimes you'll want to filter the PivotTable results. You can easily do this by dragging-and-dropping the field you want to filter by into `FILTERS` in the PivotTable pane. This will add the field to a list of filters above the PivotTable, where you can fine-tune the criteria to filter by.

![Filter by Criteria](/images/filter-pivot-table.gif)

## File > Save As

Truthfully, you should've been saving your work as you went along! There's no telling when Excel might crash, destroying all of your hard work.

But this section isn't just a reminder to save! Now that you've wrangled your data, you need to make sure it's in a ready-to-use file format. You'll typically be working with two formats: **Excel** and **CSV**. Here's what you need to know about each.

### XLSX

The default file used by Excel is the **XLSX** file. This filetype is best for preserving constructs that help you clean your data, such as formulas and PivotTables. **Using XLSX is best while the data you're cleaning is a work-in-progress.**

### CSV

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
