# PowerPoint Chart Generator

I’m building this tool because its been noticed how much time goes into manually turning Excel data into PowerPoint charts. So my goal is to automate that process: upload an Excel file, pick chart types, and generate a branded PowerPoint presentation automatically.

## What It Does (so far)

I wanted the project to be dynamic and to work with any excel file since most project data would be named differently. 
While it imports the data we also needed some cleaning on the data before generating the charts. 
Things like nAn fields or blanks had to accounted for and ignored. 
Then the idea was to be able to generate the data into one of the following charts.
* Bar
* Column
* Pie
* Line
* Area
* Doughnut
* Stacked Bar
* Stacked Column

The trick was then to try and be able to choose which chart for which slide while preserving the initail pages of the powerpoint.

## The Interface

The idea for interface was to keep it simple and practical. I dont think it came out exactly as I intended but im working on it. 
There is a scrollable list of all excel sheets picked up on the import.
On this list we are able to enable/disable if I want to use the data for charting, and if I do want to use to use it I can select which chart to generate.

## How It Works

1. Load your Excel file and a PowerPoint template.
2. Choose where the generated slides should start.
3. Pick chart types for each sheet.
4. Click “Generate” — the slides are built automatically.

## Still in Development

Right now, I’m working on:

* overall styling and structure so that it meets company standards
* Better handling of large Excel files with 50+ sheets.
* Smoother error recovery when Excel data is messy.
* More customization for colors and layouts.

## Tech Behind It

* Python 3.7+
* pandas + openpyxl (for Excel)
* python-pptx (for PowerPoint)
* tkinter (for the interface)

## My Goal

Even though it’s not part of my day-to-day job, I can see how useful it would be for teams that still build every chart manually. 
The idea is to save them time, reduce mistakes, and let them focus on the real work — interpreting and presenting insights.

