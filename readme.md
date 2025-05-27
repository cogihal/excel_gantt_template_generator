# Excel template file generator for gantt chart

## Description

This is a Python script that generates an Excel template file for a gantt chart.

## How to use

1. Modify the 'config.json' file to set the parameters for the gantt chart that you want to genarate.
1. Run the Python script.
1. After generating the gantt chart template, the script asks you to imput the base file name to save.  
   The extention of the file name is used as '.xlsx' automatically.

## Description of config.json

```
{
  "font_name"  : font name that you want to use
  "tab_title"  : tab title string
  "task_number": how many rows you need for gantt chart
  "start_date" : start date for gantt chart in format YYYY/MM/DD
  "end_date"   : end date for gantt chart in format YYYY/MM/DD

  "holidays": [
    list of holidays in format YYYY/MM/DD
  ]
}
```

## Developing environments

- Python 3.13.3
- openpyxl==3.1.5

