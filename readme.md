# Excel template file generator for gantt chart

## Description

This is a Python script that generates an Excel template file for a gantt chart.

## How to use

1. Modify the 'config.toml' file refering the sample TOML file and to set the parameters for the gantt chart that you want to genarate.
1. Run the Python script.
1. The script will generate a gantt chart template in Excel format.
1. After generating the gantt chart template, the script asks you to input the excel base file name to save.  
   The extention of the file name is used as '.xlsx' automatically.

## Description of config.toml

Prepare 'config.toml' by referring the sample TOML file. The configuration file name must be 'config.toml'.

```
# config.toml

font_name   = font name that you want to use to excel : ex. "Meiryo UI"
tab_title   = excel tab title string : ex. "Project Blue"
task_number = how many rows you need for gantt chart
start_date  = start date for gantt chart in format "YYYY/MM/DD"
end_date    = end date for gantt chart in format "YYYY/MM/DD"

holidays = [
  list of holidays in format "YYYY/MM/DD", ...
]
```

