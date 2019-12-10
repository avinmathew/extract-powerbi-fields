# Extract Power BI Fields

This PowerShell script will extract a list of fields from a folder that contains .pbix or .pbit files. This is useful when undertaking impact analysis when changing a data model, to see which reports will be impacted when removing or renaming fields/tables.

## Usage

Extract fields to `fields.csv` for Power BI reports in the current directory:

```powershell
Extract-Fields
```

Extract fields to a named CSV for Power BI reports in a specified directory:

```powershell
Extract-Fields -ReportsPath C:\reports -CsvPath C:\reports\pbifields.csv
```

To overwrite an existing CSV file, use the `-Force` parameter.

## Limitations

Due to the complexities of Layout structure with filters on visuals, I've taken a shortcut and used regex rather than JSON navigation to mop up any missing properties not found the the first two passes. This will result in the Table column in the CSV file being blank in these cases.
