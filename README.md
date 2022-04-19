# csv2xlsx

## Usage: 
```bash
csv2xlsl {input.csv} {output.xlsx} [settings.json]
```

## Example settings file
```json
{
    "width": 2000,
    "height": 700,
    "first_row_freeze": true,
    "first_row_bold": true,
    "first_row_height": 20,
    "first_row_autofilter": true,
    "column_widths": [100,20,5,20,30],
    "text_wrap": true
}
```