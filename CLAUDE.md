# CLAUDE.md — CRM-to-Ytel-Merging

## What This Project Does

Matches phone numbers from a Ytel report (`.xlsx` or `.csv`) against a CRM export (`.csv`) and writes an enriched Excel output with CRM data appended to each matched row.

## Main Script

`create_ytel_with_crm_status.py`

### Usage
```bash
python create_ytel_with_crm_status.py <crm-csv-or-folder> <ytel-xlsx/csv-or-folder> <output.xlsx> [home-phone-column] [status-column]
```
- If a folder is given for CRM or Ytel, the most recently modified matching file is used automatically.
- Default CRM phone column: `Home Phone`
- Default CRM status column: `Status`

## CRM Columns Imported into Output

Defined in `CRM_IMPORT_COLUMNS` (each entry: `(crm_column, output_column, required)`):

| CRM Column | Output Column | Required |
|---|---|---|
| Enrolled Debt | Enrolled Debt | Yes |
| Last Credit Pulled Date | Last Credit Pulled Date | Yes |
| Cordoba Enrolled Date | Cordoba Enrolled Date | Yes |
| Credit Score | Credit Score | Yes |
| ID | AMOD | Yes |
| Assigned To | Assigned To | No |
| State | CRM_State | No |

- **Required = Yes**: script raises an error if the column is missing from the CRM CSV.
- **Required = No**: script prints a warning and leaves the column blank if missing.
- `CRM_State` is intentionally named differently from Ytel's own `State` column to avoid overwriting it.

## Output Columns Added by This Script

Appended after the existing Ytel columns on every matched row:

- `CRM Status` — cleaned status value from CRM
- `Enrolled Debt`
- `Last Credit Pulled Date`
- `Cordoba Enrolled Date`
- `Credit Score`
- `AMOD` (from CRM `ID` column)
- `Assigned To`
- `CRM_State` (from CRM `State` column — separate from Ytel's `State`)
- `Recording With CRM Status` — recording filename/path with CRM status appended

## Columns Removed from Output

Defined in `REMOVE_OUTPUT_COLUMNS` — Ytel columns stripped before saving to keep the output clean. Includes: `call_list_id`, `vendor_lead_code`, `list_id`, `gmt_offset_now`, `phone_code`, `title`, `gender`, `date_of_birth`, `alt_phone`, `email`, `security_phrase`, `comments`, `list_name`, `owner`, `postal_code`, `rank`, `country_code`, `province`, `state`, `city`, `address1`, `address2`, `address3`, `middle_initial`.

## Adding New CRM Columns

To bring in an additional CRM column, add a line to `CRM_IMPORT_COLUMNS` in the script:
```python
("CRM Column Name", "Output Column Name", False),  # False = optional, True = required
```
Also add the output column name to `EXTRA_CRM_COLUMNS` for reference.

## Dependencies

```bash
pip install openpyxl
```
