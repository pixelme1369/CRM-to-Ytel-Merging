# CRM-to-Ytel-Merging

Matches phone numbers from a Ytel report against a CRM export and enriches the Ytel output with CRM data.

## What It Does

For each row in the Ytel report, the script looks up the dialed phone number in the CRM CSV. When a match is found, it appends the following CRM columns to the output:

| Output Column | CRM Source Column |
|---|---|
| CRM Status | Status |
| Enrolled Debt | Enrolled Debt |
| Last Credit Pulled Date | Last Credit Pulled Date |
| Cordoba Enrolled Date | Cordoba Enrolled Date |
| Credit Score | Credit Score |
| AMOD | ID |
| Assigned To | Assigned To |
| Recording With CRM Status | *(derived — recording filename with CRM status appended)* |

## Usage

```bash
python create_ytel_with_crm_status.py <crm-csv-or-folder> <ytel-xlsx/csv-or-folder> <output.xlsx> [home-phone-column] [status-column]
```

### Arguments

| Argument | Description |
|---|---|
| `crm-csv-or-folder` | Path to the CRM `.csv` file, or a folder — the most recently modified `.csv` is used |
| `ytel-xlsx/csv-or-folder` | Path to the Ytel `.xlsx` or `.csv` file, or a folder — the most recently modified file is used |
| `output.xlsx` | Path for the output Excel file |
| `home-phone-column` | *(optional)* Column name in the CRM CSV containing the phone number (default: `Home Phone`) |
| `status-column` | *(optional)* Column name in the CRM CSV containing the status (default: `Status`) |

### Example

```bash
python create_ytel_with_crm_status.py crm_export.csv ytel_report.xlsx output.xlsx
```

## Requirements

```bash
pip install openpyxl
```
