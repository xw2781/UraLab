2027/9/1

# Set model to Opus 4.8 (1M context) for a session

/model claude-opus-4-8[1m]
/model claude-opus-4-8
/model claude-opus-4-6

# Commit and Push Workflow

git pull --ff-only

Run the commit workflow.
Run the commit workflow and push.

[$arcrho-commit-workflow](E:\\XWSpace\\Repos\\ArcRho\\.claude\\skills\\arcrho-commit-workflow\\SKILL.md)


# JSON Contract Validations

Macros
"Import ResQ Reserving Class"
"Import ResQ Reserving Classes"
"Export Reserving Class to ResQ"
"Sync Reserving Class with ResQ"


Project: NJ_Annual_Prod_202605_Fake
Path: PRNJ - PA\PA\All States\Direct Group\COL

Project: NJ_Annual_Prod_202605_Fake
Path: HPPREF\HO+DF\NJ\Legacy\HOL


# Server Components Rebuild and Deploy

Run the commit workflow, rebuild and deploy server components if needed.
server-components\deploy.bat


# Production Q3-Aug

## Create 4 Adjustment Vectors

& "C:\Program Files\Python310\python.exe" tools/create_reserve_review_input_datasets.py --project "NJ_Annual_Prod_2026 Q2-May Test" --quarter "2026Q2" --dry-run


## Reconciliations

### DFM triangles
py -3.10 python-api/migration/validation/dfm_ratio_side_by_side_review.py

### Result Selections
py -3.10 python-api/migration/validation/rs_dataset_side_by_side_review.py

### Datasets (no method)
py -3.10 python-api/migration/validation/dataset_side_by_side_review.py

py -3.10 python-api/migration/validation/dataset_side_by_side_review.py --source-kind all --rc "Legacy\HOL"

### Datasets + RS
py -3.10 python-api/migration/validation/combined_side_by_side_review.py --rc "HPPREF\HO+DF\NJ\Legacy\HOL"


