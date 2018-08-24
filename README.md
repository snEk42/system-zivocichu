# Taxonomic units import script

A script enabling automatic import of Class, Order and Family into xlsx file from [ITIS.gov database](https://www.itis.gov/downloads/index.html).
Script assumes the same input format as in `example.xlsx` and that all subjects belong to the Animalia kingdom.
Matches are primarily made with Genus and Species name combination, if that fails only Genus is matched. If
even Genus match fails for a subjects imported fields are left blank. Script orders subjects by their latin name
and duplicate rows are left in the output. Script can hadle multiple sheets, but will unfortunatelly drop data formating.
All fields are trimmed of whitespaces at the beginning and end. Latin names will have first letter capitalised on output.

## How to run the script

1. Install Node.js latest stable release (<https://nodejs.org/en/>).
2. Download latest SQLite version of full [ITIS.gov database](https://www.itis.gov/downloads/index.html) and extract `ITIS.sqlite` source file into the root of this repository.
3. Rename the file you want to import taxonomic units for as `data.xlsx` and copy it to the root of this repository.
4. In the root of this repository run `node import.js`.
5. Wait while the script fills in the table. This will take minute or two per each 1000 rows, so it might take a while. Unmatched animals are reported in the script output.
6. Once the script finishes, you'll find your filled in tables in `filled.xlsx`. The `data.xlsx` file will be left untouched.
