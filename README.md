# Excel Data Processor for BPS Tables

## Description

This Python script processes Excel files containing BPS (Badan Pusat Statistik) data tables. It filters data based on user input, creates new sheets, and applies specific data processing rules. The script is designed to handle reference and derived tables, applying custom logic to generate a processed output file.

## Author

- **Name:** Fajrian Aidil Pratama
- **Email:** <fajrianaidilp@gmail.com>

## Features

- Processes BPS Excel tables for specific kabupaten (district) codes
- Handles reference and derived tables
- Allows user input for kabupaten code, table codes, and column keywords
- Applies custom processing rules to data (e.g., setting values to 'NA' based on conditions)
- Generates a new Excel file with processed data in multiple sheets (acuan, riil, template)
- Formats animal names for better readability
- Customizes output file names based on kabupaten information

## Requirements

- Python 3.x
- pandas
- openpyxl

## Installation

1. Ensure you have Python 3.x installed on your system.
2. Install required packages:

   ```bash
   pip install pandas openpyxl
   ```

3. Download the `check_na.py` script to your local machine.

## Usage

1. Place your input Excel files in the `data` directory within the script's directory.
2. Run the script:

   ```bash
   python check_na.py
   ```

3. Follow the prompts to enter:

   - Kabupaten code (e.g., 7205)
   - Reference table code (e.g., 6_06)
   - Derived table code (e.g., 6_30)
   - Column keywords (choose from options or enter custom keywords)

4. The script will process the files and generate an output Excel file in the `output` directory.

## Directory Structure

```dir
/CheckNA/
├── check_na.py
├── data/
│   ├── [input Excel files]
└── output/
    └── [processed Excel files]
```

## Customization

- Modify the `directory` variable in the `main()` function to change the base directory path.
- Adjust the `get_column_keywords()` function to modify or add default keyword options.

## Output

The script generates a new Excel file with the following sheets:

- 'acuan': Reference data for the specified kabupaten
- 'riil': Derived data for the specified kabupaten
- 'template': Processed data with applied rules

The output file name is formatted as: `PROCESSED_[Original_Filename]_[KABUPATEN_NAME].xlsx`

## Notes

- Ensure that your input Excel files follow the expected naming convention and sheet structure.
- The script assumes specific column naming patterns. Adjust the code if your data structure differs.
- For best results, use consistent naming conventions across your input files.

## Contributing

Feel free to fork this repository and submit pull requests with any improvements or bug fixes.

## License

This project is open-source and available under the [MIT License](https://opensource.org/licenses/MIT).
