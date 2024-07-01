# Synthesizer

![GitHub Release](https://img.shields.io/github/v/release/engineered-in/Synthesizer)
![GitHub Stars](https://img.shields.io/github/stars/engineered-in/Synthesizer?style=social)
[![Changelog](https://img.shields.io/badge/Changelog-📄-blue)](https://github.com/engineered-in/Synthesizer/blob/main/CHANGELOG.md)
![GitHub Issues](https://img.shields.io/github/issues/engineered-in/Synthesizer)
![GitHub Forks](https://img.shields.io/github/forks/engineered-in/Synthesizer)
![GitHub License](https://img.shields.io/github/license/engineered-in/Synthesizer)
![GitHub All Releases](https://img.shields.io/github/downloads/engineered-in/Synthesizer/total)

Synthesizer is an Excel-based tool designed for Synthesis of Excel Data.

It aims to reduce the repetitive and time-consuming tasks and makes working with Template based excel files a breeze.

<a href="https://github.com/engineered-in/Synthesizer/releases/latest/download/Synthesizer.xlsb" style="display: inline-block; padding: 10px 20px; font-size: 16px; font-weight: bold; text-align: center; color: #fff; background-color: #007bff; border: none; border-radius: 5px; text-decoration: none; cursor: pointer;">
  Download Synthesizer
</a> &nbsp; &nbsp;
<a href="mailto:swarup+synthesizer@engineered.co.in?subject=Synthesizer%20-%20Feedback%20-%20reg.&body=Dear%20Swarup,%0D%0A%0D%0APlease%20find%20below%20my%20feedback%20on%20Synthesizer.xlsb%0D%0A%0D%0AFeedback [Positive/Negative]: %0D%0A%0D%0AComments:" style="display: inline-block; padding: 10px 20px; font-size: 16px; font-weight: bold; color: #ffffff; background-color: #28a745; border: none; border-radius: 5px; text-decoration: none; cursor: pointer;" target="_blank">
  Give Feedback
</a> &nbsp; &nbsp;
<a href="https://www.linkedin.com/company/engineeredin" style="display: inline-block; padding: 10px 20px; font-size: 16px; font-weight: bold; color: white; background-color: #0077b5; border: none; border-radius: 5px; text-decoration: none; text-align: center;">
  <img src="https://cdn.jsdelivr.net/npm/simple-icons@v3/icons/linkedin.svg" alt="LinkedIn" style="width: 20px; height: 20px; vertical-align: middle; margin-right: 8px;"/>
  Engineered-In
</a>

## Features

1. **Generate Multiple Files**:
   - Create datasheets or calculations for multiple items using the same template.
   - Automatically duplicate the template file with the correct data for each calculation.

2. **Generate Summary Table**:
   - Summarize input and output data from multiple workbooks into a single summary table.

## Download

1. Download the <a href="https://github.com/engineered-in/Synthesizer/releases/latest/download/Synthesizer.xlsb" target="_blank">Synthesizer.xlsb</a>.
2. Open the file in Microsoft Excel.
3. Enable macros when prompted to ensure all functionalities work correctly.  Optionally you can save the workbook in a Trusted Location.

## Usage

### A) Generating Multiple Files with a Template

1. **Open Synthesizer**: Open the `Synthesizer.xlsb` file in Excel. A new ribbon "Synthesizer" will appear.
2. **Select Template**: Click on `Select Template` button from the Synthesizer ribbon (available in Map Data group)
3. **Map Data Wizard**: Use the Map Data Wizard to interactively map all the input and output variables from the selected Template File
4. **Generate Summary Table**:
   - Go to the "Summary" sheet.
   - Click "Generate Summary Table" in the Synthesizer ribbon. All variables defined in the "Mapper" sheet will appear as headers.
5. **Input Data**: Copy and paste all the input variable information into the respective columns in the "Summary" sheet.
6. **Set Output Folder**:
   - Select the "Input/Output Folder" icon in the Synthesizer ribbon.
   - Choose a folder where you want the files to be saved.
7. **Select File Format**:
   - Choose the file format (e.g., pdf, xlsx) in the "Export" dropdown.
8. **Generate Files**:
   - Click the "All" button in the Synthesizer ribbon to generate the files.
   - The files will be available in the selected folder.

### B) Creating a Summary from Multiple Workbooks

1. **Open Synthesizer**: Open the `Synthesizer.xlsb` file in Excel. A new ribbon "Synthesizer" will appear.
2. **Select Template**: Click on `Select Template` button from the Synthesizer ribbon (available in Map Data group)
3. **Map Data Wizard**: Use the Map Data Wizard to interactively map all the input and output variables from the selected Template File
4. **Generate Summary Table**:
   - Go to the "Summary" sheet.
   - Click "Generate Summary Table" in the Synthesizer ribbon. All variables defined in the "Mapper" sheet will appear as headers.
5. **Consolidate Workbooks**: Place all workbooks in a single folder.
6. **Set Input Folder**:
   - Select the "Input/Output Folder" icon in the Synthesizer ribbon.
   - Choose the folder containing the workbooks.
7. **Import Data**:
   - Click "Import Data" in the Synthesizer ribbon to populate the "Summary" sheet with data from the workbooks.

<!-- ## Contribution

Contributions are welcome! Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/your-feature`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a pull request. -->

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

Special thanks to all [contributors](https://github.com/engineered-in/Synthesizer/graphs/contributors) and users for their feedback and support.
