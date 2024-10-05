
# Bulk-Calculate&nbsp;<img src="https://github.com/user-attachments/assets/cfa7e6b8-202f-4ba0-b48e-bafe7a3d5cf6" height="26px">

<i>Work with data, not files.</i>

[Download Bulk-Calculate](https://github.com/engineered-in/Bulk-Calculate/releases/latest/download/Bulk-Calculate.xlsb)&nbsp; | &nbsp;
[Walkthrough](https://view.genially.com/66ef09bc2d8d928848f09bb2/interactive-content-bulk-calculate-getting-started-guide) &nbsp; | &nbsp;
<a href="mailto:swarup+bulk-calculate@engineered.co.in?subject=Bulk-Calculate%20-%20Feedback%20-%20reg.&body=Dear%20Swarup,%0D%0A%0D%0APlease%20find%20below%20my%20feedback%20on%20Bulk-Calculate.xlsb%0D%0A%0D%0AFeedback [Positive/Negative]: %0D%0A%0D%0AComments:"  target="_blank">Give Feedback</a> &nbsp; | &nbsp;
<a href="https://github.com/sponsors/engineered-in" target="_blank">Sponsor Engineered-In</a>

![GitHub Release](https://img.shields.io/github/v/release/engineered-in/Bulk-Calculate)&nbsp;![GitHub Stars](https://img.shields.io/github/stars/engineered-in/Bulk-Calculate?style=social)&nbsp;[![Changelog](https://img.shields.io/badge/Changelog-📄-blue)](https://github.com/engineered-in/Bulk-Calculate/blob/main/CHANGELOG.md)&nbsp;![GitHub Issues](https://img.shields.io/github/issues/engineered-in/Bulk-Calculate)&nbsp;![GitHub Forks](https://img.shields.io/github/forks/engineered-in/Bulk-Calculate)&nbsp;![GitHub License](https://img.shields.io/github/license/engineered-in/Bulk-Calculate)&nbsp;![GitHub All Releases](https://img.shields.io/github/downloads/engineered-in/Bulk-Calculate/total)&nbsp;

## What is Bulk Calculate?

**Bulk-Calculate** is a tool designed to streamline repetitive, **template**-based engineering calculations by centralizing inputs and outputs in a single summary table. 

Ideal solution for organizations transitioning to a **data-driven workflow** using Excel.

## What Problem Does It Solve?

In many engineering projects, template-driven calculations are routine, such as the design of structural elements like beams. 
Typically, these calculations are done using a standard, validated Excel file (referred to as a "**Template**") that is **reused** for **multiple calculations** across projects.
While the calculation process remains consistent across different projects, only project-specific details (e.g., client name, logo) change.


However, the conventional approach of creating separate Excel files for each calculation leads to several issues:

&nbsp;&nbsp;❌ **File Multiplication** - Managing multiple similar files and updating each one manually is a time drain.  
&nbsp;&nbsp;❌ **Data Accuracy Risks** - Errors can easily occur when manually entering or pasting data across multiple similar looking files.  
&nbsp;&nbsp;❌ **Data Entry Overload** - Excessive time and effort is spent on data entry instead of engineering innovation.   
&nbsp;&nbsp;❌ **Inefficient Data Extraction** - Summarizing results requires manually opening and extracting data from each file one-by-one.  
&nbsp;&nbsp;❌ **Time-consuming Updates** - Any changes to the template must be applied manually to every file.  
&nbsp;&nbsp;❌ **Losing Sight of the Big Picture** - The scattered nature of data across files makes it difficult to see the overall picture.   
&nbsp;&nbsp;❌ **Lack of Data Visibility** - There's no centralized view of all the calculations, making it difficult to track or reference past work.  
  

For example, if I need to check if a similar beam design was performed in another project, I must search through numerous files, which is difficult and inefficient.

### How Does It Solve the Problem?

A more efficient solution would be a centralized system that records all inputs and outputs from each calculation in a single summary table, with each row representing a calculation file.
This provides:

&nbsp;&nbsp;✅ Clear visibility of data  
&nbsp;&nbsp;✅ Easier access to past calculations  
&nbsp;&nbsp;✅ More efficient updates across projects  
&nbsp;&nbsp;✅ Ability to export the summary table to spreadsheets or PDFs, preserving the look of the conventional method 


## Key Features

- **Template-based calculations**: Use any of your calculation file as a template
- **Mapping Wizard**: Easily map calculation input and output cells
- **Summary table**: Get a clear view of your calculation inputs and outputs
- **Bulk calculation**: Perform multiple calculations at once with a single click
- **Bulk export**: Export individual calculation spreadsheets or PDF files from Summary table
- **Bulk import**: Import data from existing calculation spreadsheets into a Summary table


## Download

1. Download the <a href="https://github.com/engineered-in/Bulk-Calculate/releases/latest/download/Bulk-Calculate.xlsb" target="_blank">Bulk-Calculate.xlsb</a> to a [Trusted Location](https://github.com/engineered-in/Bulk-Calculate/wiki/Excel-Trusted-Location).
2. Open the file in Microsoft Excel.
3. Enable macros when prompted to ensure all functionalities work correctly.

[Demo Video Playlist](https://www.youtube.com/watch?v=J667nX5zhAE&list=PLEv5wGuO-nlCG0vGYjktEjpwVfhTBWX8P)

## How to Use?

### 1. Selecting template spreadsheet

- Open the `Bulk-Calculate.xlsb` file in Excel. A new ribbon "Bulk-Calculate" will appear.
- Click on `Select Template` button from the Bulk-Calculate ribbon (available in Map Data group)

### 2. Interactively map your template using Map Data Wizard

- Use the `Map Data Wizard` to interactively map all the input and output variables from your selected Template File

### 3. Generate Summary Table

- Go to the "Summary" sheet.
- Click `Generate Summary Table` in the Bulk-Calculate ribbon. All variables defined in the "Mapper" sheet will appear as headers.

### 4. Fill the Summary Table with your Inputs

- Copy and paste all the input variable information into the respective columns in the "Summary" sheet.

### 5. Bulk Calculate

- Click on the `Calculate All` button to perform bulk calculation on all the records of the Summary Table

### 6. Bulk Export

- Select the `Input/Output Folder` icon in the Bulk-Calculate ribbon.
- Choose a folder where you want the files to be saved.
- Choose the file format (e.g., pdf, xlsx) in the "Export" dropdown.
- Click the `Export All` button in the Bulk-Calculate ribbon to bulk export the files.
- The files will be available in the selected folder.

### 7. Bulk Import

- Place all workbooks in a single folder.
- Select the "Input/Output Folder" icon in the Bulk-Calculate ribbon.
- Choose the folder containing the workbooks.
- Click "Import Data" in the Bulk-Calculate ribbon to populate the "Summary" sheet with data from the workbooks.

<!-- ## Contribution

Contributions are welcome! Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/your-feature`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a pull request. -->

## Frequently Asked Questions

<details><summary>Is Bulk-Calculate free?</summary>

Yes, Bulk-Calculate is completely free and open-source. You can use, view the code, and even modify it to suit your needs without any cost (except for the Microsoft Excel license ofcourse).
</details>

<details><summary>Why should I trust Bulk-Calculate?</summary>

Bulk-Calculate relies on macros, which are disabled by default in Excel unless you trust the file or its publisher. As an open-source project, the VBA code is fully transparent, allowing anyone to review it for vulnerabilities. Only verified maintainers (using GPG keys) can update the source code and release new versions.
</details>

<details><summary>Is it valuable even though it's free?</summary>

Absolutely! The goal of Bulk-Calculate is **collective progress**, not profit. Pricing it based on its value would make it inaccessible to many. Think of it like air—free, but invaluable.
</details>

<details><summary>Do I need to know VBA or coding to use Bulk-Calculate?</summary>

No coding or VBA knowledge is required. Bulk-Calculate has a user-friendly interface that allows you to map input and output cells using the Map Data Wizard and perform bulk calculations with just a few clicks.
</details>

<details><summary>Can I use my own Excel file with Bulk-Calculate?</summary>

Yes, you can use any standalone Excel file (without external references) as a template in Bulk-Calculate. Simply map your input and output cells using the Map Data Wizard.
</details>

<details><summary>What happens if my Excel template changes?</summary>

If your template changes, you can easily update the mappings by re-running the Map Data Wizard. Bulk-Calculate will adapt to the new structure and ensure all calculations are performed correctly.
</details>

<details><summary>Can Bulk-Calculate handle large datasets?</summary>

Yes, Bulk-Calculate processes data sequentially, one datapoint at a time. While it can handle large datasets, calculations and exports may take a bit longer for larger volumes of data.
</details>

<details><summary>Can Bulk-Calculate import data from files with different structures?</summary>

No, Bulk-Calculate requires that all files used for bulk import have the same structure. The input and output cells need to be mapped consistently across all files for successful data import.
</details>

<details><summary>Does Bulk-Calculate work with older versions of Excel?</summary>

Bulk-Calculate is compatible with Excel 2013 and later. However, for the best experience and performance, it's recommended to use the latest version of Excel.
</details>

<details><summary>Can I use Bulk-Calculate for non-engineering projects?</summary>

Yes! While Bulk-Calculate is designed for engineering calculations, it can be applied to any repetitive data-driven task. As long as you can map the input and output cells, it will work for your needs, whether in finance, research, or other fields.
</details>

<details><summary>Can I suggest improvements to Bulk-Calculate?</summary>

Yes, feel free to share ideas for improvement by using the Feedback button in the Bulk-Calculate ribbon menu. The maintainers will review your suggestion and prioritize it accordingly. You can also fast-track development by sponsoring the project <a href="https://github.com/sponsors/engineered-in" target="_blank">here</a>.
</details>

<details><summary>Can you create similar tools for me?</summary>

For custom development requests, please reach out through our <a href="https://www.linkedin.com/company/engineeredin/" target="_blank">LinkedIn Page</a>. Avoid using the Feedback button for these inquiries.
</details>

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

Special thanks to all [contributors](https://github.com/engineered-in/Bulk-Calculate/graphs/contributors) and users for their feedback and support.
