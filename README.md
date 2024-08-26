<div align="center">
  
  <h1>Moon's Tool</h1>
</div >
This project is a GUI application developed in Python that generates Excel reports based on CSV data inputs for ease of use.

<div align="center" style="max-width:20px;">
  <br>
<img src="https://github.com/user-attachments/assets/16734d8e-22bd-475d-a5b5-3e0d7692876b"  width="800px" />
</div>

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Libraries and Frameworks](#libraries-and-frameworks)
- [Main Functions](#main-functions)
- [Break Schedule](#break-schedule)
- [Additional Notes](#additional-notes)
- [Contributing](#contributing)
- [License](#license)

## Installation

1. **Clone the repository:**
    ```sh
    git clone https://github.com/your-repo/gui-application.git
    cd gui-application
    ```

2. **Create a virtual environment:**
    ```sh
    python -m venv venv
    ```

3. **Activate the virtual environment:**
    - On Windows:
        ```sh
        venv\Scripts\activate
        ```
    - On macOS/Linux:
        ```sh
        source venv/bin/activate
        ```

4. **Install the required dependencies:**
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. **Create a folder** for the project.
2. **Copy `Generate-Reports.exe`** to this folder.
3. **Export CSVs** from CZ for CSAT and from ZOHO for productivity:
    - For daily productivity, name the file `L2 Intraday.csv`.
    - For hourly productivity, name the file `ghassan.csv`.
    - For CSAT, ensure the file name starts with `IVR`.

4. **Copy the exported CSVs** to the folder.
5. **Run `Generate-Reports.exe`** and follow the on-screen instructions.

### Break Schedule

To create a break schedule automatically, follow the same steps to set up your folder and run the executable. 

## Libraries and Frameworks

- **Pillow (PIL):** 
    - Used for opening, manipulating, and saving image files.
    - [Documentation](https://pillow.readthedocs.io/en/stable/)

- **pandas:** 
    - Used for data manipulation and analysis.
    - [Documentation](https://pandas.pydata.org/)

- **openpyxl:** 
    - Used for reading and writing Excel 2010 xlsx/xlsm/xltx/xltm files.
    - [Documentation](https://openpyxl.readthedocs.io/en/stable/)

- **tkinter:** 
    - Used for creating graphical user interfaces.
    - [Documentation](https://docs.python.org/3/library/tkinter.html)

## Main Functions

### `C-SAT` Report

**Functionality**:
- Generates an Excel report with conditional formatting for CSAT data.

**Logic**:
1. **Read Data**: Reads CSAT data from a CSV file starting with "IVR".
2. **Process Data**: Calculates CSAT scores and formats them.
3. **Export Data**: Exports the processed data to an Excel file with conditional formatting.

**Usage**:
- Ensure the CSAT CSV file starts with "IVR".
- Place the CSV file in the same directory as `Generate-Reports.exe`.
- Run the executable and follow the instructions.

### `Productivity` Report

#### Daily Productivity

**Functionality**:
- Generates an Excel report for daily productivity data.

**Logic**:
1. **Read Data**: Reads productivity data from a CSV file named "L2 Intraday".
2. **Process Data**: Calculates daily productivity metrics.
3. **Export Data**: Exports the processed data to an Excel file.

**Usage**:
- Ensure the CSV file is named "L2 Intraday.csv".
- Place the CSV file in the same directory as `Generate-Reports.exe`.
- Run the executable and follow the instructions.

#### Hourly Productivity

**Functionality**:
- Generates an Excel report for hourly productivity data.

**Logic**:
1. **Read Data**: Reads productivity data from a CSV file named "ghassan".
2. **Process Data**: Calculates hourly productivity metrics.
3. **Export Data**: Exports the processed data to an Excel file.

**Usage**:
- Ensure the CSV file is named "ghassan.csv".
- Place the CSV file in the same directory as `Generate-Reports.exe`.
- Run the executable and follow the instructions.

## Additional Notes

- Ensure that all dependencies are installed before running the application.
- If you encounter any issues, please refer to the official documentation of the respective libraries.

## Contributing

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes and commit them (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Create a new Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
