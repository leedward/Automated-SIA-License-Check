# Automated SIA License Check

## Prerequisites

- Python 3.9+
- Python Libraries: `openpyxl`, `selenium`
- Chrome browser
- [ChromeDriver](https://sites.google.com/chromium.org/driver/)

## Installation

### Windows

1. Make sure you have Python 3.8 or above installed. You can verify this by running `python --version` in your command prompt.

2. Clone this repository or download the `main.py` script.

3. Install the necessary Python libraries. You can do this by running the following command in your command prompt:

    ```shell
    pip install openpyxl selenium
    ```

4. You need to have the Chrome browser installed. The ChromeDriver version you'll use should match the version of your Chrome browser. To check your Chrome version:
   * Click on the three dots on the top right corner of the browser.
   * Go to `Help` -> `About Google Chrome`.
   * Your Chrome version will be displayed.

5. Download the corresponding ChromeDriver from the [ChromeDriver download page](https://chromedriver.chromium.org/downloads). Make sure to download the version that matches your Chrome version and system architecture.

6. Extract the `chromedriver.exe` from the downloaded zip file.

7. Add the path to `chromedriver.exe` to your System PATH.

### Linux and macOS

1. Make sure you have Python 3.8 or above installed. You can verify this by running `python3 --version` in your terminal.

2. Clone this repository or download the `main.py` script.

3. Install the necessary Python libraries. You can do this by running the following command in your terminal:

    ```shell
    pip3 install openpyxl selenium
    ```

4. You need to have the Chrome browser installed. The ChromeDriver version you'll use should match the version of your Chrome browser. To check your Chrome version:
   * Click on the three dots on the top right corner of the browser.
   * Go to `Help` -> `About Google Chrome`.
   * Your Chrome version will be displayed.

5. Download the corresponding ChromeDriver from the [ChromeDriver download page](https://chromedriver.chromium.org/downloads). Make sure to download the version that matches your Chrome version and system architecture.

6. Move the downloaded `chromedriver` file to `/usr/local/bin/chromedriver` using the following command:

    ```shell
    mv /path/to/downloaded/chromedriver /usr/local/bin/chromedriver
    ```

7. Give executable permissions to the `chromedriver` binary with `chmod +x /usr/local/bin/chromedriver`.

8. On macOS, you may also need to remove the quarantine attribute that the system applies to downloaded files. You can do this with the following command:
    ```shell
    xattr -d com.apple.quarantine /usr/local/bin/chromedriver
    ```

## Usage

The script requires two command line arguments: the path to the spreadsheet containing the license numbers and the path to the ChromeDriver executable. The license numbers should be in column A of the spreadsheet, starting from the second row (A2).

1. Open the command line in the folder where you have saved `main.py`.

2. Run the script with the paths to the spreadsheet and the ChromeDriver as command line arguments:

    ```shell
    python main.py -s /path/to/your/spreadsheet.xlsx -d /path/to/chromedriver
    ```

   Make sure to replace `/path/to/your/spreadsheet.xlsx` with the actual path to your spreadsheet and `/path/to/chromedriver` with the actual path to the ChromeDriver executable.

3. The script will take some time to run, depending on the number of license numbers in the spreadsheet. It will update the spreadsheet with the results of each license number check. The results will be stored in columns B-H of the spreadsheet.

## Troubleshooting

If the script fails to run or completes with errors, ensure the following:

1. The spreadsheet and the ChromeDriver paths are correct and both the files exist at the provided locations.

2. The ChromeDriver version matches the version of the Google Chrome installed on your system.

3. The Python version is 3.8 or above and all the required libraries are installed.

4. The website's structure has not changed. If it has, the script may need to be updated to correctly interact with the form and extract the results.

If you're getting an `OSError: [Errno 8] Exec format error`, it means that the ChromeDriver binary is not compatible with your system. Check your Chrome version and system architecture, and make sure you've downloaded the correct ChromeDriver version

## License

This project is licensed under the terms of the [MIT license](License.md).
