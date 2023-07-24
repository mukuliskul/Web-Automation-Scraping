# Web Scraping/Automation

## Description:
This repository contains Python code to perform web scraping and automation from Google Scholar using Selenium. The script utilizes Selenium along with Chromedriver to access Google Scholar and extract information. Additionally, the xlwings and docx libraries are employed to read input queries and generate the desired output.

## Skills:
- Web Scraping
- Automation
- Python
- Selenium
- Chromedriver
- xlwings
- docx

## Learning Outcomes:
By using this repository and engaging with the code, you can expect to achieve the following learning outcomes:

1. **Proficiency in Web Scraping Techniques:** You will gain hands-on experience in web scraping methodologies, understanding how to interact with websites programmatically and extract relevant data.

2. **Automation Skills:** You will learn how to automate repetitive tasks using Python and Selenium, enabling you to save time and effort in data collection and processing.

3. **Familiarity with Selenium and Chromedriver:** You will become proficient in using Selenium along with the Chromedriver executable to control the Chrome browser for web scraping tasks.

4. **Working with Excel and Word Documents:** You will learn how to use the xlwings and docx libraries to interact with Excel files for input data and create Word documents to store the extracted information.

5. **Problem-Solving and Error Handling:** Throughout the code, you will encounter various problem-solving techniques and error handling approaches that are essential for robust and reliable automation.

6. **Understanding Data Pagination and Extraction:** As Google Scholar may have multiple pages of search results, you will learn how to navigate through these pages and extract data efficiently.

7. **Application of Python Libraries:** You will see practical applications of popular Python libraries like Selenium, xlwings, and docx, broadening your knowledge of their capabilities.

8. **Project Organization and Version Control:** By working with this repository, you will get familiar with organizing projects, using Git for version control, and collaborating with others on GitHub.

## Instructions:

### Prerequisites:
1. Make sure you have Python installed on your system.
2. Install the required libraries using the following command:
   ```
   pip install selenium xlwings docx
   ```

### Getting Started:
1. Clone this repository to your local machine using the following command:
   ```
   git clone https://github.com/your_username/your_repository.git
   ```

2. Navigate to the cloned repository:
   ```
   cd your_repository
   ```

3. Download the appropriate version of Chromedriver and place it in the project directory.

### Usage:
1. Prepare an Excel file with the list of titles of publications you want to search for. Save the file as 'exampleExcelOfPublications.xlsx' in the project directory.

2. Run the Python script 'web_scraping_google_scholar.py':
   ```
   python web_scraping_google_scholar.py
   ```

3. The script will prompt you to enter the FILE NUMBER. Enter the number corresponding to the row in the Excel file where you want to start scraping.

4. The script will open a Chrome browser and navigate to Google Scholar. It will search for the title from the Excel file and start extracting citation details.

5. During the process, you will be prompted to enter an Operation. If you enter 'n' or 'N', the script will skip processing that title and move to the next one.

6. Once the citations are extracted, a Word document will be created for each title, containing the title itself and the corresponding citation details. The document will be saved in the project directory.

7. The script will continue to scrape the next titles from the Excel file until it reaches the end.

### Note:
- The 'chromedriver' path in the code should be updated to match the location where you placed the Chromedriver executable on your system.
- Ensure that you have a stable internet connection during the scraping process.

Feel free to use, modify, and share this code as per your requirements! Happy scraping!