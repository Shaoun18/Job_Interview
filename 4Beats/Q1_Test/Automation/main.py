import openpyxl
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

options = Options()
# options.add_experimental_option('detach',True)


def find_longest_shortest_options(driver, keyword):
    # Find the search box and input the keyword
    search_box = driver.find_element_by_name('q')
    search_box.clear()
    search_box.send_keys(keyword)
    driver.implicitly_wait(2)  # Wait for suggestions to appear
    suggestions = driver.find_elements_by_xpath('//li[@role="presentation"]//span')

    if not suggestions:
        return None, None

    options = [suggestion.text for suggestion in suggestions]
    longest_option = max(options, key=len)
    shortest_option = min(options, key=len)
    return longest_option, shortest_option


def main():
    # Path to your Excel file
    excel_file_path = 'E:/Python/Assignment/4BeatsQ1.xlsx'

    # Load the Excel workbook
    workbook = openpyxl.load_workbook(excel_file_path)

    # Determine the current day of the week
    day_of_week = datetime.now().strftime('%A')
    sheet = workbook[day_of_week]  # Access the sheet for the current day

    # Set up the WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://www.google.com/")

    for row in sheet.iter_rows(min_row=2, max_col=1):
        keyword = row[0].value
        if keyword:
            longest, shortest = find_longest_shortest_options(driver, keyword)
            # Write the longest and shortest options back to the Excel file
            if longest and shortest:
                sheet.cell(row=row[0].row, column=2).value = longest
                sheet.cell(row=row[0].row, column=3).value = shortest

    # Save the workbook
    workbook.save(excel_file_path)

    # Close the browser
    driver.quit()


if __name__ == "__main__":
    main()


# import openpyxl
# from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
# from datetime import datetime
#
#
# def find_longest_shortest_options(driver, keyword):
#     # Find the search box and input the keyword
#     search_box = driver.find_element("name", "q")
#     search_box.clear()
#     search_box.send_keys(keyword)
#     search_box.send_keys(Keys.RETURN)  # Use RETURN to submit the search
#
#     driver.implicitly_wait(2)  # Wait for suggestions to appear
#     suggestions = driver.find_elements("xpath", '//li[@role="presentation"]//span')
#
#     if not suggestions:
#         return None, None
#
#     options = [suggestion.text for suggestion in suggestions]
#     longest_option = max(options, key=len)
#     shortest_option = min(options, key=len)
#     return longest_option, shortest_option
#
#
# def main():
#     # Path to your Excel file
#     excel_file_path = 'E:/Python/Assignment/4BeatsQ1.xlsx'
#
#     # Load the Excel workbook
#     try:
#         workbook = openpyxl.load_workbook(excel_file_path)
#     except Exception as e:
#         print(f"Error loading Excel file: {e}")
#         return
#
#     # Determine the current day of the week
#     day_of_week = datetime.now().strftime('%A')
#     try:
#         sheet = workbook[day_of_week]  # Access the sheet for the current day
#     except KeyError:
#         print(f"No sheet found for {day_of_week}")
#         return
#
#     # Set up the WebDriver
#     driver = webdriver.Chrome('C:/Users/User/.wdm/drivers/chromedriver/win64/128.0.6613.86/chromedriver-win32')
#     driver.get("https://www.google.com")
#
#     for row in sheet.iter_rows(min_row=2, max_col=1):
#         keyword = row[0].value
#         if keyword:
#             longest, shortest = find_longest_shortest_options(driver, keyword)
#             # Write the longest and shortest options back to the Excel file
#             if longest and shortest:
#                 sheet.cell(row=row[0].row, column=2).value = longest
#                 sheet.cell(row=row[0].row, column=3).value = shortest
#
#     # Save the workbook
#     try:
#         workbook.save(excel_file_path)
#     except Exception as e:
#         print(f"Error saving Excel file: {e}")
#
#     # Close the browser
#     driver.quit()
#
#
# if __name__ == "__main__":
#     main()



