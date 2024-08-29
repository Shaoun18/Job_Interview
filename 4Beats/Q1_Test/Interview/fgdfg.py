import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from datetime import datetime


def find_longest_shortest_options(driver, keyword):
    print(f"Attempting to search for keyword: {keyword}")
    try:
        search_box = driver.find_element(By.NAME, "q")
        print("Search box found, entering keyword...")
        search_box.clear()
        search_box.send_keys(keyword)
        search_box.send_keys(Keys.RETURN)
        print("Keyword entered, waiting for suggestions...")

        driver.implicitly_wait(5)
        suggestions = driver.find_elements(By.XPATH, '//li[@role="presentation"]//span')
        print(f"Found {len(suggestions)} suggestions.")

        if not suggestions:
            print(f"No suggestions found for keyword: {keyword}")
            return None, None

        options = [suggestion.text for suggestion in suggestions]
        longest_option = max(options, key=len)
        shortest_option = min(options, key=len)
        print(f"Longest option: {longest_option}, Shortest option: {shortest_option}")
        return longest_option, shortest_option

    except Exception as e:
        print(f"Error during search for keyword '{keyword}': {e}")
        return None, None


def main():
    excel_file_path = r'E:\Python\Interview\4BeatsQ1.xlsx'

    print("Loading Excel workbook...")
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        print(f"Excel workbook loaded from: {excel_file_path}")
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    day_of_week = datetime.now().strftime('%A')
    print(f"Today is: {day_of_week}")

    try:
        sheet = workbook[day_of_week]
        print(f"Working on sheet: {day_of_week}")
    except KeyError:
        print(f"No sheet found for {day_of_week}")
        return

    print("Setting up WebDriver...")
    try:
        service = Service(r'C:\Users\User\.wdm\drivers\chromedriver\win64\128.0.6613.86\chromedriver-win32')
        driver = webdriver.Chrome(service=service)
        driver.get("https://www.google.com")
        print("WebDriver set up and Google loaded.")
    except Exception as e:
        print(f"Error setting up WebDriver: {e}")
        return

    for row in sheet.iter_rows(min_row=2, max_col=1):
        keyword = row[0].value
        if keyword:
            print(f"Searching for keyword: {keyword}")
            longest, shortest = find_longest_shortest_options(driver, keyword)
            if longest and shortest:
                sheet.cell(row=row[0].row, column=2).value = longest
                sheet.cell(row=row[0].row, column=3).value = shortest
                print(f"Longest: {longest}, Shortest: {shortest}")

    print("Saving the workbook...")
    try:
        workbook.save(excel_file_path)
        print(f"Workbook saved successfully at {excel_file_path}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

    driver.quit()
    print("Browser closed.")


if __name__ == "__main__":
    main()
