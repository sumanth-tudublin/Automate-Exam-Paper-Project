import os
import time
import getpass
import itertools
import urllib.parse
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import NoSuchElementException

year = input("Please enter which year you are in (1/2/3/4) : ")

while True:
    season = int(input('\nSearch in season (\n1 : Winter\n2 : Summer\n ) : '))
    if season == 1:
        season = 'Winter'
        break
    elif season == 2:
        season = 'Summer'
        break
    else:
        print("Invalid Number. Does not match any season\n")

print("\nCollecting Modules Information...")

# Settings for chrome driver.
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.headless = True

# Open chrome driver (below this statements are functions for the tabs functions).
driver = webdriver.Chrome(executable_path=r'C:\bin\chromedriver.exe', options=options)
driver.get("https://www.dit.ie/catalogue/Programmes/Details/DT228?tab=Programme%20Structure")

module_name_odd = driver.find_elements_by_xpath(
    "//table[@class='listing full_width']/tbody/tr[@class='section_detail row_odd "
    "dt228" + str(year) + " hidden']/td[2]/a")  # 6
module_name_even = driver.find_elements_by_xpath(
    "//table[@class='listing full_width']/tbody/tr[@class='section_detail row_even "
    "dt228" + str(year) + " hidden']/td[2]/a")  # 6
module_name = module_name_odd + module_name_even

module_id_odd = driver.find_elements_by_xpath(
    "//table[@class='listing full_width']/tbody/tr[@class='section_detail row_odd "
    "dt228" + str(year) + " hidden']/td[3]")
module_id_even = driver.find_elements_by_xpath(
    "//table[@class='listing full_width']/tbody/tr[@class='section_detail row_even "
    "dt228" + str(year) + " hidden']/td[3]")
module_id = module_id_odd + module_id_even

modules_pack = list(itertools.chain(*[(i, j) for i, j in zip(module_id, module_name)]))

modules_pack = [x.get_attribute("textContent") for x in modules_pack]

driver.quit()

modules_pack = {modules_pack[i]: modules_pack[i + 1] for i in range(0, len(modules_pack), 2)}

modules_pack = {k: v for k, v in sorted(modules_pack.items(), key=lambda i: i[1])}

print("\nDone Collecting Modules Information...")

print("\nDownloading all Files. Please Wait...\n")


# modules_pack = {
#     # course code   course name
#     'CMPU3042': 'Systems Security',
#     'CMPU3004': 'Applied Intelligence',
#     'CMPU3005': 'Business & Enterprise',
#     ...etc...
# }

def launch_browser():
    # Settings for chrome driver.
    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
                                    )
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.headless = True

    # Open chrome driver (below this statements are functions for the tabs functions).
    driver = webdriver.Chrome(executable_path=r'C:\bin\chromedriver.exe', options=options)
    print()
    # Open Exam Papers Website.
    driver.get("http://Student:ThunderRoad@student.dit.ie/exampapers/KT/")
    return driver


for module, subject in modules_pack.items():
    print("Working on subject: %s" % subject)

    # Directory to keep all the exam papers in.
    download_dir = "C:\\Users\\" + \
                   getpass.getuser() + "\\Downloads\\Exam_Paper1\\" + subject

    # Check if directory exists, or else make the directories.
    try:
        os.makedirs(download_dir)
    except OSError:
        print("Proceeding on existing file...")

    driver = launch_browser()


    def go_to_tab(tab_number):
        """Switch tabs by given tab number."""
        inner_tabs = driver.window_handles
        driver.switch_to.window(inner_tabs[tab_number])


    def close_tab():
        """Closes tab and switch tabs to given tab_number"""
        driver.close()
        inner_tabs = driver.window_handles
        driver.switch_to.window(inner_tabs[-1])


    def open_in_new_tab(tab_to_be_opened):
        """Opens given element in new tab."""
        ActionChains(driver).key_down(Keys.CONTROL).click(
            tab_to_be_opened).key_up(Keys.CONTROL).perform()


    # Store all years to search in.
    year1 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202013-2014%20Sci/"]')
    year2 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202014-2015%20Sci/"]')
    year3 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202015-2016%20Sci/"]')
    year4 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202016-2017%20Sci/"]')
    year5 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202017-2018%20Sci/"]')
    year6 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202018-2019%20Sci/"]')
    year7 = driver.find_element_by_xpath(
        '//a[@href="Academic%20Year%202019-2020%20Sci/"]')
    years = [year1, year2, year3, year4, year5, year6]

    # Search in all years.
    for year in years:
        # Get current year that we are searching in.
        academic_year = year.get_attribute("href")[74:83]


        def search_document(season_chosen):
            """Open season (winter/summer) and search for module to download."""
            print("Searching Document..")
            season_chosen = driver.find_element_by_partial_link_text(season_chosen)
            open_in_new_tab(season_chosen)
            go_to_tab(2)
            pdfs_save = driver.find_elements_by_partial_link_text(module.upper())
            return pdfs_save


        def download(pdf_save):
            """Download PDF and return file name (to rename later)."""
            print("Downloading PDF...")
            pdf_save.click()
            file_name = os.path.basename(pdf_save.get_attribute('href'))
            file_name = urllib.parse.unquote(file_name)
            return file_name


        def rename(download_file_name):
            """Waits for file to exist (finish downloading) and changes it's name to corresponding year."""
            print("Renaming Document...\n")
            while not os.path.exists(download_dir + "\\" + download_file_name):
                time.sleep(1)

            if os.path.isfile(download_dir + "\\" + download_file_name):
                old_file = os.path.join(download_dir, download_file_name)
                new_file_name = download_file_name.replace(module.upper(), academic_year)
                new_file = os.path.join(download_dir, new_file_name)
                os.rename(old_file, new_file)
            else:
                # Throws error if it isn't a file.
                raise ValueError("%s isn't a file!" % download_file_name)


        def open_up(season_chosen):
            """Executes all three functions (searches documents in web browser, download the file locally and renames
            it). """
            pdfs_save = search_document(season_chosen)
            for pdf_save in pdfs_save:
                download_file_name = download(pdf_save)
                rename(download_file_name)


        try:
            # Try in Winter
            print("Trying in %s: %s" % (season, academic_year))
            open_in_new_tab(year)
            # Wait for the tab to open (or else it will just execute next steps even when element are not there and
            # throws an error).
            WebDriverWait(driver, 60).until(
                expected_conditions.number_of_windows_to_be(2))
            go_to_tab(1)  # Once tab is open, go to the tab.
            open_up(season)
            close_tab()  # Close Exam Papers tab -> Now in Season tab.
        except NoSuchElementException:
            # Does not exist
            close_tab()  # Close Summer (we opened this in the try statement).
            print("Not In Year: %s\n\n" % academic_year)

        close_tab()  # Close Season tab -> Now in All Academic Years tab.

    # All are finished. Close Browser.
    driver.quit()
    print("\n\nFinished All Files")

# Open Folder to show user downloaded files.
os.startfile("C:\\Users\\" +
             getpass.getuser() + "\\Downloads\\Exam_Paper1\\")
