import os
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
import selenium.webdriver.support.expected_conditions as ec
from colorama import init, Fore, Style
from PyQt6 import QtWidgets
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import QMessageBox, QTableWidgetItem
from ui.UploadScrapper import Ui_UploadScrapper


class MainApp(QtWidgets.QMainWindow):
    form_fields_value = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        self.scrapping_worker = None
        self.ui = Ui_UploadScrapper()
        self.ui.setupUi(self)
        self.initializer()
        self.ui.btn_scrape.clicked.connect(self.evt_scrapping_process)
        self.show()

    # Function to initialize UI components
    def initializer(self):
        # Populate the web driver drop down box with the list below
        webdriver_list = ['Chrome web driver', 'Firefox web driver', 'Opera web driver', 'Edge web driver']
        self.ui.cbo_webdriver.addItems(webdriver_list)
        self.ui.cbo_webdriver.setCurrentIndex(0)            # Remove in production
        # Set the tool tips for the UI components
        self.ui.txt_email.setToolTip('Your NDR login email.')
        self.ui.txt_password.setToolTip('Your NDR login password.')
        self.ui.txt_search.setToolTip('Search by either DATIM code OR the name of facility. This will narrow your request.')
        # Config the column width of the table
        self.ui.tbw_processed.setColumnWidth(0, 300)
        self.ui.tbw_processed.setColumnWidth(1, 350)
        self.ui.tbw_processed.setColumnWidth(2, 180)
        self.ui.tbw_processed.setColumnWidth(3, 200)
        self.ui.tbw_processed.setColumnWidth(4, 450)
        self.ui.tbw_processed.setColumnWidth(5, 100)
        self.ui.tbw_processed.setColumnWidth(6, 100)
        self.ui.tbw_processed.setColumnWidth(7, 100)
        self.ui.tbw_processed.setColumnWidth(8, 100)

    # Method to give user user_feedback through message box. The title and icon parameters are option;
    # if not provide the default values are used.
    @staticmethod
    def user_feedback(feedback, title='Info', icon=QMessageBox.Icon.Information):
        mbox = QMessageBox()
        mbox.setText(feedback)
        mbox.setWindowTitle(title)
        mbox.setIcon(icon)
        mbox.exec()

    # Start scrapping the upload page of NDR, but ensure all required parameters are entered by the user before starting.
    def evt_scrapping_process(self):
        self.ui.tbw_processed.clearContents()                                       # Clear all rows
        self.ui.tbw_processed.setRowCount(0)                                        # Reset row counter
        self.ui.pgb_processing.setValue(0)                                          # Clear progress bar
        self.scrapping_worker = ScrappingWorkerThread()
        self.scrapping_worker.start()

        self.form_fields_value.connect(self.scrapping_worker.form_values)
        form_data = {'driver': self.ui.cbo_webdriver.currentText(),
                     'url': self.ui.txt_url.text().strip(),
                     'email': self.ui.txt_email.text().strip(),
                     'pwd': self.ui.txt_password.text().strip(),
                     'search': self.ui.txt_search.text().strip(),
                     'max_page': self.ui.spb_number_of_page_scrape.value(),
                     'sleep': self.ui.spb_sleep_time.value()
                     }
        self.form_fields_value.emit(form_data)

        # Hook signals from the worker thread to the main thread
        self.scrapping_worker.update_progress.connect(self.evt_update_progress)
        self.scrapping_worker.user_feedback.connect(self.evt_feedbacks)
        self.scrapping_worker.processed_data.connect(self.evt_update_table_progress)
        self.scrapping_worker.enable_scrape_button.connect(self.evt_enable_scrape_button)

    def evt_feedbacks(self, feedback):
        if feedback['message_type'] == 'success':
            self.user_feedback(feedback['message'], feedback['title'], QMessageBox.Icon.Information)
        if feedback['message_type'] == 'required':
            self.user_feedback(feedback['message'], feedback['title'], QMessageBox.Icon.Warning)
        if feedback['message_type'] == 'error':
            self.user_feedback(feedback['message'], feedback['title'], QMessageBox.Icon.Critical)

    def evt_update_progress(self, value):
        self.ui.pgb_processing.setValue(value)

    def evt_update_table_progress(self, data):
        current_row_count = self.ui.tbw_processed.rowCount()                       # necessary even when there are no rows in the table
        self.ui.tbw_processed.insertRow(current_row_count)
        self.ui.tbw_processed.setItem(current_row_count, 0, QTableWidgetItem(data['username']))
        self.ui.tbw_processed.setItem(current_row_count, 1, QTableWidgetItem(data['facility']))
        self.ui.tbw_processed.setItem(current_row_count, 2, QTableWidgetItem(data['upload_date']))
        self.ui.tbw_processed.setItem(current_row_count, 3, QTableWidgetItem(data['batch']))
        self.ui.tbw_processed.setItem(current_row_count, 4, QTableWidgetItem(data['zip_file']))
        self.ui.tbw_processed.setItem(current_row_count, 5, QTableWidgetItem(data['total']))
        self.ui.tbw_processed.setItem(current_row_count, 6, QTableWidgetItem(data['fails']))
        self.ui.tbw_processed.setItem(current_row_count, 7, QTableWidgetItem(data['passes']))
        self.ui.tbw_processed.setItem(current_row_count, 8, QTableWidgetItem(data['pending']))

    def evt_enable_scrape_button(self, status):
        self.ui.btn_scrape.setEnabled(status)


class ScrappingWorkerThread(QThread):
    user_feedback = pyqtSignal(dict)
    required_feedback = pyqtSignal(str)
    update_progress = pyqtSignal(int)
    processed_data = pyqtSignal(dict)
    enable_scrape_button = pyqtSignal(bool)
    input_data = {}                 # Dictionary initialization

    @pyqtSlot(dict)
    def form_values(self, data):
        self.input_data = data                                 # Load the data into the empty dictionary
        print(Style.BRIGHT + Fore.LIGHTYELLOW_EX + 'Calling from Thread.\n' + str(self.input_data))

    def run(self):
        if self.input_data['driver'] == '':
            self.user_feedback.emit({'message': 'Please specify the web driver to use for this scrapping process.', 'title': 'Web driver required', 'message_type': 'required'})
            return
        if self.input_data['url'] == '':
            self.user_feedback.emit({'message': 'Provide the target URL for the scrapping.', 'title': 'URL required','message_type': 'required'})
            return
        if self.input_data['email'] == '':
            self.user_feedback.emit({'message': 'Provide sign on email address.', 'title': 'Email required','message_type': 'required'})
            return
        if self.input_data['pwd'] == '':
            self.user_feedback.emit({'message': 'Provide sign on password.', 'title': 'Password required','message_type': 'required'})
            return

        init(autoreset=True)                                                    # Initializes Colorama
        start_time = time.time()                                                # Record the time the processing started
        move_to_next_page = 1
        max_page_to_scrape = self.input_data['max_page']                        # Set default number of page(s) to iterate
        relax_seconds = self.input_data['sleep']                                # Set a default delay of in seconds if the user did not supply any value
        timeout_duration = 90                                                   # Set default page timeout to in seconds
        search_query = self.input_data['search']                                # Ser default search string to blank
        ndr_driver = ''                                                         # Default URL is blank
        self.enable_scrape_button.emit(False)                                   # Send a signal to disable the scrape button when the processing start

        try:
            # Decide which web drive to use
            if 'chrome' in self.input_data['driver'].lower():
                ndr_driver = webdriver.Chrome()
            if 'firefox' in self.input_data['driver'].lower():
                ndr_driver = webdriver.Firefox()
            if 'opera' in self.input_data['driver'].lower():
                ndr_driver = webdriver.Opera()
            if 'edge' in self.input_data['driver'].lower():
                ndr_driver = webdriver.Edge()

            ndr_driver.get(self.input_data['url'])
        except Exception as e:
            # I am interested in showing the user the exception thrown, but I want to only show the message part and not the full
            # exception. To achieve this I had to call the __dic__ list and give it the msg key which is the dictionary I am interested in.
            self.user_feedback.emit({'message': e.__dict__['msg'] + '\n\nUpdate the web driver to a version that supports your current browser\'s version and try again.', 'title': 'Browser launch failed', 'message_type': 'error'})
            ndr_driver.close()
            self.enable_scrape_button.emit(True)                            # Send a signal to enable the scrape button
            return

        wait_timeout = WebDriverWait(ndr_driver, timeout_duration)
        # The wait_timeout above allow for the web page to at least load properly before sending the email and password values.
        ndr_driver.find_element(By.XPATH, '//*[@id="Input_Email"]').send_keys(self.input_data['email'])
        ndr_driver.find_element(By.XPATH, '//*[@id="Input_Password"]').send_keys(self.input_data['pwd'])
        ndr_driver.find_element(By.XPATH, '//*[@id="account"]/div[5]/button').send_keys('\n')

        # After login, set page zoom. If you are using zoom functionality note that click() method will not work, you have to use send_keys('\n') method for ENTER action.
        ndr_driver.execute_script("document.body.style.zoom='75%'")

        # # Load the option to show 100 (last dropdown option) entries per page. However, this is not yet functional as the process throws exception when looping many pages
        # # this is hoped to be perfected.
        # show_size = Select(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable_length"]/label/select'))
        # option_count = len(show_size.options)
        # show_size.select_by_index(option_count - 1)

        # wait_timeout.until(EC.presence_of_element_located((By.XPATH, '//*[@id="uploadDataTable_length"]/label/select/option[' + str(option_count - 1) + ']')))

        # Search string was supplied. This code has to be placed at this point if not the search criteria will not take effect on the first page.
        if len(search_query) != 0:
            wait_timeout.until(ec.presence_of_element_located((By.XPATH, '//*[@id="uploadDataTable_filter"]/label/input'))).send_keys(search_query)

        # This timer sleep is very important to allow the page to load at the initial
        time.sleep(relax_seconds)

        try:
            next_p = wait_timeout.until(ec.presence_of_element_located((By.XPATH, '//*[@id="uploadDataTable_next"]'))).get_attribute('data-dt-idx')
            # print('Next_Page_Button_ndex: ' + str(next_p))

            last_val = ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable_paginate"]/span/a[' + str(int(next_p) - 1) + ']').text
            # Ensure the user does not supply more than the existing page(s) to avoid app crashing.
            # Last page is determined by the values that is set for the Show_entries and Search filters controls on the NDR web page
            if max_page_to_scrape > int(last_val):
                max_page_to_scrape = int(last_val)

            print(Style.BRIGHT + Fore.LIGHTCYAN_EX + 'The processing will scrape ' + str(max_page_to_scrape) + ' page(s) from the NDR upload web page.')
        except TimeoutException as te:
            print(Style.BRIGHT + Fore.RED + "Loading took too much time!\n" + str(te))
            self.user_feedback.emit({'message': 'Loading took too much time!\nEnsure that the web driver selected has been installed on this PC and the path set in the environment variable.', 'title': 'Timeout', 'message_type': 'error'})
            ndr_driver.close()
            self.enable_scrape_button.emit(True)  # Send a signal to enable the scrape button
            return

        # Get the number of entries shown per page from the drop-down box
        entries_per_page = int(Select(ndr_driver.find_element(By.XPATH,'//*[@id="uploadDataTable_length"]/label/select')).first_selected_option.text)
        print(Style.BRIGHT + Fore.MAGENTA + 'Entries per page is set to ' + str(entries_per_page))

        # define column names
        username = []
        facility = []
        upload_date = []
        batch = []
        zip_file = []
        total = []
        fails = []
        passes = []
        pending = []

        print(Style.BRIGHT + Fore.BLUE + '\nGo grasp a cup of coffee while I run your errand. Cheers!')

        rows_of_records_processed_counter = 0
        while move_to_next_page <= int(max_page_to_scrape):
            # Do not click the Next button if it's the first page that is loaded
            try:
                if move_to_next_page != 1:
                    ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable_next"]').send_keys('\n')
                    # After moving to the next page, allow for the page to load completely
                    time.sleep(relax_seconds)
            except Exception as e:
                print(Style.BRIGHT + Fore.RED + 'Move to next page error:\n' + str(e))
                self.user_feedback.emit({'message': 'Move to next page error!\n' + e.__dict__['msg'], 'title': 'Timeout', 'message_type': 'error'})
                ndr_driver.close()
                self.enable_scrape_button.emit(True)  # Send a signal to enable the scrape button
                return

            for row in range(1, entries_per_page + 1):
                try:
                    # # Just checking if the records on current row are load. However, I only check the username field, but if it loads then it is assumed that other loaded as well
                    # print('Page-' + str(move_to_next_page) + ' :: User name displayed @ row ' + str(row) + ' :' + str(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[1]/div[2]'.format(str(row))).is_displayed()))

                    username.append(ndr_driver.find_element(By.XPATH,'//*[@id="uploadDataTable"]/tbody/tr[{}]/td[1]/div[2]'.format(str(row))).text.strip())
                    facility.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[2]/div[3]/b'.format(str(row))).text.strip())
                    upload_date.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[1]/div[3]'.format(str(row))).text.strip())
                    batch.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[2]/div[1]'.format(str(row))).text.strip())
                    zip_file.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[2]/div[2]/span[2]'.format(str(row))).text.strip())
                    total.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[4]'.format(str(row))).text.strip())
                    fails.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[5]'.format(str(row))).text.strip())
                    passes.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[6]'.format(str(row))).text.strip())
                    pending.append(ndr_driver.find_element(By.XPATH, '//*[@id="uploadDataTable"]/tbody/tr[{}]/td[7]'.format(str(row))).text.strip())

                    # Send out the processed row of data to be used to update the table widget on the UI.
                    # Because the variables are Array object, the current record is accessed using the [row - 1] index counter as array index start from 0 to n-1.
                    # The for loop start at index 1 which if used that way will flag an out of bound index.
                    row_data = {'username': username[row - 1], 'facility': facility[row - 1], 'upload_date': upload_date[row - 1], 'batch': batch[row - 1], 'zip_file': zip_file[row - 1], 'total': total[row - 1], 'fails': fails[row - 1], 'passes': passes[row - 1], 'pending': pending[row - 1]}
                    self.processed_data.emit(row_data)

                    # I am monitoring the progress of how rows of the total record entries has been process time 100% and use it to update the value of the progress bar for the user to see
                    rows_of_records_processed_counter +=1
                    progress_counter = (int((rows_of_records_processed_counter / (entries_per_page * max_page_to_scrape)) * 100))
                    self.update_progress.emit(progress_counter)
                    # print('progress_counter: ' + str(progress_counter) + '% ::rows_of_records_processed_counter: ' + str(rows_of_records_processed_counter) + ' @ move_to_next_page: ' + str(move_to_next_page))
                except StaleElementReferenceException as e:
                    print(Style.BRIGHT + Fore.RED + "\nThe scrapping was not successful.\nTry running the application again with a high page load waiting time, this will help if you have\n many pages to scrape or your internet connection is of poor quality.\n\n" + str(e))
                    msg = '\nThe scrapping was not successful.\nTry running the application again with a high page load waiting time, this will help if you have many pages to scrape or your internet connection is of poor quality.\n\nError message:\n' + str(e.__dict__['msg'])
                    self.user_feedback.emit({'message': msg, 'title': 'Scrapping failed', 'message_type': 'error'})
                    ndr_driver.close()
                    self.enable_scrape_button.emit(True)                    # Send a signal to enable the scrape button
                    return

            # Increment counter to move to the next page
            move_to_next_page = move_to_next_page + 1

        datafile = pd.DataFrame({'Login user': username, 'Facility': facility, 'Upload date': upload_date, 'Batch No.': batch, 'Zip file': zip_file, 'Total': total, 'Fails': fails, 'Passes': passes, 'Pending': pending})

        # Define the output path/file
        user_path = os.path.expanduser("~")
        output_file = str(user_path) + '\\Downloads\\ndr_upload_tracker.xlsx'

        # Delete file if it does exist on the location specified.
        if os.path.exists(output_file):
            os.remove(output_file)

        # Write scrapped data to disk. It is very important to install the openpyxl module for the export to .xlsx file type.
        datafile.to_excel(output_file, sheet_name='NDR upload tracker')
        ndr_driver.close()
        self.enable_scrape_button.emit(True)              # Send a signal to enable the scrape button when the processing finish

        # Record the time processing finished and compute duration
        elapse_time = time.time() - start_time
        hrs = elapse_time // 3600
        mints = (elapse_time - 3600) // 60 if elapse_time > 3600 else elapse_time // 60
        secs = elapse_time % 60
        print(Style.BRIGHT + Fore.GREEN + '\nElapse_time: ' + str(elapse_time))

        process_feedback = 'A total row of ' + str(len(username)) + ' datasets were scrapped and saved to a file located at ' + output_file
        process_feedback += '\n\nProcessed time took: {:0>2d}'.format(int(hrs)) + ' hr : ' + '{:0>2d}'.format(int(mints)) + ' min : ' + '{:0>2d}'.format(int(secs)) + ' sec'
        print(Style.BRIGHT + Fore.GREEN + '\n' + process_feedback + '\n')
        self.user_feedback.emit({'message': process_feedback, 'title': 'Scrapping successful', 'message_type': 'success'})


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_app = MainApp()
    # Load the style sheet file and apply it.
    with open('qss/theme.qss', 'r') as file:
        main_app.setStyleSheet(file.read())
        # app.setStyleSheet(file.read())
    sys.exit(app.exec())
