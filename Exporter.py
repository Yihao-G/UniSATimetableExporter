import csv
import datetime
import re

import easygui
import requests
from bs4 import BeautifulSoup


def prompt_user_info():
    """getting username and password"""
    username_typed_in = False
    while not username_typed_in:
        username = easygui.enterbox("Your UniSA username: ", "Login")
        if username is None:
            if easygui.ynbox("Do you want to exit? ", "Exit"):
                exit(0)
        elif not len(username) == 8:
            easygui.msgbox("Your username should be eight digits. Please retype. ")
        else:
            username_typed_in = True

    password_typed_in = False
    while not password_typed_in:
        password = easygui.passwordbox("Your UniSA password: ", "Login")
        if password is None:
            if easygui.ynbox("Do you want to exit? ", "Exit"):
                exit(0)
        elif len(password) == 0:
            easygui.msgbox("The password should not be empty. Please retype. ")
        else:
            password_typed_in = True

    return {"username": username, "password": password}


def select_study_period(study_periods):
    """Prompt user to select the study period, and return the name of that study period"""
    selected_sp = None
    while selected_sp is None:
        selected_sp = easygui.choicebox("Please select the study period you want to export: ",
                                        "Select Study Period",
                                        study_periods)
        if selected_sp is None:
            if easygui.ynbox("Do you want to exit? ", "Exit"):
                exit(0)
    return selected_sp


easygui.msgbox("Welcome. \n"
               "This program is an open-source project on Github, developed by Yihao Gao. \n"
               "You can contact me by emailing to: gaoyy054@mymail.unisa.edu.au\n"
               "\n"
               "Click \"OK\" to continue. ")
easygui.msgbox("Please notice that we do not save your login details, and this program sends the login details to "
               "UniSA websites only. \n"
               "\n"
               "Please also notice that we have not handled the exceptions on internet problems yet. If the program "
               "crashed caused by internet problems, please rerun the program manually. \n"
               "\n"
               "Click \"OK\" to continue. ")
# login starts
login_successfully = False
while not login_successfully:
    user_info = prompt_user_info()
    response = requests.get('https://my.unisa.edu.au/Student/myEnrolment/myEnrolment/EnrolmentSummary.aspx',
                            auth=requests.auth.HTTPBasicAuth(**user_info))
    if response.status_code != 200:
        easygui.msgbox("Login failed. Please check your username or password is correct. ")
    else:
        login_successfully = True
# login ends

bs = BeautifulSoup(response.content, "html.parser")
study_periods_select_element = bs.find("li", {"class": "studyperiod"}).find("select")
study_periods_select_element_post_name = study_periods_select_element["name"]
study_periods_option_elements = study_periods_select_element.find_all("option")
study_periods_text = []
for sp in study_periods_option_elements:
    study_periods_text.append(sp.text)
# getting asp.net form general essential data starts
aspnetForm_data = bs.find(id="aspnetForm").find_all("input", type="hidden")
aspnetForm_data_dict = {}
# convert aspnetForm plain html data to dict object
for afd in aspnetForm_data:
    if afd.has_attr("value"):
        aspnetForm_data_dict[afd["name"]] = afd["value"]
# getting asp.net form general essential data ends

# select study period and get course table starts
no_course = True
while no_course:
    selected_study_period_text = select_study_period(study_periods_text)
    selected_study_period_tag = study_periods_option_elements[study_periods_text.index(selected_study_period_text)]
    selected_study_period_tag_value = selected_study_period_tag["value"]

    aspnetForm_data_dict[study_periods_select_element_post_name] = selected_study_period_tag_value

    courses = BeautifulSoup(requests
                            .post("https://my.unisa.edu.au/Student/myEnrolment/myEnrolment/EnrolmentSummary.aspx",
                                  data=aspnetForm_data_dict,
                                  auth=requests.auth.HTTPBasicAuth(**user_info)).content,
                            "html.parser").find_all(class_="DataTable")[1]
    if courses.find(class_="DataTableEmptyRow") is not None:
        easygui.msgbox("You have no courses enrolled in " + selected_study_period_text + ". Please reselect. ")
    else:
        no_course = False
# select study period and get course table ends

easygui.msgbox("Your timetable is ready to fetch. Please click \"OK\" to start. \n"
               "This will take a while. ")

# get all links of courses info page and save to courses_links variable
courses_links = []
for ct in courses.find_all("a", href=True):
    courses_links.append("https://my.unisa.edu.au/Student/myEnrolment/myEnrolment/" + ct["href"])

# all courses' info will be stored here
courses_time = []
# get into each course info page to fetch useful info
for course_link in courses_links:
    course_page = requests.get(course_link,
                               auth=requests.auth.HTTPBasicAuth(**user_info))
    course_page_bs = BeautifulSoup(course_page.content, "html.parser")

    course_info_block = course_page_bs.find(class_="EditableContent").nextSibling

    course_name_full = course_info_block.find("h2").text
    course_name_start = course_name_full.find(" - ") + 3
    course_name = course_name_full[course_name_start:]

    entries = course_info_block.find_all(class_="DEARow")
    course_type = entries[0].div.nextSibling.text
    instructor = re.sub(" +", " ", entries[7].div.nextSibling.text).strip()

    course_page_schedule = course_page_bs.find(class_="DataTable")

    # get date and time
    cp_trs = course_page_schedule.find_all("tr", class_=False)
    for tr in cp_trs:
        cp_tds = tr.find_all("td")

        location = cp_tds[0].text.strip()
        date_period = cp_tds[1].text
        start_date_end = date_period.find(" - ")
        day = cp_tds[2].text
        start_date = datetime.datetime.strptime(day + " " + date_period[:start_date_end], "%A %d %b %Y")
        end_date = datetime.datetime.strptime(day + " " + date_period[start_date_end + 3:], "%A %d %b %Y")
        time_period = cp_tds[3].text
        start_time_end = time_period.find("-")
        start_time = datetime.datetime.strptime(time_period[:start_time_end], "%I:%M %p").time()
        end_time = datetime.datetime.strptime(time_period[start_time_end + 1:], "%I:%M %p").time()

        # calculating all dates when have lessons
        this_course_all_dates = [start_date]
        while end_date not in this_course_all_dates:
            this_course_all_dates.append(this_course_all_dates[len(this_course_all_dates) - 1]
                                         + datetime.timedelta(days=7))
        # organise all info and put them into courses_time list ready for csv export
        for course_date in this_course_all_dates:
            courses_time.append({
                "Subject": course_type + " - " + course_name,
                "Start Date": course_date.strftime("%m/%d/%Y"),
                "Start Time": start_time.strftime("%I:%M %p"),
                "End Date": course_date.strftime("%m/%d/%Y"),
                "End Time": end_time.strftime("%I:%M %p"),
                "Description": instructor,
                "Location": location
            })

easygui.msgbox("Completed fetching your timetable. \n"
               "Click \"OK\" to choose where you want to save the .csv file to. ")

# prompt for saving location
save_path_selected = False
while not save_path_selected:
    save_file_to = easygui.filesavebox("Choose where you want to export your timetable to...",
                                       default="Timetable of " + selected_study_period_text + ".csv")
    if save_file_to is None:
        if easygui.ynbox("Do you want to exit without saving? ", "Exit"):
            exit(0)
    else:
        save_path_selected = True

# exporting csv file starts
with open(save_file_to, "w+") as csv_file:
    CSV_HEADER = ["Subject",
                  "Start Date", "Start Time", "End Date", "End Time",
                  "Description", "Location"]
    writer = csv.DictWriter(csv_file, delimiter=',', quoting=csv.QUOTE_ALL, fieldnames=CSV_HEADER, lineterminator="\n")
    writer.writeheader()
    for line in courses_time:
        writer.writerows(courses_time)
# exporting csv file ends

easygui.msgbox("Your timetable has been successfully exported to " + save_file_to + ". \n"
               "\n"
               "Please go to Google Calendar or other calendar providers to import the .csv file. ")
