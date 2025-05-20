import datetime
import json
from os import system
import os
from os.path import split


class InputUtil:

    def __init__(self):
        self.project_keys_json_file_path = "Project_Keys.json"
        self.project_names_json_file_path = "Project_names.json"
        self.start_date = None
        self.end_date = None
        self.project_keys = []
        self.project_names = []
        self.input_Date()
        self.input_project()

    def get_project_names(self):
        return self.project_names

    def get_project_keys(self):
        return self.project_keys

    def get_startDate(self):
        return self.start_date

    def get_endDate(self):
        return self.end_date

    def get_current_week_date_range_file_name(self):
        sd = self.start_date
        sd_day = sd.strftime("%d")
        sd_month = sd.strftime("%b")
        sd_year = sd.strftime("%Y")
        sd_string = sd_day + " " + sd_month + " " + sd_year

        ed = self.end_date
        ed = self.start_date + datetime.timedelta(days=6)  # Subtract 1 day from end date
        ed_day = ed.strftime("%d")
        ed_month = ed.strftime("%b")
        ed_year = ed.strftime("%Y")
        ed_string = ed_day + " " + ed_month + " " + ed_year

        full_date_string = sd_string + " - " + ed_string
        return full_date_string

    def get_current_week_date_range_full(self):
        sd = self.start_date
        sd_day = sd.strftime("%d")
        sd_month = sd.strftime("%b")
        sd_year = sd.strftime("%Y")
        sd_string = sd_month + " " + sd_day + ", " + sd_year

        ed = self.end_date
        ed = self.start_date + datetime.timedelta(days=6) # Subtract 1 day from end date
        ed_day = ed.strftime("%d")
        ed_month = ed.strftime("%b")
        ed_year = ed.strftime("%Y")
        ed_string = ed_month + " " + ed_day + ", " + ed_year

        full_date_string = sd_string + " - " + ed_string
        return full_date_string

    def input_Date(self):
        # Input Start date
        start_date = None
        while start_date is None:  # Keep asking for input, until the correct date input is provided
            try:
                start_date_input = input("\tEnter the Start Date (DD/MM/YYYY) : ")
                start_date = datetime.datetime.strptime(start_date_input, "%d/%m/%Y")
            except ValueError as ve:
                print("\t\tError: ", ve)
                print("\t\tPlease enter the start date again")

        # Input End Date
        self.start_date = start_date
        self.end_date = start_date + datetime.timedelta(days=6)

    def input_project(self):

        project_key = None
        print("")
        with open(self.project_keys_json_file_path) as project_keys:
            project_keys_json = json.load(project_keys)

        with open(self.project_names_json_file_path) as project_names:
            project_names_json = json.load(project_names)

        while project_key is None:
            try:
                for key, value in project_keys_json.items():
                    try:
                        project_name_json = project_names_json[value]
                    except KeyError:
                        project_name_json = value

                    print("\t" + key + " - " + value + " - " + project_name_json)

                project_key = input("\nEnter the number corresponding the projects (separated by comma) that you need the report for (E.g. 1,2,3):  ")
                project_key = project_key.replace(" ", "")
                project_keys = project_key.split(",")

                for k in project_keys:
                    self.project_keys.append(project_keys_json[k])

                for p in self.project_keys:
                    self.project_names.append(project_names_json.get(p))


            except ValueError as ve:
                print("\tError: ", ve)
                print("Please enter the Project Name again by referring to the list above")





