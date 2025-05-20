import win32com.client
import os
import glob

from pptx import Presentation

from ITSM_PPT.changeRequestPPT import ChangeRequestPPT
from ITSM_PPT.incidentPPT import IncidentPPT
from ITSM_PPT.requestPPT import RequestPPT


class WriteFilePPT:

    def __init__(self, input_util):
        self.input_util = input_util
        self.date_range = self.input_util.get_current_week_date_range_file_name()
        self.current_week_date = self.input_util.get_current_week_date_range_file_name()
        self.current_week_date_full = self.input_util.get_current_week_date_range_full()

    def write(self):


        for project in self.input_util.get_project_names():

            ppt_template = Presentation("Template\\GCO - HP Weekly ITSM Report - Template.pptx")

            incident_ppt = IncidentPPT(ppt_template, project, self.current_week_date, self.current_week_date_full)
            incident_ppt.fill_data()
            cr_ppt = ChangeRequestPPT(ppt_template, project, self.current_week_date)
            cr_ppt.fill_data()
            req_ppt = RequestPPT(ppt_template, project, self.current_week_date)
            req_ppt.fill_data()

            output_file_name = "GCO - Weekly ITSM Report - " + project + " " + self.date_range + ".pptx"
            output_ppt_path = "Output/" + output_file_name
            ppt_template.save(output_ppt_path)



