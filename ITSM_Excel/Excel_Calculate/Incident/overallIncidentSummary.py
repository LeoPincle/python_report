import numpy
import pandas as pd
from pandas.core.groupby.groupby import DataError
import openpyxl
from ITSM_Excel.Excel_FetchData.fetchData import FetchData


class OverallIncidentSummary:

    def __init__(self, fetch_data):
        self.incident_sheet_data = fetch_data.get_incident_sheet_data()
        self.response_resolution_sheet_data = fetch_data.get_response_resolution_sheet_data()
        self.priority_list = ["1 - Critical", "2 - High", "3 - Moderate", "4 - Low", "5 - Planning"]
        self.start_date = fetch_data.start_date
        self.end_date = fetch_data.end_date


    def get_data_list(self, grouped_df):
        datalist = []
        sum = 0
        for priority in self.priority_list:
            try:
                datalist.append(grouped_df["Number"].count()[priority])
                # print(priority, grouped_df["Number"].count())
                sum = sum + grouped_df["Number"].count()[priority]
            except KeyError as ke:
                datalist.append(0)
            except DataError:
                datalist.append(0)
        datalist.append(sum)
        return datalist

    def get_data_list_BET_hrs(self, grouped_df):
        datalist = []
        for priority in self.priority_list[0:4]:
            try:
                datalist.append(format(grouped_df['BET in hrs'].mean()[priority], ".2f"))
            except KeyError:
                datalist.append('')
            except DataError:
                datalist.append('')
        datalist.append('NA')
        datalist.append('NA')
        return datalist

    def get_data_list_resp_resl_perct(self, grouped_df, grouped_df_total):
        datalist = []
        for priority in self.priority_list[0:4]:
            try:
                div = grouped_df['Number'].count()[priority] / grouped_df_total['Number'].count()[priority]
                div = "{:.0%}".format(div) if div == 1 else "{:.2%}".format(div)
                datalist.append(div)
            except KeyError:
                datalist.append('')
            except DataError:
                datalist.append('')
        datalist.append('NA')
        datalist.append('NA')
        return datalist

    def get_data(self):
        data = []
        df1 = pd.DataFrame(self.incident_sheet_data[1:], columns=self.incident_sheet_data[0])

        # Open Incidents
        df_open = df1[df1["Final State"].eq("Open")]
        prod_grouped = df_open.groupby("Priority")
        data.append(self.get_data_list(prod_grouped))

        # Closed Incidents
        df_closed = df1[df1["Final State"].eq("Closed")]
        closed_grouped = df_closed.groupby("Priority")
        data.append(self.get_data_list(closed_grouped))

        # Total Incidents
        df_total = df1.groupby("Priority")
        data.append(self.get_data_list(df_total))

        df2 = pd.DataFrame(self.response_resolution_sheet_data[1:], columns=self.response_resolution_sheet_data[0])

        # Backlog
        mask1 = (df2['Resolved'] >= self.start_date) & (df2['Resolved'] <= self.end_date)
        mask2 = (df2['Created'] < self.start_date) | (df2['Created'] > self.end_date)
        df_backlog = df2[mask1 & mask2]
        df_backlog = df_backlog[df_backlog['SLA definition'].str.contains(pat=r'resolution', regex=True, case=False)]
        df_backlog = df_backlog[df_backlog['Stage'] != "Cancelled"]
        backlog_grouped = df_backlog.groupby("Priority")
        data.append(self.get_data_list(backlog_grouped))

        # Response
        df_response = df2[df2['State'] != 'New']
        df_response = df_response[df_response['Stage'] != 'Cancelled']
        df_response = df_response[df_response['SLA definition'].str.contains(pat=r'response', regex=True, case=False)]
        df_response_violated = df_response[df_response['Has violated'] == False]
        response_grouped = df_response.groupby("Priority")
        response_grouped_violated = df_response_violated.groupby("Priority")
        data.append(self.get_data_list_resp_resl_perct(response_grouped_violated, response_grouped))

        # Response Avg in hours
        df_response_avg_hr = df2[df2['State'] != 'New']
        df_response_avg_hr = df_response_avg_hr[df_response_avg_hr['Stage'] != 'Cancelled']
        df_response_avg_hr = df_response_avg_hr[df_response_avg_hr['SLA definition'].str.contains(pat=r'response',
                                                                                                  regex=True,
                                                                                                  case=False)]
        response_avghr_grouped = df_response_avg_hr.groupby("Priority")
        data.append(self.get_data_list_BET_hrs(response_avghr_grouped))  # Flag

        # Resolution
        df_resolution = df2[df2['State'].isin(['Closed', 'Resolved'])]
        df_resolution = df_resolution[df_resolution['Stage'] != 'Cancelled']
        mask = (df_resolution['Resolved'] >= self.start_date) & (df_resolution['Resolved'] <= self.end_date)
        df_resolution = df_resolution[mask]
        df_resolution = df_resolution[df_resolution['SLA definition'].str.contains(pat=r'resolution',
                                                                                   regex=True, case=False)]
        df_resolution_violated = df_resolution[df_resolution['Has violated'] == False]
        resolution_grouped = df_resolution.groupby("Priority")
        resolution_grouped_violated = df_resolution_violated.groupby("Priority")
        data.append(self.get_data_list_resp_resl_perct(resolution_grouped_violated, resolution_grouped))

        # Resolution Avg in hours
        df_resolution_avg_hr = df2[df2['State'].isin(['Closed', 'Resolved'])]
        df_resolution_avg_hr = df_resolution_avg_hr[df_resolution_avg_hr['Stage'] != 'Cancelled']
        mask = (df_resolution_avg_hr['Resolved'] >= self.start_date) & (
                    df_resolution_avg_hr['Resolved'] <= self.end_date)
        df_resolution_avg_hr = df_resolution_avg_hr[mask]
        df_resolution_avg_hr = df_resolution_avg_hr[
            df_resolution_avg_hr['SLA definition'].str.contains(pat=r'resolution', regex=True, case=False)]
        resolution_avg_hr_grouped = df_resolution_avg_hr.groupby("Priority")
        data.append(self.get_data_list_BET_hrs(resolution_avg_hr_grouped))

        # Prod Violated
        df_prod_violated = df2[df2['Final State'] == 'Closed']
        mask = (df_prod_violated['Resolved'] >= self.start_date) & (df_prod_violated['Resolved'] <= self.end_date)
        df_prod_violated = df_prod_violated[mask]
        df_prod_violated = df_prod_violated[
            df_prod_violated['SLA definition'].str.contains(pat=r'resolution', regex=True, case=False)]
        df_prod_violated = df_prod_violated[df_prod_violated["Final Environment"] == 'Prod']
        df_prod_violated = df_prod_violated[df_prod_violated['Has violated'] == True]
        prod_violated_grouped = df_prod_violated.groupby("Priority")
        data.append(self.get_data_list(prod_violated_grouped))

        # Converting data into row-wise list of list format
        priority_column = self.priority_list.copy()
        priority_column.append('Total')
        data.insert(0, priority_column)  # adding first column to the data containing priority list
        data_temp = numpy.array(data)
        data = data_temp.transpose().tolist()  # transposing list to store data in row-wise order
        header_list = ['Priority', 'Open State', 'Closed State', 'Total Created', 'Backlog', 'Response',
                       'Response Avg. in Hrs', 'Resolution', 'Resolution Avg. in Hrs', 'PRD Violated']
        data.insert(0, header_list)  # appending the header row

        return data

    def get_piechart_data(self):
        data = []
        df = pd.DataFrame(self.incident_sheet_data[1:], columns=self.incident_sheet_data[0])
        df_total = df.groupby("Priority")
        data.append(self.get_data_list(df_total))
        data[0].pop()
        priority_column = self.priority_list.copy()
        data.insert(0, priority_column)
        return data


'''''
# Testing purpose only
fetch_data = FetchData('2024-11-11', '2024-11-17', 'PROJ0040158')
inc_created = OverallIncidentSummary(fetch_data)
inc_created.get_data()
'''
