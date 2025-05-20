from ITSM_Excel.Excel_FetchData.fetchData import FetchData
import pandas as pd


class OverallCRSummary:

    def __init__(self, fetch_data):
        self.change_request_data = fetch_data.get_change_request_sheet_data()
        self.state_list = ['New', 'Assess', 'Authorize', 'Scheduled', 'Implement', 'Review', 'Closed', 'Canceled']

    def get_data_list_by_state(self, grouped_df):
        data_list = []
        sum = 0
        for state in self.state_list:
            try:
                data_list.append(grouped_df['Number'].count()[state])
                sum = sum + grouped_df['Number'].count()[state]
            except KeyError:
                data_list.append(0)
        data_list.append(sum)
        return data_list

    def get_data(self):
        data = []
        df = pd.DataFrame(self.change_request_data[1:], columns=self.change_request_data[0])
        overall_cr_summary_grouped = df.groupby('Final Assignment Group')
        overall_cr_summary = overall_cr_summary_grouped['Number'].count().sort_values(ascending=False)
        overall_cr_summary_index = overall_cr_summary.index.values.tolist()
        for ag in overall_cr_summary_index:
            df_final = df[df["Final Assignment Group"] == ag]
            total_cr_grouped = df_final.groupby("State")
            data_ag = self.get_data_list_by_state(total_cr_grouped)
            data_ag.insert(0, ag)
            data.append(data_ag)

        # add grand total row
        grand_total_list = [0] * 9
        grand_total_list.insert(0, "Grand Total")
        for l in data:
            for i in range(1, 10):
                grand_total_list[i] += l[i]
        data.append(grand_total_list)

        # Adding header
        header_list = self.state_list.copy()
        header_list.insert(0, 'Assignment Group')
        header_list.append('Total')
        data.insert(0, header_list)

        return data

