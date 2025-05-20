import pandas as pd
from ITSM_Excel.Excel_FetchData.fetchData import FetchData


class Top5IncidentTypes:
    def __init__(self, fetch_data):
        self.incident_sheet_data = fetch_data.get_incident_sheet_data()

    def get_data(self):
        data = []
        df = pd.DataFrame(self.incident_sheet_data[1:], columns=self.incident_sheet_data[0])
        top5_grouped = df.groupby("Short description")
        top5 = top5_grouped['Number'].count().sort_values(ascending=False).head(5)
        top5_index = top5.index.values.tolist()

        # To calculate Grand Total
        grand_total = 0
        i = 0
        for t in top5:
            l = []
            l.append(top5_index[i])
            l.append(top5[top5_index[i]])
            grand_total += top5[top5_index[i]]
            data.append(l)
            i += 1

        grand_total_list = ["Grand Total", grand_total]
        data.append(grand_total_list)

        # to add header row
        header_row = ['Short Description', 'Total']
        data.insert(0, header_row)

        return data
