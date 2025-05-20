import pandas as pd

from ITSM_Excel.Excel_FetchData.fetchData import FetchData


class OverallRequestSummary:
    def __init__(self, fetch_data):
        self.request_sheet_data = fetch_data.get_request_sheet_data()  # getting data from excel in List o list format

    def get_data(self):
        data = []
        df = pd.DataFrame(self.request_sheet_data[1:], columns=self.request_sheet_data[0])
        items_list = df.groupby('Item')['Number'].count().sort_values(ascending=False).index.values.tolist()
        # To calculate Grant Total
        grand_total = 0

        for item in items_list:
            df2 = df[df['Item'] == item]
            overall_grouped = df2.groupby(["Item", "Short description"])
            overall = overall_grouped['Number'].count().sort_values(ascending=False)
            overall_index = overall.index.values.tolist()
            i = 0
            for entries in overall:
                l = list(overall_index[i])
                l.append(overall[overall_index[i]])  # Appending the value(count) to the end of list
                grand_total += overall[overall_index[i]]
                data.append(l)
                i += 1

        # grand total row
        grand_total_list = ["Grand Total", '', grand_total]
        data.append(grand_total_list)

        # add header row
        header_list = ["Item", 'Short description', 'Total']
        data.insert(0, header_list)

        return data
