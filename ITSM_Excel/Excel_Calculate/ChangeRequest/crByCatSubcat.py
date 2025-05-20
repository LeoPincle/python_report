import pandas as pd
from ITSM_Excel.Excel_FetchData.fetchData import FetchData


class CRByCatSubcat:
    def __init__(self, fetch_data):
        self.change_request_sheet_data = fetch_data.get_change_request_sheet_data() # getting data from excel in List o list format

    def get_data(self):
        data = []
        df = pd.DataFrame(self.change_request_sheet_data[1:], columns=self.change_request_sheet_data[0])

        category_list = df.groupby('Category')['Number'].count().sort_values(ascending=False).index.values.tolist()
        # To calculate Grand Total
        grand_total = 0

        for category in category_list:
            df2 = df[df['Category'] == category]
            overall_grouped = df2.groupby(["Category", "Subcategory"])
            overall = overall_grouped['Number'].count().sort_values(ascending=False)
            overall_index = overall.index.values.tolist()
            i = 0
            for _ in overall:
                l = list(overall_index[i])
                l.append(overall[overall_index[i]])  # Appending the value (count) in the end of list
                grand_total += overall[overall_index[i]]
                data.append(l)
                i += 1

        grand_total_list = ["Grand Total", '', grand_total]
        data.append(grand_total_list)

        # to add header row
        header_list = ["Category", 'Subcategory', 'Total']
        data.insert(0, header_list)

        return data
