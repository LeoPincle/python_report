import pandas as pd

from ITSM_Excel.Excel_FetchData.fetchData import FetchData


class Top5CRTypes:
    def __init__(self, fetch_data):
        self.change_request_sheet_data = fetch_data.get_change_request_sheet_data()
        self.start_date = fetch_data.start_date
        self.end_date = fetch_data.end_date

    def get_data(self):
        data = []
        df = pd.DataFrame(self.change_request_sheet_data[1:], columns=self.change_request_sheet_data[0])
        mask = (df['Created'] >= self.start_date) & (df['Created'] <= self.end_date)
        df = df[mask]
        top5_grouped = df.groupby("Short description")
        top5 = top5_grouped['Number'].count().sort_values(ascending=False).head(5)
        top5_index = top5.index.values.tolist()

        # Calculate Grand Total
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

        # add header row
        header_list = ['Short Description', 'Total']
        data.insert(0, header_list)

        return data
