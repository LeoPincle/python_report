import pandas as pd
from ITSM_Excel.Excel_FetchData.fetchData import FetchData


class IncByCatSubcat:

    def __init__(self, fetch_data):
        self.incident_sheet_data = fetch_data.get_incident_sheet_data()
        self.priority_list = ["1 - Critical", "2 - High", "3 - Moderate", "4 - Low", "5 - Planning"]

    def getDataList(self,grouped_df):
        datalist = []
        sum = 0
        for priority in self.priority_list:
            try:
                datalist.append(grouped_df['Number'].count()[priority])
                sum = sum + grouped_df['Number'].count()[priority]
            except KeyError:
                datalist.append(0)
        datalist.append(sum)
        return datalist

    def get_subcategory_data(self, category, data):
        df = pd.DataFrame(self.incident_sheet_data[1:], columns=self.incident_sheet_data[0])
        df = df[df["Category"] == category]
        subcat_grouped = df.groupby(["Subcategory"])
        subcategory = subcat_grouped['Number'].count().sort_values(ascending=False)
        subcategory_index = subcategory.index.values.tolist()
        for subcat in subcategory_index:
            temp_data = self.get_priority_data(category, subcat)
            temp_data.insert(0, category)
            data.append(temp_data)

    def get_priority_data(self, category, subcategory):
        df = pd.DataFrame(self.incident_sheet_data[1:], columns=self.incident_sheet_data[0])
        df = df[df["Category"] == category]
        df = df[df["Subcategory"] == subcategory]
        cloud_grouped = df.groupby(["Priority"])
        data = self.getDataList(cloud_grouped)
        data.insert(0, subcategory)
        return data

    def get_data(self):
        data = []
        df = pd.DataFrame(self.incident_sheet_data[1:], columns=self.incident_sheet_data[0])

        cat_grouped = df.groupby(["Category"])
        category = cat_grouped['Number'].count().sort_values(ascending=False)
        category_index = category.index.values.tolist()

        for cat in category_index:
            self.get_subcategory_data(cat, data)

        # Header row
        header_list = self.priority_list.copy()
        header_list.insert(0, "Category")
        header_list.insert(1, "Subcategory")
        header_list.append('Total')
        data.insert(0, header_list)
        # Grand total row
        grandtotal_list = [0] * len(self.priority_list)
        grandtotal_list.insert(0, "Grand Total")
        grandtotal_list.insert(1, "")
        # For summation
        for l in data[1:]:
            for i in range(2, len(self.priority_list) + 2):
                grandtotal_list[i] += l[i]
        grandtotal_list.append(sum(grandtotal_list[2:]))
        data.append(grandtotal_list)

        return data
