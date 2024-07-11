from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import pandas as pd


def main():
    pd.set_option('display.max_columns', 500)

    raw_df = pd.read_csv('all_delays.csv')
    station_analysis = raw_df.groupby(['Station'])['Day'].count().reset_index()
    station_analysis.columns = ['Station', 'CountOfDelays']
    station_analysis = station_analysis.sort_values(by=['CountOfDelays'], ascending=False)


    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(station_analysis, index=True, header=True):
        ws.append(r)


    chart1 = BarChart()
    values = Reference(ws, min_col=3, min_row=3, max_row=12)
    cats = Reference(ws, min_col=2, min_row=3, max_row=12)

    chart1.title = "Delay Codes"
    chart1.add_data(values, titles_from_data=True)
    chart1.set_categories(cats)

    ws.add_chart(chart1, "E2")

    print(values)
    print(cats)

    
    wb.save('aggregation.xlsx')


if __name__ == "__main__":
    main()
