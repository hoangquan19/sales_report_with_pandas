import pandas as pd
import xlsxwriter

pd.options.display.width= None
pd.options.display.max_columns= None
pd.set_option('display.max_rows', 3000)
pd.set_option('display.max_columns', 3000)

data = pd.read_csv("OfficeSupplies.txt")

data["Sales"] = data["Units"] * data["Unit Price"]

data["Date"] = data["OrderDate"].str.split("-")
data["Month"] = data["Date"].apply(lambda x: x[1])

data = data.drop(columns=["Date"])

sales_by_month = data.groupby(by="Month").agg({"Sales": "sum"})
sales_by_month = sales_by_month.reset_index()
# Khởi tạo Workbook và Worksheet
workbook = xlsxwriter.Workbook('chart.xlsx')
worksheet = workbook.add_worksheet()

# Ghi dữ liệu vào Worksheet
worksheet.write_column('A1', sales_by_month["Month"])
worksheet.write_column('B1', sales_by_month["Sales"])

# Tạo biểu đồ
n = len(sales_by_month.index)
chart = workbook.add_chart({'type': 'bar'})
chart.add_series({'categories': f'=Sheet1!$A$2:$A${n}',
                  'values': f'=Sheet1!$B$1:$B${n}'})

# Thêm biểu đồ vào Worksheet
worksheet.insert_chart('D1', chart)

# Đổi tên biểu đồ
chart.set_title({'name': 'Sales by Month'})

# Lưu Workbook
workbook.close()