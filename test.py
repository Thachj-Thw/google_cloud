from google_cloud import GoogleSheet

sheet = "Trang tính1"
spreadsheet = GoogleSheet("D:\\Python\\MyLib\\google_cloud\\test\\key.json", "1KjUq04WQcy4l5XoXDQR-uOAxoSUqrCB2eU-Q4303lJY")
# spreadsheet.write_row(sheet, "D7", [10, 1, 2, 4, 5, 6])
print(spreadsheet.write(sheet, "E6:H6", [[1, 2, 3, 4]]))
# spreadsheet.copy_paste(sheet, "H1:H5", "A1:D5")
