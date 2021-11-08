#test python script for xls sheet automation
import openpyxl

Ratings_File = openpyxl.load_workbook("NoteRating.xlsx")
Rating_List = Ratings_File["Sheet1"]

# sheet.maxrow will give the info of how many rows are filled uo in the sheet
# to run a for loop we need to cretae a list of numbers inorder to loop thru each item in that list
# to get that list we need to use range funtion.
print(range(Rating_List.max_row))

# range starts from value 0, our sheet has no 0, start reading the rows only from row 2 since row 1 is title
# create a dictionary with key-value pair -  childname as key and number of tests as value
tests_per_child = {}

for each_row in range(2, Rating_List.max_row+1):
    Child_name = Rating_List.cell(each_row,1).value

    if Child_name in tests_per_child:
        tests_per_child[Child_name] += 1
    else:
        tests_per_child[Child_name] = 1



print(tests_per_child)

#create new worksheet
workbook = Ratings_File.create_sheet("Mysheet", 0) # insert at the end (default)
workbook.title = "Dashboard"
Ratings_File.save("NoteRating.xlsx")