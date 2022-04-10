from src.duct import Duct
from src.category import Category
import csv
import os
import xlsxwriter

csv_file_list = os.listdir(os.getcwd() + '/data')

category1 = Category('1', 1.00)
category2 = Category('2', 2.50)
category3 = Category('3', 4.00)
category4 = Category('4', 5.00)
category5 = Category('5', 6.00)

def read_CSV_file(filename):
    with open(os.getcwd() + '/data/' + filename) as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            if row['Size'] == '':
                pass
            else:
                component_in_row = Duct(row['Size'], row['Area'], row['Surface Area'])
                component_in_row.set_width_and_height()
                component_in_row.set_duct_category()
                add_component_to_relevant_category(component_in_row)

def write_CSV_file(filename):
    with open(os.getcwd() + '/output/csv/' + filename[0:-4] + '(Quantified).csv', 'w', newline='') as new_file:
        calculate_category_costs()
        csv_writer = csv.writer(new_file)
        csv_writer.writerow(['category', 'quantity', 'rate', 'cost'])
        csv_writer.writerow([category1.categoryNo, round(category1.quantity, 3), category1.rate, round(category1.cost, 2)])
        csv_writer.writerow([category2.categoryNo, round(category2.quantity, 3), category2.rate, round(category2.cost, 2)])
        csv_writer.writerow([category3.categoryNo, round(category3.quantity, 3), category3.rate, round(category3.cost, 2)])
        csv_writer.writerow([category4.categoryNo, round(category4.quantity, 3), category4.rate, round(category4.cost, 2)])
        csv_writer.writerow([category5.categoryNo, round(category5.quantity, 3), category5.rate, round(category5.cost, 2)])
        total = category1.cost + category2.cost + category3.cost + category4.cost + category5.cost
        csv_writer.writerow(['total', round(total, 2)])

def write_Excel_file(filename):
    workbook = xlsxwriter.Workbook(os.getcwd() + '/output/excel/' + filename[0:-4] + '(BOQ).xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'Item', bold)
    worksheet.write('B1', 'Description', bold)
    worksheet.write('C1', 'Unit', bold)
    worksheet.write('D1', 'Quantity', bold)
    worksheet.write('E1', 'Rate', bold)
    worksheet.write('F1', 'Amount', bold)

    worksheet.write('B3', filename[0:-4], bold)

    worksheet.write('B5', )

    workbook.close()

def add_component_to_relevant_category(component):
    if component.category == '1':
        category1.add_component(component)
    elif component.category == '2':
        category2.add_component(component)
    elif component.category == '3':
        category3.add_component(component)
    elif component.category == '4':
        category4.add_component(component)
    elif component.category == '5':
        category5.add_component(component)
    else:
        print(f"Category {component.category} not found")

def calculate_category_costs():
    category1.calculate_cost()
    category2.calculate_cost()
    category3.calculate_cost()
    category4.calculate_cost()
    category5.calculate_cost()

def clear_categories():
    category1.clear_category()
    category2.clear_category()
    category3.clear_category()
    category4.clear_category()
    category5.clear_category()

for file in csv_file_list:
    read_CSV_file(file)
    write_CSV_file(file)
    write_Excel_file(file)
    clear_categories()
