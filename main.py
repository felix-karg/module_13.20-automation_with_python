import openpyxl

# Open file and save it in a variable
inv_file = openpyxl.load_workbook("inventory.xlsx")
# Save content of the sheet named 'Sheet1' in another variable
product_list = inv_file["Sheet1"]


### Exercise 1: Calculate number of products per supplier ###
products_per_supplier = {}
# Start loop from second row to skip headers
for product_row in range(2, product_list.max_row + 1):
    # extract value of 4th column of each row
    supplier_name = product_list.cell(product_row, 4).value

    if supplier_name in products_per_supplier:
        products_per_supplier[supplier_name] = products_per_supplier[supplier_name] + 1
    else:
        products_per_supplier[supplier_name] = 1

print(products_per_supplier)


### Exercise 2: Calculate total inventory value per supplier ###
total_value_per_supplier = {}

for product_row in range(2, product_list.max_row + 1):
    supplier_name = product_list.cell(product_row, 4).value
    inventory = product_list.cell(product_row, 2).value
    price = product_list.cell(product_row, 3).value
    if supplier_name in total_value_per_supplier:
        current_total_value = total_value_per_supplier.get(supplier_name)
        total_value_per_supplier[supplier_name] = current_total_value + inventory * price
    else:
        total_value_per_supplier[supplier_name] = inventory * price

print(total_value_per_supplier)


### Exercise 3: Print products with invetory < 10
product_under_10_inv = {}

for product_row in range(2, product_list.max_row + 1):
    inventory = int(product_list.cell(product_row, 2).value)
    product_num = int(product_list.cell(product_row, 1).value)
    if inventory < 10:
        product_under_10_inv[product_num] = inventory

print(product_under_10_inv)


### Exercise 4: Add column with total value to file
for product_row in range(2, product_list.max_row + 1):
    inventory = product_list.cell(product_row, 2).value
    price = product_list.cell(product_row, 3).value
    inventory_price = product_list.cell(product_row, 5)
    inventory_price.value = inventory * price

# save changes to a new file
inv_file.save("inventory_with_total_value.xlsx")
