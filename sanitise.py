import string
from csv import reader, writer
from openpyxl import workbook
from openpyxl import load_workbook

class Lists:
    def __init__(self, lst):
        self.lst = lst

    def merge_sublists(self):
        new_list = []
        for i in self.lst:
            new_list += i
        self.lst = new_list
        return self.lst

def letters(end):
    # Creates a list appending letters (only up to 26)
    letters = list(string.ascii_uppercase)
    counter = 0

    while counter < end:
        print(f'Berth Code {letters[counter]}')
        counter += 1

def numbers(end):
    # Creates a list appending numbers
    for i in range(1, end):
        print(f'BC{i}')

def replace():
    # BCx list
    pc = [1000 + x for x in list(string.ascii_uppercase)]

    # From the dimensions table
    with open('ttest.csv', 'r') as f:
        read = reader(f)
        og_pc = [row for row in read]

        tmp = Lists(og_pc)
        og_pc = tmp.merge_sublists()

    # From the values table
    with open('test1.csv', 'r') as f:
        read = reader(f)
        values_pc = [row for row in read]

        tmp = Lists(values_pc)
        values_pc = tmp.merge_sublists()

    # Create dictionary from dimensions table and PCx
    sanitisation = {}
    for i in range(len(pc)):
        sanitisation[og_pc[i]] = pc[i]

    # Compare dimensions and values tables list, convert to PCx list
    final_list = []
    for i in range(len(values_pc)):
        final_list.append(sanitisation.get(values_pc[i]))

    # Write the updated list to a csv
    with open('output.csv', 'w', newline='') as file:
        write = writer(file)
        for sanitised in final_list:
            write.writerow([sanitised])

def match():
    # Get the changes required as inputs
    change = {}
    # Creates a dictionary of the user's inputs
    while True:
        try:
            items = input('Terms to look for (enter one at a time, input end when finished): ')
            # Ensures input is a string
            # Check if the input is just a number (letters + numbers are accepted)
            if items.isdigit():
                raise AttributeError
            elif items == 'end':
                break
            else: 
                # Ask for change and input both into the dictionary
                updates = input('Change to: ')
                change[items] = updates
                continue
        except AttributeError:
            print(f'Only text can be replaced')
            continue

    # Open the model
    wb = load_workbook(filename='DC_test_no_queries.xlsm', keep_vba=True, keep_links=True)

    # Iterate over the sheets and read the values of each cell
    for ws in wb:
        for row in ws.rows:
            for cell in row:
                # Loop through the original inputs keys
                for item in change.keys():
                    # cell is the actual cell reference and cell.value gives the value inside the reference
                    try:
                        if cell.value is not None:
                            if item.lower() in cell.value.lower():
                                # Need to replace only the instance and not the whole cell value
                                # Change the string based on location rather than replace (due to case)
                                position = cell.value.lower().find(item.lower())
                                tmp = cell.value
                                if position == 0:
                                    cell.value = ''.join((change[item], tmp[len(item):]))
                                else:
                                    cell.value = ''.join((tmp[:position], change[item], tmp[position + len(item):]))
                                cell.value.replace('  ', ' ')
                                # Clean the cells that now have a space at the start
                                if cell.value[0] == ' ':
                                    temp = cell.value
                                    cell.value = temp[1:]
                        else:
                            continue
                    except AttributeError:
                        # AttributeError occurs when the cell contents is not a string
                        continue
    
    wb.save('sanitised_workbook.xlsm')