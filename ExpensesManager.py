from openpyxl import load_workbook
import xml.etree.ElementTree as ET


# Classes:
class Expense:
    def __init__(self, BusinessName, TransactionSum, ChargeAmount) :
        self.BusinessName = BusinessName
        self.TransactionSum = TransactionSum
        self.ChargeAmount = ChargeAmount

class CategoryCounter:
    def __init__(self, CategoryName, Counter) :
        self.CategoryName = CategoryName
        self.Counter = Counter


# Helpers functions:
def isThisStringHasHebrewChars(s):
    return any("\u0590" <= c <= "\u05EA" for c in s)
    
def getExpensesList():
    ThisMonthExpensesFileName = "ThisMonthExpenses.xlsx"
    wb = load_workbook(filename = ThisMonthExpensesFileName)
    ws = wb.active
    ThisMonthExpensesList = []
    for row in ws.iter_rows(min_row=4, max_col=5, max_row=127):
        if isThisStringHasHebrewChars(str(row[1].value)):
            BusinessName = str(row[1].value)[::-1]
        else:
            BusinessName = str(row[1].value)
        TransactionSum = str(row[2].value)
        ChargeAmount = str(row[3].value)
        ThisMonthExpensesList.append(Expense(BusinessName, TransactionSum, ChargeAmount))
    
    return ThisMonthExpensesList

def initiallizeCategoriesCounters():
    xml_file_name = 'ExpenseManagerConfigFile.xml'
    XML_tree = ET.parse(xml_file_name)
    XML_root = XML_tree.getroot()
    CategoriesCountersList = []

    for category in XML_root.iter('Category'):
        should_i_insert = True
        for cc in CategoriesCountersList:
            if cc.CategoryName == category.text:
                should_i_insert = False
                break
        if should_i_insert:
            CategoriesCountersList.append(CategoryCounter(category.text, 0))

    return CategoriesCountersList





if __name__ == "__main__":
    # get all this month expenses in  a list
    ThisMonthExpensesList = getExpensesList()

    #Initialize counters for each known category:
    CategoriesCountersList = initiallizeCategoriesCounters()

    #iterate over the list:
    xml_file_name = 'ExpenseManagerConfigFile.xml'
    XML_tree = ET.parse(xml_file_name)
    for expense in ThisMonthExpensesList:
        
        #if the business named exists in the xml add to its category counter the sum of the charge
        category = ""
        for business in XML_tree.getroot():
            if business[0].text == expense.BusinessName:
                category = business[1].text
                break
        if category != "":
            for categoryCounter in CategoriesCountersList:
                if categoryCounter.CategoryName == category:
                    categoryCounter.Counter = float(categoryCounter.Counter) + float(expense.ChargeAmount)
                    break
        else:
            # else, ask the user what is this business category,
            print("\nhello, from the following categories: ")
            i = 0
            for categotyCounter in CategoriesCountersList:
                print(str(i) + ". " +categotyCounter.CategoryName)
                i= i+1
            print(str(i) + ". " + "New Category\n")
            try:
                categoryNumber = input("to which one belongs the business named: " +  expense.BusinessName + "?\n")    
                categoryNumber = int(categoryNumber)
            except ValueError:
                categoryNumber = input("That's not an int!, insert number between 0 to "+str(i) + ".\n")
                categoryNumber = int(categoryNumber)
            if categoryNumber<0 or categoryNumber>i :
                categoryNumber = input("illegal number")
            if categoryNumber < i:
                category = CategoriesCountersList[categoryNumber].CategoryName
                CategoriesCountersList[categoryNumber].Counter = CategoriesCountersList[categoryNumber].Counter + float(expense.ChargeAmount)
            else:
                category = input("so what is the category then ?\n")
                CategoriesCountersList.append(CategoryCounter(category, float(expense.ChargeAmount)))
                print("\n")

            # add this business with his category to the xml, add the category to the list of the counters 

            new_business = ET.SubElement(XML_tree.getroot(), 'Business')
            ET.SubElement(new_business, 'BusinessName').text = expense.BusinessName
            ET.SubElement(new_business, 'Category').text = category

    print("your total is: ")
    for categotyCounter in CategoriesCountersList:
        print("wasted on "+categotyCounter.CategoryName +" " + str(categotyCounter.Counter) +  "NIS")

    XML_tree.write("ExpenseManagerConfigFile.xml")



    
