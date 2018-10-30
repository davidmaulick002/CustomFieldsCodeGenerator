
import xlrd
import xlwt
import string
import sys

ASSIGNEE_CF_NAME_COLUMN, ASSIGNEE_CF_ID_COLUMN = 1, 0
ASSIGNMENT_CF_NAME_COLUMN, ASSIGNMENT_CF_ID_COLUMN = 1, 0
COUNTRY_CF_NAME_COLUMN, COUNTRY_CF_ID_COLUMN = 4, 0
FIRST_OUT_COL, SECOND_OUT_COL, THIRD_OUT_COL = 0, 3, 6


def createCustomFieldCodes(customFieldSheet, customFieldCodeArr, customFieldNameArr, customFieldTitleCol, customFieldIdCol):

    for i in range(customFieldSheet.nrows):

        rowCustomFieldName = customFieldSheet.cell_value(i, customFieldTitleCol)
        customFieldNameArr.append(rowCustomFieldName)

        rowCustomFieldId = str(customFieldSheet.cell_value(i, customFieldIdCol)).replace('.0', '') # workaround to get rid of float trailing '.0'
        rowCustomFieldOutputCode = rowCustomFieldName.replace(' ', '_') + rowCustomFieldId
        customFieldCodeArr.append(rowCustomFieldOutputCode)


def outputCodesToFile(customFieldNameArr, customFieldCodeArr, outSheet, cfNameCol):

    cfCodeCol = cfNameCol + 1

    for i in range(len(customFieldCodeArr)):
        outSheet.write(i, cfNameCol, str(customFieldNameArr[i]))
        outSheet.write(i, cfCodeCol, str(customFieldCodeArr[i]))


if __name__ == '__main__':
    assigneeCustomFieldFile = "C:\\Users\\dmaulick002\\Documents\\AFileBox\\CustomFieldData\\assignee.xlsx"  # custom field is 2nd collum (index 1), id is index 0
    assignmentCustomFieldFile = "C:\\Users\\dmaulick002\\Documents\\AFileBox\\CustomFieldData\\assignee_assignment.xlsx"  # custom field is 2nd column (index 1), id is index 0
    countryCustomFieldFile = "C:\\Users\\dmaulick002\\Documents\\AFileBox\\CustomFieldData\\custom_country_with_id.xlsx"  # custom field is 5TH column (index4), ID is index 0

    outputFile = "C:\\Users\\dmaulick002\\Documents\\AFileBox\\CustomFieldData\\customFieldCodeOutput.csv"



    assigneeSheet = xlrd.open_workbook(assignmentCustomFieldFile).sheet_by_index(0)
    assignmentSheet = xlrd.open_workbook(assignmentCustomFieldFile).sheet_by_index(0)
    countrySheet = xlrd.open_workbook(countryCustomFieldFile).sheet_by_index(0)


    outWb = xlwt.Workbook()
    outputSheet =outWb.add_sheet('output')

    assigneeNameArr = [string]
    assigneeCodeArr = [string]

    assignmentNameArr = [string]
    assignmentCodeArr = [string]

    countryNameArr = [string]
    countryCodeArr = [string]

    createCustomFieldCodes(assigneeSheet, assigneeCodeArr, assigneeNameArr, ASSIGNEE_CF_NAME_COLUMN, ASSIGNEE_CF_ID_COLUMN)
    createCustomFieldCodes(assignmentSheet, assignmentCodeArr, assignmentNameArr, ASSIGNMENT_CF_NAME_COLUMN, ASSIGNMENT_CF_ID_COLUMN)
    createCustomFieldCodes(countrySheet, countryCodeArr, countryNameArr, COUNTRY_CF_NAME_COLUMN, COUNTRY_CF_ID_COLUMN)

    outputCodesToFile(assigneeNameArr, assigneeCodeArr, outputSheet, FIRST_OUT_COL)
    outputCodesToFile(assignmentNameArr, assigneeCodeArr, outputSheet, SECOND_OUT_COL)

    outWb.save(outputFile)
    #outputCodesToFile(countryNameArr, assigneeCodeArr, outputSheet, THIRD_OUT_COL)

    if len(assigneeCodeArr) > len(set(assigneeCodeArr)):
        print("Assignee custom fields were not converted to a unique set of codes.")
        sys.exit(1)

    if len(assignmentCodeArr) > len(set(assignmentCodeArr)):
        print("Assignment custom fields were not converted to a unique set of codes.")
        sys.exit(1)

    if len(countryCodeArr) > len(set(countryCodeArr)):
        print("Country custom fields were not converted to a unique set of codes.")
        sys.exit(1)

    print("All codes sets are fully unique")



