import pandas as pd

# SETTING CONSTANT INDEX.
STATUS = 0
MEMBER_DOB = 1
SPOUSE_DOB = 2
PAYEE_GENDER = 3
SPOUSE_GENDER = 4
PROVINCE = 5
POSTAL_CODE = 6
RETIREMENT_DATE = 7
DEATH_DATE = 8
PENSION = 9
YEARS_GUARANTEE = 10
GUARANTEE_END = 11
UNLOCATED_MEMBER = 12
SURNAME = 13
GIVEN_NAME = 14
# MEMBER_NAME = 9
SPOUSE_SURNAME = 15
SPOUSE_GIVENNAME = 16
# SPOUSE_NAME = 10
MARITAL_STATUS = 17
BEN_SURNAME = 18
BEN_GIVEN = 19
# BEN_NAME = 11


def process_data(filename):
    """
    Precondition: filename is a file in .xlsx format that is in the same folder with this python file.
    Post-condition: Outputs a file name "Review *filename*.xlsx" in the same folder with this python file.
    """
    data = pd.read_excel(filename)
    output_headers = ["ID", "Status", "Member DOB", "Missing DOB?", "Spouse Date of Birth", "Missing Spouse DOB",
                      "Payee Gender", "Member Gender Anamoly", "Spouse Gender", "Spouse Gender Anamoly",
                      "Gender Mismatch",
                      "Province of Residence",
                      "Postal Code", "Postal Code Check",
                      "Original Member's Date of Retirement",
                      "Original Member's Date of Death", "Lifetime Monthly Pension", "Pension Amount Check",
                      "Original Guarantee (Years)", "Date Guarantee End", "Guarantee Check", "Unlocated Member",
                      "Unlocated Check", "Member Name", "Member Name Check", "Spouse Name", "Spouse Name Check",
                      "Marital Status", "Beneficiary Name", "Beneficiary Name Check"]
    existing_headers = ["Status", "Member DOB", "Spouse Date of Birth", "Payee Gender", "Spouse Gender",
                        "Province of Residence",
                        "Postal Code", "Original Member's Date of Retirement",
                        "Original Member's Date of Death", "Lifetime Monthly Pension",
                        "Original Guarantee (Years)", "Date Guarantee End", "Unlocated Member", "Marital Status"]
    constant = [STATUS, MEMBER_DOB, SPOUSE_DOB, PAYEE_GENDER, SPOUSE_GENDER,
                PROVINCE,
                POSTAL_CODE, RETIREMENT_DATE,
                DEATH_DATE, PENSION, YEARS_GUARANTEE, GUARANTEE_END, UNLOCATED_MEMBER, MARITAL_STATUS]
    output = pd.DataFrame(columns=output_headers)
    for i in range(len(constant)):
        output[existing_headers[i]] = data.iloc[:, constant[i]]
    output["ID"] = ["10 digit unique numbers"] * len(data.index)
    output["Missing DOB?"] = output["Member DOB"].isnull()
    output["Missing Spouse DOB"] = output["Spouse Date of Birth"].isnull()
    output["Payee Gender"].fillna("N/A", inplace=True)
    output.loc[:,"Member Gender Anamoly"] = output["Payee Gender"].apply(gender_anamoly)
    output["Spouse Gender"].fillna("N/A", inplace=True)
    output.loc[:,"Spouse Gender Anamoly"] = output["Spouse Gender"].apply(gender_anamoly)
    output["Gender Mismatch"] = output["Payee Gender"] == output["Spouse Gender"]
    output["Postal Code Check"] = output["Postal Code"].apply(check_pc)
    output["Pension Amount Check"] = output["Lifetime Monthly Pension"] > 0
    ########## Member Unlocated Check ############
    output.loc[output["Unlocated Member"] == "Y", "Unlocated Check"] = "Investigate"
    output.loc[output["Unlocated Member"] != "Y", "Unlocated Check"] = "Correct"
    ###################### END #######################
    output["Member Name"] = data.iloc[:, GIVEN_NAME] + " " + data.iloc[:, SURNAME]
    # output["Member Name"] = data.iloc[:, MEMBER_NAME]
    output["Member Name"].fillna("N/A", inplace=True)
    ########## Member Name Check ##################
    output.loc[output["Member Name"] == "N/A", "Member Name Check"] = "No"
    output.loc[output["Member Name"] != "N/A", "Member Name Check"] = "Yes"
    ###################### END #########################
    output["Spouse Name"] = data.iloc[:, SPOUSE_GIVENNAME] + " " + data.iloc[:, SPOUSE_SURNAME]
    output["Beneficiary Name"] = data.iloc[:, BEN_GIVEN] + " " + data.iloc[:, BEN_SURNAME]
    # output["Spouse Name"] = data.iloc[:, SPOUSE_NAME]
    # output["Beneficiary Name"] = data.iloc[:, BEN_NAME]
    ######### Guarantee Check ##################
    output["Date Guarantee End"].fillna("N/A", inplace=True)
    output.loc[(output["Date Guarantee End"] == "N/A") & (
            output["Original Guarantee (Years)"] != 0), "Guarantee Check"] = "Incorrect"
    output.loc[(output["Date Guarantee End"] != "N/A") | (
            output["Original Guarantee (Years)"] == 0), "Guarantee Check"] = "Correct"
    ######### END ############################
    ######### SPOUSE NAME CHECK #############
    output["Spouse Name"].fillna("N/A", inplace=True)
    output.loc[(output["Spouse Name"] == "N/A") & (output["Marital Status"] == "Yes"), "Spouse Name Check"] = "No"
    output.loc[(output["Spouse Name"] != "N/A") | (output["Marital Status"] == "No"), "Spouse Name Check"] = "Yes"
    ########################################
    ########### BENEFICIARY CHECK ############
    output["Beneficiary Name Check"].fillna("N/A", inplace=True)
    output.loc[
        (output["Beneficiary Name"] == "N/A") & (output["Status"] == "Beneficiary"), "Beneficiary Name Check"] = "No"
    output.loc[
        (output["Beneficiary Name"] != "N/A") | (output["Status"] != "Beneficiary"), "Beneficiary Name Check"] = "Yes"
    ##########################################
    output_name = "Review " + filename
    output.to_excel(output_name)


def gender_anamoly(gender):
    """ Returns whether the gender is binary or missing."""
    if gender in ["F", "M", "1", "2", 1, 2]:
        return 'Correct'
    elif gender == "N/A":
        return "Missing gender"
    else:
        return "Incorrect gender code"


def check_pc(code):
    """ Returns whether the postal code is valid, having the form XYX XYX
    where X represents alphabetical value and Y represents numerical value."""
    ignore_ws = code.replace(" ", "")
    if len(ignore_ws) != 6:
        return "Incorrect code"
    elif not ignore_ws[0].isalpha():
        return "Incorrect code"
    elif not ignore_ws[1].isnumeric():
        return "Incorrect code"
    elif not ignore_ws[2].isalpha():
        return "Incorrect code"
    elif not ignore_ws[3].isnumeric():
        return "Incorrect code"
    elif not ignore_ws[4].isalpha():
        return "Incorrect code"
    elif not ignore_ws[5].isnumeric():
        return "Incorrect code"
    return "Correct code"


if __name__ == '__main__':
    process_data("Client A.xlsx")
    # process_data("Client B.xlsx")
