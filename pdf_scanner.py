import os
import re
import lexnlp.extract.en.money
import pandas as pd

from PyPDF2 import PdfReader
from nltk import tokenize

# Constants
ENVIROMENTAL_KEYWORDS = ["Material", "Energy", "Water", "Biodiversity", "Emission", "Effluent", "Waste", "Climate",
                         "Paper", "Compliance", "Transport", "Environment", "Environmental initiatives",
                         "Environmental assessment", "Environmental operations"]

SOCIAL_KEYWORDS = ["Employment", "Labor management", "Employee", "Health", "Safety", "Training", "Education",
                   "Diversity", "Woman management", "Equal opportunity", "Equal remuneration", "Employee grievance",
                   "Discrimination", "Freedom of association", "Collective bargaining", "Child labor", "Forced labor", "Compulsory labor",
                   "Security", "Indigenous rights", "Human rights", "Public policy", "Corruption", "Anti-competitive behaviour",
                   "Compliance", "Society", "Impact society", "Local communities", "Labelling", "Marketing Communication",
                   "Customer privacy", "Customer compliance", "Economic impact", "Procurement practices", "Pay Ratio",
                   "Employee turnover", "Temporary worker", "Injury rate", "Labour"]

GOVERNANCE_KEYWORDS = ["Board Diversity", "Board Independence", "Incentives", "Collective Bargaining", "Code of Conduct",
                       "Ethics", "Corruption", "Data Privacy", "ESG Reporting", "Disclosure practices", "External Assurance",
                       "Data Protection", "Fair remuneration", "Independent Director", "Board of Directors", "Board Meeting",
                       "Committee"]


# test
PDF_PAGES_TO_INCLUDE = [
    {"file_name": "LPLA US.pdf", "page_start": 33, "page_end":54},
    {"file_name": "LOW US.pdf", "page_start": 5, "page_end":10},
]

OUTPUT_EXCEL_FILE_NAME = "!output.xlsx"
LOG_FILE_NAME = "!logs.txt"


def extract_pdf(file, file_params):

    reader = PdfReader(file)

    extracted_pages = []

    if file_params is None:
        for page in reader.pages:
            page_data = page.extract_text()
            extracted_pages.append(page_data)
    else:
        for index, page in enumerate(reader.pages):
            if index+1 < file_params["page_start"] or index >= file_params["page_end"]:
                continue
            else:
                page_data = page.extract_text()
                extracted_pages.append(page_data)

    tokenized_page_list = []
    for page in extracted_pages:
        tokenized_page = tokenize.sent_tokenize(page)
        tokenized_page_list.append(tokenized_page)


    for index, page in enumerate(tokenized_page_list):
        tokenized_page_list[index] = [x.replace('\n',' ').strip() for x in page]

    file_sentences_list = []
    for page in tokenized_page_list:
        for sentence in page:
            file_sentences_list.append(sentence)

    return file_sentences_list

def if_string_has_number(input_string):

    return any(char.isdigit() for char in input_string)

def if_string_has_currency(input_string):

    # input_string = "This is a test sentence Employment 50.45$sdfasdf"#vb
    input_string =  re.sub('([+-]?(?=\.\d|\d)(?:\d+)?(?:\.?\d*))(?:[Ee]([+-]?\d+))?', r' \1 ', input_string)            #re.sub("[A-Za-z]+", lambda ele: " " + ele[0] + " ", input_string)

    # for currency_name_shortened in MAJOR_CURRENCIES:
    #     currency_symbol = CurrencySymbols.get_symbol(currency_name_shortened)

    currency_found = False
    found_currency_data = list(lexnlp.extract.en.money.get_money(input_string))

    if len(found_currency_data) != 0:
        currency_found = True

    return currency_found

if __name__ == '__main__':

    with open(LOG_FILE_NAME, 'w', encoding="UTF-8") as f:
        f.write("PDF Scanner logs\n\n")

    if os.path.exists(OUTPUT_EXCEL_FILE_NAME):
        try:
            os.rename(OUTPUT_EXCEL_FILE_NAME, OUTPUT_EXCEL_FILE_NAME)
        except OSError as e:
            print('Access-error on file "' + OUTPUT_EXCEL_FILE_NAME + '"! \n' + str(e))
            raise Exception(f"Close {OUTPUT_EXCEL_FILE_NAME} and try again")



    #Create list of .pdf files in folder where python script is located
    files = [file for file in os.listdir('.') if os.path.isfile(file)]
    files = [file for file in files if file.endswith(".pdf")]

    output_list = []
    log_list = []
    error_counter = 0

    print(f"Job Starting: {len(files)} PDF files found\n")
    for file in files:

        file_params = None
        for idx, file_to_find in enumerate(PDF_PAGES_TO_INCLUDE):
            if file_to_find["file_name"] == file:
                file_params = file_to_find
                break



        log_string = f"Scanning {file} ... "
        print(log_string, end="")

        try:
            extracted_sentences = extract_pdf(file, file_params)
            sentences_count = len(extracted_sentences)

            sentences_list_with_ratings = []

            for sentence in extracted_sentences:

                item = {"E": None, "S": None, "G": None, "sentence":sentence}
                sentences_list_with_ratings.append(item)

            for item in sentences_list_with_ratings:
                sentence = item["sentence"]

                keywords_environmental_in_sentence = False
                keywords_social_in_sentence = False
                keywords_governance_in_sentence = False

                for keyword in ENVIROMENTAL_KEYWORDS:
                    if re.search(rf"\b{keyword.lower()}\b", sentence.lower()):
                        keywords_environmental_in_sentence = True
                        break

                for keyword in SOCIAL_KEYWORDS:
                    if re.search(rf"\b{keyword.lower()}\b", sentence.lower()):
                        keywords_social_in_sentence = True
                        break

                for keyword in GOVERNANCE_KEYWORDS:
                    if re.search(rf"\b{keyword.lower()}\b", sentence.lower()):
                        keywords_governance_in_sentence = True
                        break

                if keywords_environmental_in_sentence == False:
                    item["E"] = 0

                if keywords_social_in_sentence == False:
                    item["S"] = 0

                if keywords_governance_in_sentence == False:
                    item["G"] = 0


                #if keyword accurance in sentance occured:
                string_has_number = False
                string_has_currency = False

                string_has_number = if_string_has_number(sentence)
                string_has_currency = if_string_has_currency(sentence)

                #tweak item rating number based on has number or has currency
                if keywords_environmental_in_sentence == True:
                    item["E"] = 1

                    if string_has_number:
                        item["E"] = 2

                    if string_has_number and string_has_currency:
                        item["E"] = 3

                if keywords_social_in_sentence == True:
                    item["S"] = 1

                    if string_has_number:
                        item["S"] = 2

                    if string_has_number and string_has_currency:
                        item["S"] = 3

                if  keywords_governance_in_sentence == True:
                    item["G"] = 1

                    if string_has_number:
                        item["G"] = 2

                    if string_has_number and string_has_currency:
                        item["G"] = 3

            e0_total = 0
            e1_total = 0
            e2_total = 0
            e3_total = 0

            s0_total = 0
            s1_total = 0
            s2_total = 0
            s3_total = 0

            g0_total = 0
            g1_total = 0
            g2_total = 0
            g3_total = 0

            for sentence in sentences_list_with_ratings:

                e0_total += 1 if sentence["E"] == 0 else 0
                e1_total += 1 if sentence["E"] == 1 else 0
                e2_total += 1 if sentence["E"] == 2 else 0
                e3_total += 1 if sentence["E"] == 3 else 0

                s0_total += 1 if sentence["S"] == 0 else 0
                s1_total += 1 if sentence["S"] == 1 else 0
                s2_total += 1 if sentence["S"] == 2 else 0
                s3_total += 1 if sentence["S"] == 3 else 0

                g0_total += 1 if sentence["G"] == 0 else 0
                g1_total += 1 if sentence["G"] == 1 else 0
                g2_total += 1 if sentence["G"] == 2 else 0
                g3_total += 1 if sentence["G"] == 3 else 0

            output_row = {
                "name_of_file": os.path.splitext(file)[0],
                "total_sentences": sentences_count,
                "E0": e0_total,
                "E1": e1_total,
                "E2": e2_total,
                "E3": e3_total,
                "S0": s0_total,
                "S1": s1_total,
                "S2": s2_total,
                "S3": s3_total,
                "G0": g0_total,
                "G1": g1_total,
                "G2": g2_total,
                "G3": g3_total
                }

            output_list.append(output_row)

            log_string += "success"
            print("success")

            with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
                f.write(log_string + "\n")
        except:
            log_string += "error"
            print("error")

            error_counter += 1

            with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
                f.write(log_string + "\n")

    #end of files loop
    try:
        df = pd.DataFrame(output_list, columns=["name_of_file", "total_sentences",
                                                "E0","E1","E2","E3",
                                                "S0","S1","S2","S3",
                                                "G0","G1","G2","G3"])

        df.to_excel(OUTPUT_EXCEL_FILE_NAME, index=False)
    except:

        with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
            f.write("Error creating result excel file" + "\n")

        print("Error creating result excel file")
        raise Exception("Error creating result excel file")

    with open(LOG_FILE_NAME, 'a', encoding="UTF-8") as f:
        f.write(f"\nJob finished: files processed={len(files)}, errors={error_counter} (ctrl+f search for 'error')")

    print(f"\nJob finished: files processed={len(files)}, errors={error_counter}")




