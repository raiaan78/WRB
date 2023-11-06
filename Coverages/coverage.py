import itertools
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict
import pandas as pd
import openpyxl
from datetime import date

class ExcelLoaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title('ProductModel WRB Accelerator')
        
        # Dataframes for each file type
        self.Coverages = {}
        self.Forms = {}
        self.Inference = {}
        self.QRG_inference = {}
        self.Transaction_types = {}
        self.Limits = {}
        self.QRG_forms = {}
        self.SBT_model = {}
        self.SBT_model_covterms = {}
        self.SBT_model_covterm_options = {}
        self.SBT_model_covterm_states = {}
        self.SBT_model_covterm_options_states = {}
        self.Exclusions = {}
        self.Prod_Coverages = {}
        self.today = date.today()
        self.files_used = ""

        #set line of business
        self.lob = ""

        #Path for template file
        self.template = ""

        # Filepaths for display
        self.loaded_files = []

        # Buttons for each file type

        self.sbt_extract_btn = tk.Button(self.root, text='Load SBT Product Model Extract', command=lambda: self.load_file('SBT_extract'))
        self.sbt_extract_btn.pack(pady=10)

        self.production_btn = tk.Button(self.root, text='Load PROD Coverage File', command=lambda: self.load_file('prod_coverage'))
        self.production_btn.pack(pady=10)

        self.coverage_btn = tk.Button(self.root, text='Load CPU Coverage File', command=lambda: self.load_file('coverage'))
        self.coverage_btn.pack(pady=10)

        self.forms_btn = tk.Button(self.root, text='Load Forms File', command=lambda: self.load_file('forms'))
        self.forms_btn.pack(pady=10)

        self.inference_btn = tk.Button(self.root, text='Load Form Inference Steps File', command=lambda: self.load_file('inference'))
        self.inference_btn.pack(pady=10)

        self.qrg_btn = tk.Button(self.root, text='Load Forms QRG File', command=lambda: self.load_file('QRG'))
        self.qrg_btn.pack(pady=10)

        self.coverage_exclusions_btn = tk.Button(self.root, text='Load Coverage Exclusions File', command=lambda: self.load_file('coverage_exclusions'))
        self.coverage_exclusions_btn.pack(pady=10)

        self.covterm_options_btn = tk.Button(self.root, text='Load Limit Deductible File', command=lambda: self.load_file('covterm_options'))
        self.covterm_options_btn.pack(pady=10)

        self.input_template_btn = tk.Button(self.root, text='Load Template File', command=lambda: self.load_file('input_template'))
        self.input_template_btn.pack(pady=10)

        options = ["GL","CP","CA","IM"]
        self.clicked = tk.StringVar(self.root)
        self.clicked.set("Select your line of business")
        self.option = tk.OptionMenu(self.root, self.clicked, *options)
        self.option.pack()

        self.loaded_label = tk.Label(self.root, text='')
        self.loaded_label.pack(pady=20)

        self.process_btn = tk.Button(self.root, text='Process Files', command=self.process_files, state=tk.DISABLED)
        self.process_btn.pack(pady=20)

    def set_lob(self):
        self.lob = self.clicked.get()

    def generate_text(self, group):
        logic_texts = []
        step = 1

        for _, row in group.iterrows():
            step_logic = f"{step} -"
            if pd.isna(row['GOTO_STEP_ON_TRUE']) and pd.isna(row['GOTO_STEP_ON_FALSE']):
                step_logic += f" {row['STEP_NAME']}"
            if pd.notna(row['GOTO_STEP_ON_TRUE']):
                step_logic += f" If {row['STEP_NAME']}, then go to step {row['GOTO_STEP_ON_TRUE']}."
            if pd.notna(row['GOTO_STEP_ON_FALSE']):
                step_logic += f" If not, then go to step {row['GOTO_STEP_ON_FALSE']}."

            logic_texts.append(step_logic)
            step+=1

        return "\n".join(logic_texts)

    def load_file(self, file_type):
        filepath = filedialog.askopenfilename(title=f'Select {file_type} Excel File', filetypes=(('Excel Files', '*.xls;*.xlsx;*.xlsm'), ('All Files', '*.*')))
        if not filepath:
            return

        filename = os.path.basename(filepath)

        if file_type == 'coverage' and "CPU" in filename:
            self.Coverages = pd.read_excel(io=filepath, usecols = "A, C:G, J, S:T, X, Y, AM, BA")
            self.Coverages = self.Coverages.dropna(subset=['ENTITY_C'])
            self.coverage_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'forms' and "Forms To Coverages" in filename:
            self.Forms = pd.read_excel(io=filepath, usecols = "A:C, F, H:K, AB")
            self.Forms = self.Forms.dropna(subset=['COVERAGE_CODE'])
            self.forms_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'inference' and "Steps" in filename:
            self.Inference = pd.read_excel(io=filepath, usecols = "B, D:E, T:U")
            self.Inference = self.Inference.groupby('ROLL_ON_CND3_CODE').apply(self.generate_text).to_dict()
            self.inference_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'QRG' and "QRG" in filename:
            self.Transaction_types = pd.read_excel(io=filepath, usecols = "B, H, I")
            self.Transaction_types = self.Transaction_types[~self.Transaction_types['PROGRAM_NAME'].str.contains("FPP")]
            self.Transaction_types.drop_duplicates(inplace=True)

            self.QRG_forms = pd.read_excel(io=filepath, usecols = "B:D, H:J")
            self.QRG_forms['Form Edition'] = self.QRG_forms['Form Edition'].dt.strftime('%m/%y')
            self.QRG_forms = self.QRG_forms.groupby(['Form Number', 'Form Title', 'Form Edition','PROGRAM_NAME','RENEWAL_ACTION_C'])['STATE_CODE'].apply(lambda x: x.values.tolist()).reset_index()

            self.QRG_transactions = self.QRG_forms[['Form Number', 'Form Title', 'Form Edition', 'RENEWAL_ACTION_C']]
            self.QRG_transactions['RENEWAL_ACTION_C'] = self.QRG_transactions['RENEWAL_ACTION_C'].apply(lambda x: x.rstrip())
            self.QRG_transactions = self.QRG_transactions.set_index(['Form Number', 'Form Title', 'Form Edition']).to_dict()['RENEWAL_ACTION_C']
            
            result = {}  
            for name, group in self.QRG_forms.groupby(['Form Number', 'Form Title', 'Form Edition']):  
                result[name] = {}  
                for program, state in group.groupby('PROGRAM_NAME')['STATE_CODE']:  
                    result[name][program] = state.tolist()
                    result[name][program] = list(itertools.chain.from_iterable(result[name][program]))   

            self.QRG_forms = result

            self.QRG_inference = pd.read_excel(io=filepath, usecols = "B:D, X")
            self.QRG_inference['Form Edition'] = self.QRG_inference['Form Edition'].dt.strftime('%m/%y')
            self.QRG_inference = self.QRG_inference.dropna(subset=['ROLL_ON_CND3_CODE'])
            self.QRG_inference = self.QRG_inference.set_index(['Form Number', 'Form Title', 'Form Edition']).to_dict()['ROLL_ON_CND3_CODE']

            self.qrg_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'covterm_options' and "Limit" in filename:
            self.Limits = pd.read_excel(io=filepath, usecols = "A, B:D, E:I, L, R")
            self.Limits = self.Limits[~self.Limits['PROGRAM_NAME'].str.contains("FPP")]
            self.covterm_options_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'SBT_extract' and "ProductModelExport" in filename:
            self.SBT_model = pd.read_excel(io=filepath, sheet_name = "Clause", usecols = "A:C, E:F, I")
            #Parse the SBT model since multiple form IDs are within one cell in some cases
            self.SBT_model = self.SBT_model.assign(Form_ID = self.SBT_model['Form(s)'].str.split(r'\n')).explode('Form(s)')
            self.SBT_model = self.SBT_model.explode('Form_ID')
            self.SBT_model = self.SBT_model[["ClausePatternCode", "Description", "Type", "Existence", "Category", "Form_ID"]]
            self.SBT_model.drop_duplicates(inplace=True)

            self.SBT_model_states = pd.read_excel(io=filepath, sheet_name = "Clause Availability", usecols = "A, C")
            self.SBT_model_states = self.SBT_model_states.dropna()
            self.SBT_model_states = self.SBT_model_states.groupby("ClausePatternCode")["Jurisdiction"].apply(lambda x: x.values.tolist()).to_dict()
            
            self.SBT_model_covterms = pd.read_excel(io=filepath, sheet_name = "CovTerms", usecols = "A:C, F, H")
            self.SBT_model_covterms = self.SBT_model_covterms.groupby("ClausePatternCode")[["CovTermPatternCode","CovTerm Description","Required", "Default"]].apply(lambda x: x.values.tolist()).to_dict()

            self.SBT_model_covterm_options = pd.read_excel(io=filepath, sheet_name = "Options", usecols = "A:B, F")
            self.SBT_model_covterm_options = self.SBT_model_covterm_options.groupby(["ClausePatternCode", "CovTermPatternCode"])["Value"].apply(lambda x: x.values.tolist()).to_dict()

            self.SBT_model_covterm_states = pd.read_excel(io=filepath, sheet_name = "CovTerm Availability", usecols = "A:B, D")
            self.SBT_model_covterm_states = self.SBT_model_covterm_states.dropna()
            self.SBT_model_covterm_states = self.SBT_model_covterm_states.groupby(["ClausePatternCode", "CovTermPatternCode"])["Jurisdiction"].apply(lambda x: x.values.tolist()).to_dict()
             
            self.SBT_model_covterm_options_states = pd.read_excel(io=filepath, sheet_name = "Option Availability", usecols = "A:B, D")
            self.SBT_model_covterm_options_states = self.SBT_model_covterm_options_states.dropna()
            self.SBT_model_covterm_options_states = self.SBT_model_covterm_options_states.groupby(["ClausePatternCode", "CovTermPatternCode"])["Jurisdiction"].apply(lambda x: x.values.tolist()).to_dict()
            
            self.sbt_extract_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'coverage_exclusions' and "Exclusion" in filename:
            self.Exclusions = pd.read_excel(io=filepath, usecols = "A, C:D")
            self.coverage_exclusions_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'input_template' and "Product Model" in filename:
            self.template = filepath
            self.input_template_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'prod_coverage' and "PROD".casefold() in filename.casefold():
            self.Prod_Coverages = pd.read_excel(io=filepath, usecols="F, G, AS")
            self.Prod_Coverages['COVERAGE_CODE'] = self.Prod_Coverages['COVERAGE_CODE'].apply(lambda x: x.rstrip())
            self.Prod_Coverages = self.Prod_Coverages.dropna(subset=['ENTITY_C'])
            self.Prod_Coverages['ENTITY_C'] = self.Prod_Coverages['ENTITY_C'].apply(lambda x: x.rstrip())
            self.Prod_Coverages.drop_duplicates(inplace=True)
            self.production_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        else:
            messagebox.showerror("Error", f"Invalid file selected for {file_type}. Please select the correct file.")

        # Update the loaded files label
        self.loaded_label.config(text=', '.join(self.loaded_files))

        # Enable the process button if all files are loaded
        if len(self.loaded_files) == 9:  # Assuming you have 9 files to load
            self.process_btn.config(state=tk.NORMAL)

    def process_files(self):
        self.set_lob()

        def print_qrg_forms(sheet, row, form):
            form_number = form[0]
            form_name = form[1]
            form_edition = form[2].replace('/'," ")
            form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

            for program in self.QRG_forms[form]:
                #Updated By
                cell0 = "C" + str(row)
                sheet[cell0] = "Automation Script"

                #ISO/Proprietary
                cell1 = "D" + str(row)

                if self.lob == "GL":
                    if (form_pattern[:2] == "CG" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell1] = "Proprietary"
                    else:
                        sheet[cell1] = "ISO"

                if self.lob == "CP":
                    if (form_pattern[:2] == "CP" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell1] = "Proprietary"
                    else:
                        sheet[cell1] = "ISO"

                if self.lob == "CA":
                    if (form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell1] = "Proprietary"
                    else:
                        sheet[cell1] = "ISO"

                if self.lob == "IM":
                    if (form_pattern[:2] == "IM" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell1] = "Proprietary"
                    else:
                        sheet[cell1] = "ISO"

                cell2 = "E" + str(row)
                sheet[cell2] = form_pattern

                cell3 = "F" + str(row)
                sheet[cell3] = form_number

                cell4 = "G" + str(row)
                sheet[cell4] = form_edition

                cell5 = "H" + str(row)
                sheet[cell5] = form_name

                cell6 = "I" + str(row)

                state_set = set(self.QRG_forms[form][program])

                if len(state_set) == len(US_states) or "A1" in state_set:
                    sheet[cell6] = "All States"
                elif len(state_set) <= 10:
                    sheet[cell6] = ','.join(state_set)
                else:
                    difference = US_states.difference(state_set)
                    sheet[cell6] = "All states except: " + ','.join(difference)

                cell7 = "J" + str(row)
                sheet[cell7] = program

                #Populate Transaction Types
                cell8 = "O" + str(row)

                if form in self.QRG_transactions:
                    if self.QRG_transactions[form] == "RETAIN":
                        sheet[cell8] = "Submission, Policy, Change, Rewrite, Rewrite New Account, Renewal"
                    else:
                        sheet[cell8] = "Submission, Policy, Change, Rewrite, Rewrite New Account"

                if form in self.QRG_inference and self.QRG_inference[form] in self.Inference:
                    if sheet != product_model["State Amendatory Endorsements"] and sheet != product_model["Common Forms"]:
                        cell28 = "AC" + str(row)
                    else:
                        cell28 = "N" + str(row)
                    
                    sheet[cell28] = self.Inference[self.QRG_inference[form]]
                
                row+=1

            return row

        def print_qrg_sbt_forms(sheet, row, form):
            form_number = form[0]
            form_name = form[1]
            form_edition = form[2].replace('/'," ")
            form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

            category_idx = 0
            if form_pattern in sbt_form_to_category:
                category_idx = len(sbt_form_to_category[form_pattern]) - 1
            
            for program in self.QRG_forms[form]:
                while category_idx >= 0:
                    if self.lob == "GL":
                        cell15 = "AH" + str(row)
                        cell16 = "AI" + str(row)
                        cell17 = "AJ" + str(row)
                        cell18 = "AK" + str(row)
                    else:
                        cell15 = "V" + str(row)
                        cell16 = "W" + str(row)
                        cell17 = "X" + str(row)
                        cell18= "Y" + str(row)

                    sheet[cell15] = form_pattern
                    sheet[cell16] = form_number
                    sheet[cell17] = form_edition
                    sheet[cell18] = form_name

                    #Populate SBT/OOTB
                    cell19 = "H" + str(row)
                    sheet[cell19] = "SBT"

                    #Change coverage name to whatever is in the SBT model
                    sheet["I" + str(row)] = sbt[form_pattern]

                    #Change existence of coverage to whatever is in SBT model
                    sheet["N" + str(row)] = sbt_eoc[form_pattern]
                        
                    #ISO/Proprietary
                    cell20 = "J" + str(row)

                    if self.lob == "GL":
                        if (form_pattern[:2] == "CG" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                            sheet[cell20] = "Proprietary"
                        else:
                            sheet[cell20] = "ISO"

                    if self.lob == "CP":
                        if (form_pattern[:2] == "CP" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                            sheet[cell20] = "Proprietary"
                        else:
                            sheet[cell20] = "ISO"

                    if self.lob == "CA":
                        if (form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                            sheet[cell20] = "Proprietary"
                        else:
                            sheet[cell20] = "ISO"

                    if self.lob == "IM":
                        if (form_pattern[:2] == "IM" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                            sheet[cell20] = "Proprietary"
                        else:
                            sheet[cell20] = "ISO"

                    #Populate form states
                    if self.lob == "GL":
                        cell21 = "AL" + str(row)
                    else:
                        cell21 = "Z" + str(row)

                    state_set = set(self.QRG_forms[form][program])

                    if len(state_set) == len(US_states) or "A1" in state_set:
                        sheet[cell21] = "All States"
                    elif len(state_set) <= 10:
                        sheet[cell21] = ','.join(state_set)
                    else:
                        difference = US_states.difference(state_set)
                        sheet[cell21] = "All states except: " + ','.join(difference)

                    #Populate Transaction Types
                    if self.lob == "GL":
                        cell22 = "AP" + str(row)
                    else:
                        cell22 = "AD" + str(row)

                    if self.QRG_transactions[form] == "RETAIN":
                        sheet[cell22] = "Submission, Policy, Change, Rewrite, Rewrite New Account, Renewal"
                    else:
                        sheet[cell22] = "Submission, Policy, Change, Rewrite, Rewrite New Account"

                    #Populate Category
                    cell23 = "M" + str(row)
                    if form_pattern in sbt_form_to_category:
                        category_value = sbt_form_to_category[form_pattern][category_idx][3:]
                        
                        if "AddlGrp" in category_value:
                            idx = category_value.find("AddlGrp")
                            category_value = category_value[:idx] + " - Additional Coverage"
                        elif "CondGrp" in category_value:
                            idx = category_value.find("CondGrp")
                            category_value = category_value[:idx] + " - Conditions"
                        elif "ExclGrp" in category_value:
                            idx = category_value.find("ExclGrp")
                            category_value = category_value[:idx] + " - Exclusions"
                        elif "StdGrp" in category_value:
                            idx = category_value.find("StdGrp")
                            category_value = category_value[:idx] + " - Coverages"
                        elif "BlanketGrp" in category_value:
                            idx = category_value.find("BlanketGrp")
                            category_value = category_value[:idx] + " - Blanket Coverages"
                        elif "AddlInsdGrp" in category_value:
                            idx = category_value.find("AddlInsdGrp")
                            category_value = category_value[:idx] + " - Additional Insured"
                        else:
                            pass

                        sheet[cell23] = category_value

                    cell24 = "C" + str(row)
                    sheet[cell24] = self.today

                    cell25 = "D" + str(row)
                    sheet[cell25] = ', '.join(self.loaded_files[i] for i in [0, 5])

                    cell26 = "E" + str(row)
                    sheet[cell26] = "Automation Script"

                    cell27 = "L" + str(row)
                    sheet[cell27] = program

                    if form in self.QRG_inference and self.QRG_inference[form] in self.Inference:
                        if sheet != product_model["State Amendatory Endorsements"] and sheet != product_model["Common Forms"]:
                            cell28 = "AC" + str(row)
                        else:
                            cell28 = "N" + str(row)
                        
                        sheet[cell28] = self.Inference[self.QRG_inference[form]]

                    category_idx-=1
                    row+=1
            
            return row

        def print_amendatory_coverages(sheet, row):
            #Print updated by
            cell10 = "C" + str(row)
            sheet[cell10] = "Automation Script"

            #Print OU and UW
            cell8 = "K" + str(row)
            cell9 = "L" + str(row)

            #Scenario 4
            if cov_code[0] not in ou_and_uw_exclusions:
                sheet[cell8] = "All"
                sheet[cell9] = "All"
            else:
                null_operating_unit = 0
                null_underwriting_company = 0
                ou_exception = set()
                uw_exception = set()

                for pair in ou_and_uw_exclusions[cov_code[0]]:
                    if not pd.isna(pair[0]) and not pd.isna(pair[1]):
                        if pair[0].rstrip() in ou_abbreviations:
                            ou_exception.add(ou_abbreviations[pair[0].rstrip()])
                            uw_exception.add(pair[1] + "(" + ou_abbreviations[pair[0].rstrip()] + ")")
                        continue
                    
                    if pd.isna(pair[0]):
                        null_operating_unit+=1
                    else:
                        if pair[0].rstrip() in ou_abbreviations:
                            ou_exception.add(ou_abbreviations[pair[0].rstrip()])

                    if pd.isna(pair[1]):
                        null_underwriting_company+=1
                    else:
                        uw_exception.add(pair[1])

                #Scenario 2
                if null_operating_unit == len(ou_and_uw_exclusions[cov_code[0]]) and null_underwriting_company == 0:
                    sheet[cell8] = "All"
                    sheet[cell9] = "All except " + ', '.join(uw_exception)
                #Scenario 3
                elif null_operating_unit == 0 and null_underwriting_company == len(ou_and_uw_exclusions[cov_code[0]]):
                    sheet[cell8] = "All except " + ', '.join(ou_exception)
                    sheet[cell9] = "All"
                #Scenario 1
                else:
                    sheet[cell8] = "All except " + ', '.join(ou_exception)
                    sheet[cell9] = "All except " + ', '.join(uw_exception)

        def print_amendatory_forms(sheet, row, index):
            form_number = state_amendatory[cov_code][index][0]
            form_name = state_amendatory[cov_code][index][1]
            form_edition = state_amendatory[cov_code][index][2].replace('/'," ")
            form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

            #ISO/Proprietary
            cell0 = "D" + str(row)

            if self.lob == "GL":
                if (form_pattern[:2] == "CG" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell0] = "Proprietary"
                else:
                    sheet[cell0] = "ISO"

            if self.lob == "CP":
                if (form_pattern[:2] == "CP" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell0] = "Proprietary"
                else:
                    sheet[cell0] = "ISO"

            if self.lob == "CA":
                if (form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell0] = "Proprietary"
                else:
                    sheet[cell0] = "ISO"

            if self.lob == "IM":
                if (form_pattern[:2] == "IM" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell0] = "Proprietary"
                else:
                    sheet[cell0] = "ISO"

            cell1 = "E" + str(row)
            sheet[cell1] = form_pattern

            cell2 = "F" + str(row)
            sheet[cell2] = form_number

            cell3 = "G" + str(row)
            sheet[cell3] = form_edition

            cell4 = "H" + str(row)
            sheet[cell4] = form_name

            cell5 = "I" + str(row)

            state_set = set(form_states[cov_code[0], cov_code[1], cov_code[2], form_number, form_edition.replace(" ","/")])

            if len(state_set) == len(US_states) or "A1" in state_set:
                sheet[cell5] = "All States"
            elif len(state_set) <= 10:
                sheet[cell5] = ','.join(state_set)
            else:
                difference = US_states.difference(state_set)
                sheet[cell5] = "All states except: " + ','.join(difference)

            #Populate Transaction Types
            cell6 = "O" + str(row)

            if form_number in transactions:
                if transactions[form_number] == "RETAIN":
                    sheet[cell6] = "Submission, Policy, Change, Rewrite, Rewrite New Account, Renewal"
                else:
                    sheet[cell6] = "Submission, Policy, Change, Rewrite, Rewrite New Account"

            #Populate Inference Logic
            if (cov_code[0], cov_code[1], cov_code[2], state_amendatory[cov_code][index][0], state_amendatory[cov_code][index][2]) in inference and inference[cov_code[0], cov_code[1], cov_code[2], state_amendatory[cov_code][index][0], state_amendatory[cov_code][index][2]] in self.Inference:
                cell7 = "N" + str(row)
                sheet[cell7] = self.Inference[inference[cov_code[0], cov_code[1], cov_code[2], state_amendatory[cov_code][index][0], state_amendatory[cov_code][index][2]]]

            cell8 = "J" + str(row)
            sheet[cell8] = cov_code[1]

        def print_common_coverages(sheet, row):
            #Print last updated by
            cell7 = "C" + str(row)
            sheet[cell7] = "Automation Script"

            #Print OU and UW
            cell5 = "K" + str(row)
            cell6 = "L" + str(row)

            #Scenario 4
            if cov_code[0] not in ou_and_uw_exclusions:
                sheet[cell5] = "All"
                sheet[cell6] = "All"
            else:
                null_operating_unit = 0
                null_underwriting_company = 0
                ou_exception = set()
                uw_exception = set()

                for pair in ou_and_uw_exclusions[cov_code[0]]:
                    if not pd.isna(pair[0]) and not pd.isna(pair[1]):
                        if pair[0].rstrip() in ou_abbreviations:
                            ou_exception.add(ou_abbreviations[pair[0].rstrip()])
                            uw_exception.add(pair[1] + "(" + ou_abbreviations[pair[0].rstrip()] + ")")
                        continue
                    
                    if pd.isna(pair[0]):
                        null_operating_unit+=1
                    else:
                        if pair[0].rstrip() in ou_abbreviations:
                            ou_exception.add(ou_abbreviations[pair[0].rstrip()])

                    if pd.isna(pair[1]):
                        null_underwriting_company+=1
                    else:
                        uw_exception.add(pair[1])

                #Scenario 2
                if null_operating_unit == len(ou_and_uw_exclusions[cov_code[0]]) and null_underwriting_company == 0:
                    sheet[cell5] = "All"
                    sheet[cell6] = "All except " + ', '.join(uw_exception)
                #Scenario 3
                elif null_operating_unit == 0 and null_underwriting_company == len(ou_and_uw_exclusions[cov_code[0]]):
                    sheet[cell5] = "All except " + ', '.join(ou_exception)
                    sheet[cell6] = "All"
                #Scenario 1
                else:
                    sheet[cell5] = "All except " + ', '.join(ou_exception)
                    sheet[cell6] = "All except " + ', '.join(uw_exception)

        def print_common_forms(sheet, row, index):
            form_number = common_forms[cov_code][index][0]
            form_name = common_forms[cov_code][index][1]
            form_edition = common_forms[cov_code][index][2].replace('/'," ")
            form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

            cell0 = "E" + str(row)
            sheet[cell0] = form_pattern

            cell1 = "F" + str(row)
            sheet[cell1] = form_number

            cell2 = "G" + str(row)
            sheet[cell2] = form_edition

            cell3 = "H" + str(row)
            sheet[cell3] = form_name

            cell4 = "I" + str(row)

            state_set = set(form_states[cov_code[0], cov_code[1], cov_code[2], form_number, form_edition.replace(" ","/")])

            if len(state_set) == len(US_states) or "A1" in state_set:
                sheet[cell4] = "All States"
            elif len(state_set) <= 10:
                sheet[cell4] = ','.join(state_set)
            else:
                difference = US_states.difference(state_set)
                sheet[cell4] = "All states except: " + ','.join(difference)

            if form_number in transactions:
                cell5 = "O" + str(row)
                if transactions[form_number] == "RETAIN":
                    sheet[cell5] = "Submission, Policy, Change, Rewrite, Rewrite New Account, Renewal"
                else:
                    sheet[cell5] = "Submission, Policy, Change, Rewrite, Rewrite New Account"

            if (cov_code[0], cov_code[1], cov_code[2], common_forms[cov_code][index][0], common_forms[cov_code][index][2]) in inference and inference[cov_code[0], cov_code[1], cov_code[2], common_forms[cov_code][index][0], common_forms[cov_code][index][2]] in self.Inference:
                cell6 = "N" + str(row)
                sheet[cell6] = self.Inference[inference[cov_code[0], cov_code[1], cov_code[2], common_forms[cov_code][index][0], common_forms[cov_code][index][2]]]

            cell7 = "J" + str(row)
            sheet[cell7] = cov_code[1]

            #ISO/Proprietary
            cell8 = "D" + str(row)

            if self.lob == "GL":
                if (form_pattern[:2] == "CG" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell8] = "Proprietary"
                else:
                    sheet[cell8] = "ISO"

            if self.lob == "CP":
                if (form_pattern[:2] == "CP" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell8] = "Proprietary"
                else:
                    sheet[cell8] = "ISO"

            if self.lob == "CA":
                if (form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell8] = "Proprietary"
                else:
                    sheet[cell8] = "ISO"

            if self.lob == "IM":
                if (form_pattern[:2] == "IM" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell8] = "Proprietary"
                else:
                    sheet[cell8] = "ISO"

        def print_coverages(sheet, row):
            #Populate today's date
            cell0 = "C" + str(row)
            sheet[cell0] = self.today

            #Populate files used
            cell1 = "D" + str(row)
            sheet[cell1] = self.files_used

            #Population Last Updated By
            cell15 = "E" + str(row)
            sheet[cell15] = "Automation Script"

            #Populate production status
            cell2 = "F" + str(row)

            if cov_code[0] in self.Prod_Coverages['COVERAGE_CODE'].values and cov_code[1] in self.Prod_Coverages['PROGRAM_NAME'].values and cov_code[2] in self.Prod_Coverages['ENTITY_C'].values:
                sheet[cell2] = "Y"
            else:
                sheet[cell2] = "N"

            #Populate parent_id
            cell3 = "G" + str(row)
            sheet[cell3] = parent_id[cov_code]
            
            #Populate coverage name only if not written by SBT already
            cell4 = "I" + str(row)
            if sheet[cell4].value is None:
                sheet[cell4] = coverage[cov_code]

            #Populate coverage states
            cell5 = "K" + str(row)

            if len(cov_states[cov_code]) == len(US_states) or "A1" in cov_states[cov_code]:
                sheet[cell5] = "All States"
            elif len(cov_states[cov_code]) <= 10:
                sheet[cell5] = ','.join(cov_states[cov_code])
            else:
                difference = US_states.difference(cov_states[cov_code])
                sheet[cell5] = "All states except: " + ','.join(difference)

            #Populate Offering/Program
            cell6 = "L" + str(row)
            sheet[cell6] = cov_code[1]

            #Populate existence of coverage
            cell7 = "N" + str(row)
            if sheet[cell7].value is None:
                eoc = existence[cov_code]

                if eoc[0] == 'Y' and eoc[1] == 'N':
                    sheet[cell7] = "Required"
                elif eoc[0] == 'N' and eoc[1] == 'N':
                    sheet[cell7] = "Electable"
                else:
                    sheet[cell7] = "Suggested"

            #Populate Premium Bearing
            cell8 = "O" + str(row)
            sheet[cell8] = premium[cov_code]

            #Populate scheduled field
            cell9 = "P" + str(row)
            if parent_scheduled[cov_code] == "Y":
                sheet[cell9] = "Y"
            else:
                answer = False
                for child in parent_child[cov_code]:
                    if child_scheduled[child] == "Y":
                        answer = True
                        break

                if answer == True:
                    sheet[cell9] = "Y"
                else:
                    sheet[cell9] = "N"

            #Populate operating units and underwriting companies
            cell10 = "Q" + str(row)
            cell11 = "R" + str(row)

            #Scenario 4
            if cov_code[0] not in ou_and_uw_exclusions:
                sheet[cell10] = "All"
                sheet[cell11] = "All"
            else:
                null_operating_unit = 0
                null_underwriting_company = 0
                ou_exception = set()
                uw_exception = set()

                for pair in ou_and_uw_exclusions[cov_code[0]]:
                    if not pd.isna(pair[0]) and not pd.isna(pair[1]):
                        if pair[0].rstrip() in ou_abbreviations:
                            ou_exception.add(ou_abbreviations[pair[0].rstrip()])
                            uw_exception.add(pair[1] + "(" + ou_abbreviations[pair[0].rstrip()] + ")")
                        continue
                    
                    if pd.isna(pair[0]):
                        null_operating_unit+=1
                    else:
                        if pair[0].rstrip() in ou_abbreviations:
                            ou_exception.add(ou_abbreviations[pair[0].rstrip()])

                    if pd.isna(pair[1]):
                        null_underwriting_company+=1
                    else:
                        uw_exception.add(pair[1])

                #Scenario 2
                if null_operating_unit == len(ou_and_uw_exclusions[cov_code[0]]) and null_underwriting_company == 0:
                    sheet[cell10] = "All"
                    sheet[cell11] = "All except " + ', '.join(uw_exception)
                #Scenario 3
                elif null_operating_unit == 0 and null_underwriting_company == len(ou_and_uw_exclusions[cov_code[0]]):
                    sheet[cell10] = "All except " + ', '.join(ou_exception)
                    sheet[cell11] = "All"
                #Scenario 1
                else:
                    sheet[cell10] = "All except " + ', '.join(ou_exception)
                    sheet[cell11] = "All except " + ', '.join(uw_exception)

            #Populate ASOLB/Major Peril Code
            cell12 = "S" + str(row)
            
            if self.lob == "GL":
                sheet[cell12] = ','.join(major_peril[cov_code])
                #Populate Subline C items
                code = subline[cov_code]

                if code == '          ':
                    pass
                elif code == 334 or code == 336:
                    cell13 = "T" + str(row)
                    cell14 = "U" + str(row)
                    sheet[cell13] = "x"
                    sheet[cell14] = "x"
                elif code == 332:
                    cell13 = "V" + str(row)
                    cell14 = "W" + str(row)
                    sheet[cell13] = "x"
                    sheet[cell14] = "x"
                elif code == 317:
                    cell13 = "X" + str(row)
                    cell14 = "Y" + str(row)
                    sheet[cell13] = "x"
                    sheet[cell14] = "x"
                elif code == 325:
                    cell13 = "Z" + str(row)
                    cell14 = "AA" + str(row)
                    sheet[cell13] = "x"
                    sheet[cell14] = "x"
                elif code == 360:
                    cell13 = "AB" + str(row)
                    cell14 = "AC" + str(row)
                    sheet[cell13] = "x"
                    sheet[cell14] = "x"
                else:
                    sheet[cell12] = str(code) + "/" + sheet[cell12].value
            else:
                if subline[cov_code] != '          ' and subline[cov_code] != "nan" and major_peril[cov_code] != '          ' and major_peril[cov_code] != "nan":
                    sheet[cell12] = subline[cov_code] + "/" + major_peril[cov_code]
                
        def print_forms(sheet, row, index, coverage_type):
            #Populate form info
            if coverage_type == "General":
                form_number = parent_forms[cov_code][index][0]
                form_name = parent_forms[cov_code][index][1]
                form_edition = parent_forms[cov_code][index][2].replace('/'," ")
                form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

                #Populate Inference Logic
                if (cov_code[0], cov_code[1], cov_code[2], parent_forms[cov_code][index][0], parent_forms[cov_code][index][2]) in inference and inference[cov_code[0], cov_code[1], cov_code[2], parent_forms[cov_code][index][0], parent_forms[cov_code][index][2]] in self.Inference:
                    cell24 = "AC" + str(row)
                    sheet[cell24] = self.Inference[inference[cov_code[0], cov_code[1], cov_code[2], parent_forms[cov_code][index][0], parent_forms[cov_code][index][2]]]

            elif coverage_type == "Exclusion":
                form_number = exclusions[cov_code][index][0]
                form_name = exclusions[cov_code][index][1]
                form_edition = exclusions[cov_code][index][2].replace('/'," ")
                form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

                if (cov_code[0], cov_code[1], cov_code[2], exclusions[cov_code][index][0], exclusions[cov_code][index][2]) in inference and inference[cov_code[0], cov_code[1], cov_code[2], exclusions[cov_code][index][0], exclusions[cov_code][index][2]] in self.Inference:
                    cell24 = "AC" + str(row)
                    sheet[cell24] = self.Inference[inference[cov_code[0], cov_code[1], cov_code[2], exclusions[cov_code][index][0], exclusions[cov_code][index][2]]]

            else:
                form_number = conditions[cov_code][index][0]
                form_name = conditions[cov_code][index][1]
                form_edition = conditions[cov_code][index][2].replace('/'," ")
                form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

                if (cov_code[0], cov_code[1], cov_code[2], conditions[cov_code][index][0], conditions[cov_code][index][2]) in inference and inference[cov_code[0], cov_code[1], cov_code[2], conditions[cov_code][index][0], conditions[cov_code][index][2]] in self.Inference:
                    cell24 = "AC" + str(row)
                    sheet[cell24] = self.Inference[inference[cov_code[0], cov_code[1], cov_code[2], conditions[cov_code][index][0], conditions[cov_code][index][2]]]

            category_idx = 0
            if form_pattern in sbt_form_to_category:
                category_idx = len(sbt_form_to_category[form_pattern]) - 1
            
            while category_idx >= 0:
                if self.lob == "GL":
                    cell15 = "AH" + str(row)
                    cell16 = "AI" + str(row)
                    cell17 = "AJ" + str(row)
                    cell18 = "AK" + str(row)
                else:
                    cell15 = "V" + str(row)
                    cell16 = "W" + str(row)
                    cell17 = "X" + str(row)
                    cell18= "Y" + str(row)

                sheet[cell15] = form_pattern
                sheet[cell16] = form_number
                sheet[cell17] = form_edition
                sheet[cell18] = form_name

                #Populate SBT/OOTB
                cell19 = "H" + str(row)

                if form_pattern in sbt:
                    sheet[cell19] = "SBT"

                    #Change coverage name to whatever is in the SBT model
                    sheet["I" + str(row)] = sbt[form_pattern]

                    #Change existence of coverage to whatever is in SBT model
                    sheet["N" + str(row)] = sbt_eoc[form_pattern]
                    
                if self.lob == "GL" and form_pattern[:2] == "CG" and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell19] = "New"

                if self.lob == "CP" and form_pattern[:2] == "CP" and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell19] = "New"

                if self.lob == "CA" and (form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell19] = "New"

                if self.lob == "IM" and form_pattern[:2] == "IM" and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                    sheet[cell19] = "New"

                #ISO/Proprietary
                cell20 = "J" + str(row)

                if self.lob == "GL":
                    if (form_pattern[:2] == "CG" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell20] = "Proprietary"
                    else:
                        sheet[cell20] = "ISO"

                if self.lob == "CP":
                    if (form_pattern[:2] == "CP" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell20] = "Proprietary"
                    else:
                        sheet[cell20] = "ISO"

                if self.lob == "CA":
                    if (form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell20] = "Proprietary"
                    else:
                        sheet[cell20] = "ISO"

                if self.lob == "IM":
                    if (form_pattern[:2] == "IM" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                        sheet[cell20] = "Proprietary"
                    else:
                        sheet[cell20] = "ISO"

                #Populate form states
                if self.lob == "GL":
                    cell21 = "AL" + str(row)
                else:
                    cell21 = "Z" + str(row)

                state_set = {}
                
                if form_pattern in sbt:
                    if form_pattern in sbt_form_to_category:
                        clause = sbt_form_id_and_category_to_clause[(form_pattern, sbt_form_to_category[form_pattern][category_idx])]
                        if clause in self.SBT_model_states:
                            state_set = self.SBT_model_states[clause]
                    
                else:
                    state_set = set(form_states[cov_code[0], cov_code[1], cov_code[2], form_number, form_edition.replace(" ","/")])

                if len(state_set) == len(US_states) or "A1" in state_set:
                    sheet[cell21] = "All States"
                elif len(state_set) <= 10:
                    sheet[cell21] = ','.join(state_set)
                else:
                    difference = US_states.difference(state_set)
                    sheet[cell21] = "All states except: " + ','.join(difference)

                #Populate Transaction Types
                if self.lob == "GL":
                    cell22 = "AP" + str(row)
                else:
                    cell22 = "AD" + str(row)

                if form_number in transactions:
                    if transactions[form_number] == "RETAIN":
                        sheet[cell22] = "Submission, Policy, Change, Rewrite, Rewrite New Account, Renewal"
                    else:
                        sheet[cell22] = "Submission, Policy, Change, Rewrite, Rewrite New Account"

                #Populate Category
                cell23 = "M" + str(row)
                if form_pattern in sbt_form_to_category:
                    category_value = sbt_form_to_category[form_pattern][category_idx][3:]
                    
                    if "AddlGrp" in category_value:
                        idx = category_value.find("AddlGrp")
                        category_value = category_value[:idx] + " - Additional Coverage"
                    elif "CondGrp" in category_value:
                        idx = category_value.find("CondGrp")
                        category_value = category_value[:idx] + " - Conditions"
                    elif "ExclGrp" in category_value:
                        idx = category_value.find("ExclGrp")
                        category_value = category_value[:idx] + " - Exclusions"
                    elif "StdGrp" in category_value:
                        idx = category_value.find("StdGrp")
                        category_value = category_value[:idx] + " - Coverages"
                    elif "BlanketGrp" in category_value:
                        idx = category_value.find("BlanketGrp")
                        category_value = category_value[:idx] + " - Blanket Coverages"
                    elif "AddlInsdGrp" in category_value:
                        idx = category_value.find("AddlInsdGrp")
                        category_value = category_value[:idx] + " - Additional Insured"
                    else:
                        pass

                    sheet[cell23] = category_value
                
                else:
                    sheet[cell23] = cov_code[2]

                category_idx-=1
                row+=1
            
            return row

        def populate_covterm_options(option_list, covterm, coverage_code, coverage_terms_options_row, clause = ""):
            option_counter = 0

            while option_counter < len(option_list):
                #Populate current date
                current_date = "C" + str(coverage_terms_options_row)
                coverage_term_options_sheet[current_date] = self.today

                files_used2 = "D" + str(coverage_terms_options_row)
                coverage_term_options_sheet[files_used2] = self.loaded_files[7]

                last_updated2 = "E" + str(coverage_terms_options_row)
                coverage_term_options_sheet[last_updated2] = "Automation Script"

                sbt_or_new_2 = "F" + str(coverage_terms_options_row)
                coverage_term_options_sheet[sbt_or_new_2] = coverage_terms_sheet["F" + str(coverage_terms_row)].value

                covterm_options_parent_output = "G" + str(coverage_terms_options_row)
                coverage_term_options_sheet[covterm_options_parent_output] = coverage_terms_sheet["G" + str(coverage_terms_row)].value

                covterm_options_child_output = "H" + str(coverage_terms_options_row)
                coverage_term_options_sheet[covterm_options_child_output] = coverage_terms_sheet["H" + str(coverage_terms_row)].value

                covterm_program2 = "I" + str(coverage_terms_options_row)
                coverage_term_options_sheet[covterm_program2] = coverage_terms_sheet["I" + str(coverage_terms_row)].value

                covterm_category2 = "J" + str(coverage_terms_options_row)
                coverage_term_options_sheet[covterm_category2] = coverage_terms_sheet["J" + str(coverage_terms_row)].value

                option_name = "L" + str(coverage_terms_options_row)
                coverage_term_options_sheet[option_name] = option_list[option_counter]

                child_states2 = "M" + str(coverage_terms_options_row)

                if coverage_code in sbt_parent_coverages:
                    if (clause, covterm[0]) in self.SBT_model_covterm_options_states:
                        states2 = set(self.SBT_model_covterm_options_states[clause, covterm[0]])
                    
                        if len(states2) == len(US_states) or "A1" in states2:
                            coverage_term_options_sheet[child_states2] = "All States"
                        elif len(states2) <= 10:
                            coverage_term_options_sheet[child_states2] = ','.join(states2)
                        else:
                            difference = US_states.difference(states2)
                            coverage_term_options_sheet[child_states2] = "All states except: " + ','.join(difference)
                else:
                    coverage_term_options_sheet[child_states2] = coverage_terms_sheet["O" + str(coverage_terms_row)].value
                
                coverage_terms_options_row+=1
                option_counter+=1

            return coverage_terms_options_row

        def populate_sbt_covterms(items, coverage_code, sbt_form, coverage_terms_row, coverage_terms_options_row, clause):
            for covterm in items:
                option_list = {}

                #Populate date
                coverage_terms_sheet["C" + str(coverage_terms_row)] = self.today

                #Populate files used
                coverage_terms_sheet["D" + str(coverage_terms_row)] = self.loaded_files[7]

                #Populate last updated by
                coverage_terms_sheet["E" + str(coverage_terms_row)] = "Automation Script"

                #Populate coverage name from SBT
                coverage_terms_sheet["G" + str(coverage_terms_row)] = sbt[sbt_form]

                #Populate SBT
                coverage_terms_sheet["F" + str(coverage_terms_row)] = "SBT"

                #Populate covterm name from SBT
                coverage_terms_sheet["H" + str(coverage_terms_row)] = covterm[1]

                #Populate program
                coverage_terms_sheet["I" + str(coverage_terms_row)] = coverage_code[1]

                #Populate category
                if clause in sbt_clause_to_category:
                    category_value = sbt_clause_to_category[clause][3:]

                    if "AddlGrp" in category_value:
                        idx = category_value.find("AddlGrp")
                        category_value = category_value[:idx] + " - Additional Coverage"
                    elif "CondGrp" in category_value:
                        idx = category_value.find("CondGrp")
                        category_value = category_value[:idx] + " - Conditions"
                    elif "ExclGrp" in category_value:
                        idx = category_value.find("ExclGrp")
                        category_value = category_value[:idx] + " - Exclusions"
                    elif "StdGrp" in category_value:
                        idx = category_value.find("StdGrp")
                        category_value = category_value[:idx] + " - Coverages"
                    elif "BlanketGrp" in category_value:
                        idx = category_value.find("BlanketGrp")
                        category_value = category_value[:idx] + " - Blanket Coverages"
                    elif "AddlInsdGrp" in category_value:
                        idx = category_value.find("AddlInsdGrp")
                        category_value = category_value[:idx] + " - Additional Insured"
                    else:
                        pass

                    coverage_terms_sheet["J" + str(coverage_terms_row)] = category_value
                
                else:
                    coverage_terms_sheet["J" + str(coverage_terms_row)] = coverage_code[2]
                
                #Populate term type and value type
                coverage_terms_sheet["K" + str(coverage_terms_row)] = "SBT"
                coverage_terms_sheet["L" + str(coverage_terms_row)] = "SBT"

                #Populate required value
                if covterm[1] == True:
                    coverage_terms_sheet["M" + str(coverage_terms_row)] = "Yes"
                if covterm[1] == False:
                    coverage_terms_sheet["M" + str(coverage_terms_row)] = "No"
                
                #Populate default value
                if pd.isna(covterm[2]):
                    coverage_terms_sheet["N" + str(coverage_terms_row)] = "<blank>"
                else:
                    coverage_terms_sheet["N" + str(coverage_terms_row)] = covterm[2]

                #Populate states
                if (clause, covterm[0]) in self.SBT_model_covterm_states:
                    states = set(self.SBT_model_covterm_states[clause, covterm[0]])
                    
                    if len(states) == len(US_states) or "A1" in states:
                        coverage_terms_sheet["O" + str(coverage_terms_row)] = "All States"
                    elif len(states) <= 10:
                        coverage_terms_sheet["O" + str(coverage_terms_row)] = ','.join(states)
                    else:
                        difference = US_states.difference(states)
                        coverage_terms_sheet["O" + str(coverage_terms_row)] = "All states except: " + ','.join(difference)

                if (clause, covterm[0]) in self.SBT_model_covterm_options:
                    option_list = self.SBT_model_covterm_options[clause, covterm[0]]

                coverage_terms_options_row = populate_covterm_options(option_list, covterm, coverage_code, coverage_terms_options_row, clause)
                
                coverage_terms_row+=1

            return coverage_terms_row, coverage_terms_options_row

        def populate_normal_covterms(items, coverage_code, coverage_terms_row, coverage_terms_options_row):
            for covterm in items:
                option_list = {}

                for term in covterm_term_value_type[coverage_code[0], coverage_code[1], coverage_code[2], covterm]:
                    #It's a parent
                    if coverage_code in coverage:
                        #Populate coverage name
                        coverage_terms_sheet["G" + str(coverage_terms_row)] = coverage[coverage_code]
                        
                        #Populate coverage term name
                        if "LIMIT" in covterm:
                            coverage_terms_sheet["H" + str(coverage_terms_row)] = "Limit"
                        else:
                            coverage_terms_sheet["H" + str(coverage_terms_row)] = "Deductible"

                    #It's a child
                    if coverage_code in limit_children:
                        if coverage_code in limit_child_parent and limit_child_parent[coverage_code] in parent_coverage:
                            #Populate coverage name
                            coverage_terms_sheet["G" + str(coverage_terms_row)] = coverage[parent_coverage[limit_child_parent[coverage_code]]]
                        
                        #Populate coverage term name
                        if "LIMIT" in covterm:
                            coverage_terms_sheet["H" + str(coverage_terms_row)] = limit_children[coverage_code] + " - Limit"
                        if "DEDUCT" in covterm:
                            coverage_terms_sheet["H" + str(coverage_terms_row)] = limit_children[coverage_code] + " - Deductible"

                    if coverage_code in coverage or coverage_code in limit_children:
                        #Populate date
                        coverage_terms_sheet["C" + str(coverage_terms_row)] = self.today

                        #Populate files used
                        coverage_terms_sheet["D" + str(coverage_terms_row)] = self.loaded_files[7]

                        #Populate last updated by
                        coverage_terms_sheet["E" + str(coverage_terms_row)] = "Automation Script"

                        #Populate SBT/New
                        coverage_terms_sheet["F" + str(coverage_terms_row)] = "New"
                    
                        #Populate program
                        coverage_terms_sheet["I" + str(coverage_terms_row)] = coverage_code[1]

                        #Populate category
                        coverage_terms_sheet["J" + str(coverage_terms_row)] = coverage_code[2]
                        
                        #Populate term type and value type
                        if term in direct_term:
                            coverage_terms_sheet["K" + str(coverage_terms_row)] = "Direct"
                            coverage_terms_sheet["L" + str(coverage_terms_row)] = "Other"
                        if term in option_term:
                            coverage_terms_sheet["K" + str(coverage_terms_row)] = "Option"
                            coverage_terms_sheet["L" + str(coverage_terms_row)] = "Other"

                        #Populate default value
                        if (coverage_code[0], coverage_code[1], coverage_code[2], covterm) in covterm_default_value:
                            coverage_terms_sheet["N" + str(coverage_terms_row)] = covterm_default_value[coverage_code[0], coverage_code[1], coverage_code[2], covterm]
                        else:
                            coverage_terms_sheet["N" + str(coverage_terms_row)] = "<blank>"

                        #Populate states
                        states = covterm_states[coverage_code[0], coverage_code[1], coverage_code[2], covterm, term]
                    
                        if len(states) == len(US_states) or "A1" in states:
                            coverage_terms_sheet["O" + str(coverage_terms_row)] = "All States"
                        elif len(states) <= 10:
                            coverage_terms_sheet["O" + str(coverage_terms_row)] = ','.join(states)
                        else:
                            difference = US_states.difference(states)
                            coverage_terms_sheet["O" + str(coverage_terms_row)] = "All states except: " + ','.join(difference)

                        if (coverage_code[0], coverage_code[1], coverage_code[2], covterm) in covterm_options_list:
                            option_list = list(covterm_options_list[coverage_code[0], coverage_code[1], coverage_code[2], covterm])
                            option_list = [a for a in option_list if str(a) != 'nan']

                        coverage_terms_options_row = populate_covterm_options(option_list, covterm, coverage_code, coverage_terms_options_row)

                        coverage_terms_row+=1

            return coverage_terms_row, coverage_terms_options_row
    
        self.files_used = ', '.join(self.loaded_files[:7])

        ou_and_uw_exclusions = self.Exclusions.groupby('COVERAGE_CODE')[['PRODUCT_NAME', 'COMPANY_NAME']].apply(lambda x: x.values.tolist()).to_dict()

        ou_abbreviations = {"Berkley Asset Protection":"BAPU","Berkley Agribusiness":"BARS","Berkley Fire & Marine":"BFM","Berkley Life Sciences":"BLS","Berkley Oil & Gas":"BOG","Berkley Risk Administrators Co":"BRAC","Berkley Human Services":"BHS","Berkley Program Specialists":"BUP","Berkley Renewable Energy":"BRE","Berkley Technology Underwriters":"BTU","Berkley Financial Specialists":"FIN","Berkley Entertainment":"BEI","Berkley Medical Excess Underwriters":"BMU","Berkley Prime Transportation":"BPT","Intrepid Direct Insurance":"IDI","Carolina Casualty Insurance":"CCI","Berkley Healthcare":"BHC","Berkley Custom Insurance":"BCI","Berkley Construction Solutions":"BCS","Berkley Small Business":"BSB","Berkley Enterprise Risk Solutions":"BERS","Berkley Product Protection":"BPP","Non Specific Operating Unit":"NSOU","BPS - Agent Will Bill":"BPS","Berkley Shared Services":"BSS","American Mining":"AMI"}

        transactions = self.Transaction_types.set_index("Form Number").to_dict()["RENEWAL_ACTION_C"]
        
        #Dictionary for SBT Forms ID->Coverage Description
        sbt = self.SBT_model.set_index("Form_ID").to_dict()["Description"]

        #Dictionary for SBT Form ID->Type of form (exclusion, condition)
        sbt_type = self.SBT_model.set_index("Form_ID").to_dict()["Type"]

        #Dictionary for SBT Form ID->Existence of Coverage
        sbt_eoc = self.SBT_model.set_index("Form_ID").to_dict()["Existence"]

        #Dictionary for SBT Form ID->Category
        sbt_form_to_category = self.SBT_model.groupby("Form_ID")["Category"].apply(lambda x: x.values.tolist()).to_dict()

        #Dictionary for SBT ClausePatternCode->Category
        sbt_clause_to_category = self.SBT_model.set_index("ClausePatternCode").to_dict()["Category"]

        #Dictionary for SBT (Form ID, Category)->Clause
        sbt_form_id_and_category_to_clause = self.SBT_model.set_index(["Form_ID", "Category"]).to_dict()["ClausePatternCode"]

        #Dictionary for SBT Form ID->ClausePatternCode
        sbt_form_to_clause = self.SBT_model.groupby("Form_ID")["ClausePatternCode"].apply(lambda x: x.values.tolist()).to_dict()

        #List of valid US States
        US_states = {"AK","AL","AR","AZ","CA","CO","CT","DC","DE","FL","GA","HI","IA","ID","IL","IN","IZ","KS","KY","LA","MA","MD","ME","MI","MN","MO","MS","MT","NC","ND","NE","NH","NJ","NM","NV","NY","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VA","VT","WA","WI","WV","WY"}
            
        #Hashtable for parent coverage description->child coverage description
        parent_child = defaultdict(set)

        #Hashtable for coverage code->coverage description
        coverage = defaultdict()

        #Hashtable for parent coverage description->states
        cov_states = defaultdict(set)

        #Hashtable for parent coverage description->ASLOB Code
        major_peril = defaultdict()

        #Hashtable for parent coverage description->premium bearing
        premium = defaultdict()

        #Hashtable for parent coverage description->existence of coverage
        existence = defaultdict()

        #Hashtable for parent coverage description->subline C code
        subline = defaultdict()

        #Hashtable for parent coverage description->scheduled (y/n)
        parent_scheduled = defaultdict()

        #Hashtable for child coverage description->scheduled (y/n)
        child_scheduled = defaultdict()

        #Hashtable for parent id -> coverage description
        parent_id = defaultdict()

        #Hashtable for child coverage description->child state
        covterm_states = defaultdict(set)

        #Hashtable for exclusions
        exclusions = defaultdict()

        #Hashtable for conditions
        conditions = defaultdict()

        #Hashtable for common forms
        common_forms = defaultdict()

        #Dictionary for state amendatory endorsement forms
        state_amendatory = defaultdict()

        #Dictionary of limit child coverage ID->coverage description
        limit_children = defaultdict()

        #Dictionary of limit child coverage ID->limit parent coverage ID
        limit_child_parent = defaultdict()

        #Dictionary of limit coverage ID->covterm
        covterms = defaultdict(set)

        #Dictionary of [limit coverage ID, covterm]->term type
        covterm_term_value_type = defaultdict(set)

        #Dictionary of [limit coverage ID, covterm]->default value
        covterm_default_value = defaultdict()
        
        #Dictionary of [limit coverage ID, covterm]->states
        covterm_options_states = defaultdict(set)

        #Dictionary of [limit coverage ID, covterm]->covterm options
        covterm_options_list = defaultdict(set)

        #List of parent coverages within SBT model that are needed for covterm analysis
        sbt_parent_coverages = defaultdict(set)

        for index, row in self.Coverages.iterrows():
            if not pd.isna(row["PROGRAM_NAME"]) and "FPP" not in row["PROGRAM_NAME"]:
                #Child / Covterm
                if row["COVERAGE_ID"] != row["PARENT_COVERAGE_ID"]:
                    child_scheduled[row["COVERAGE_CODE"].rstrip()] = row["SCHD_COVERAGE_F"]

                #Parent
                else:
                    coverage[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = row["COVERAGE_DESC"]
                    cov_states[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()].add(row["STATE_CODE"])
                    premium[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = row["CNTRB_TO_PREMIUM_F"]
                    existence[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = [row["REQUIRED_COV_F"], row["AUTO_ADD_COV_F"]]
                    subline[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = str(row["SUBLINE_C"])
                    major_peril[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = str(row["MAJOR_PERIL_C"])
                    parent_scheduled[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = row["SCHD_COVERAGE_F"]
                    parent_id[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()] = row["PARENT_COVERAGE_ID"]
        
        for index, row in self.Coverages.iterrows():
            if row["COVERAGE_CODE"].rstrip() in parent_id and row["COVERAGE_ID"] != row["PARENT_COVERAGE_ID"] and not pd.isna(row["PROGRAM_NAME"]) and "FPP" not in row["PROGRAM_NAME"]:
                parent_child[parent_id[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_C"].rstrip()]].add(row["COVERAGE_DESC"])
        
        #Dictionary for Coverage Code->ID
        parent_coverage = dict((v,k) for k,v in parent_id.items())

        #Convert coverage dictionary to dataframe to be able to join with Forms dataframe
        self.Coverages['COVERAGE_CODE'] = self.Coverages['COVERAGE_CODE'].apply(lambda x: x.rstrip())
        self.Coverages['ENTITY_C'] = self.Coverages['ENTITY_C'].apply(lambda x: x.rstrip())
        df = self.Coverages[['COVERAGE_CODE', 'PROGRAM_NAME', 'ENTITY_C']]
        df = df.drop_duplicates()

        #Remove trailing whitespaces from form number and convert edition to mm/yy format
        self.Forms['FORM_NBR'] = self.Forms['FORM_NBR'].apply(lambda x: x.rstrip())
        self.Forms['FORM_EDITION'] = self.Forms["FORM_EDITION"].dt.strftime('%m/%y')
        self.Forms['COVERAGE_CODE'] = self.Forms['COVERAGE_CODE'].apply(lambda x: x.rstrip())
        self.Forms['ENTITY_CODE'] = self.Forms['ENTITY_CODE'].astype(str)
        self.Forms['ENTITY_CODE'] = self.Forms['ENTITY_CODE'].apply(lambda x: x.rstrip())
        self.Forms.rename(columns = {'ENTITY_CODE':'ENTITY_C'}, inplace = True)

        inference = self.Forms.set_index(['COVERAGE_CODE', 'PROGRAM_NAME', 'ENTITY_C', 'FORM_NBR','FORM_EDITION']).to_dict()['ROLL_ON_CND3_CODE']

        #Group form states by [form number, form edition]
        form_states = self.Forms.groupby(['COVERAGE_CODE', 'PROGRAM_NAME', 'ENTITY_C', 'FORM_NBR','FORM_EDITION'])['STATE_CODE'].apply(lambda x: x.values.tolist()).to_dict()
        self.Forms.drop(columns = 'STATE_CODE', inplace = True)

        #Parent coverage dictionary joined with Forms file
        self.Forms.drop_duplicates(inplace=True)
        parent_forms = pd.merge(df, self.Forms, on=['COVERAGE_CODE', 'PROGRAM_NAME', 'ENTITY_C'], how = 'inner')
        
        #Dictionary of coverage description->[form #, form title, form edition]
        parent_forms = parent_forms.groupby(['COVERAGE_CODE', 'PROGRAM_NAME', 'ENTITY_C'])[['FORM_NBR', 'Form Title', 'FORM_EDITION']].apply(lambda x: x.values.tolist()).to_dict()

        for index, row in self.Limits.iterrows():
            #child coverage
            if row["COVERAGE_ID"] != row["PARENT_COVERAGE_ID"]:
                limit_children[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip()] = row["COVERAGE_DESC"]
                limit_child_parent[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip()] = row["PARENT_COVERAGE_ID"]

            covterms[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip()].add(row["LIMIT_DED_OCCUR_C"].rstrip())
            covterm_term_value_type[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip(), row["LIMIT_DED_OCCUR_C"].rstrip()].add(row["LIMIT_DED_OPTION"].rstrip())
            covterm_states[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip(), row["LIMIT_DED_OCCUR_C"].rstrip(), row["LIMIT_DED_OPTION"].rstrip()].add(row["STATE_CODE"])

            if row["DEFAULT_FLAG"] == "Y":
                covterm_default_value[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip(), row["LIMIT_DED_OCCUR_C"].rstrip()] = row["LIMIT_DED_DESC"]

            covterm_options_states[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip(), row["LIMIT_DED_OCCUR_C"].rstrip()].add(row["STATE_CODE"])
            covterm_options_list[row["COVERAGE_CODE"].rstrip(), row["PROGRAM_NAME"], row["ENTITY_CODE"].rstrip(), row["LIMIT_DED_OCCUR_C"].rstrip()].add(row["LIMIT_DED_DESC"])

        #Gather exclusions, conditions and common forms from Forms to Coverages File
        for cov in coverage:
            if cov in parent_forms:
                for form in parent_forms[cov][:]:
                    form_pattern = form[0].replace(" ","") + form[2].replace('/'," ").replace(" ","")

                    #If the word 'Exclusion' is in the form name
                    if "Exclusion" in form[1]:
                        #If the coverage isn't already in the hashtable as a key, add it now
                        if cov not in exclusions:
                            exclusions[cov] = []

                        #Add form to exclusions dictionary and remove it from the parent form dictionary so that it only prints in the 'Exclusions & Forms' tab
                        exclusions[cov].append(form)
                        parent_forms[cov].remove(form)

                    #If the word 'Amendatory' is in the form name
                    elif "Amendatory" in form[1]:
                        #If the coverage isn't already in the hashtable as a key, add it now
                        if cov not in state_amendatory:
                            state_amendatory[cov] = []

                        #Add form to state amendatory dictionary and remove it from the parent form dictionary so that it only prints in the 'State Amendatory Endorsements' tab
                        state_amendatory[cov].append(form)
                        parent_forms[cov].remove(form)

                    elif form_pattern in sbt and form_pattern in sbt_type:
                        #Check the last 4 characters in the 'Type' column within SBT extract
                        suffix = sbt_type[form_pattern][-4:]

                        #If the last 4 characters are 'Excl' it's an exclusion
                        if suffix == "Excl":
                            #If the coverage isn't already in the hashtable as a key, add it now
                            if cov not in exclusions:
                                exclusions[cov] = []

                            #Add form to exclusions dictionary and remove it from the parent form dictionary so that it only prints in the 'Exclusions & Forms' tab
                            exclusions[cov].append(form)
                            parent_forms[cov].remove(form)

                        #If the last 4 characters are 'Cond' it is a condition
                        if suffix == "Cond":
                            #If the coverage isn't already in the hashtable as a key, add it now
                            if cov not in conditions:
                                conditions[cov] = []
                            
                            #Add form to conditions dictionary and remove it from the parent form dictionary so that it only prints in the 'Conditions & Forms' tab
                            conditions[cov].append(form)
                            parent_forms[cov].remove(form)

                        sbt_parent_coverages[cov].add(form_pattern)
                        
                    elif self.lob == "GL" and ((form_pattern[:2] != "CG" or (form_pattern[:2] == "CG" and not form_pattern[2:4].isnumeric())) or "TC" in form[0]):
                            #If the coverage isn't already in the hashtable as a key, add it now
                            if cov not in common_forms:
                                common_forms[cov] = []
                            
                            #Add form to conditions dictionary and remove it from the parent form dictionary so that it only prints in the 'Common Forms' tab
                            common_forms[cov].append(form)
                            parent_forms[cov].remove(form)
                    
                    elif self.lob == "CP" and ((form_pattern[:2] != "CP" or (form_pattern[:2] == "CP" and not form_pattern[2:4].isnumeric())) or "TC" in form[0]):
                            #If the coverage isn't already in the hashtable as a key, add it now
                            if cov not in common_forms:
                                common_forms[cov] = []
                            
                            #Add form to conditions dictionary and remove it from the parent form dictionary so that it only prints in the 'Common Forms' tab
                            common_forms[cov].append(form)
                            parent_forms[cov].remove(form)

                    elif self.lob == "CA":
                        if (form_pattern[:2] != "CA" and form_pattern[:2] != "CC") or ("TC" in form[0]) or ((form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and not form_pattern[2:4].isnumeric()):
                            #If the coverage isn't already in the hashtable as a key, add it now
                            if cov not in common_forms:
                                common_forms[cov] = []
                            
                            #Add form to conditions dictionary and remove it from the parent form dictionary so that it only prints in the 'Common Forms' tab
                            common_forms[cov].append(form)
                            parent_forms[cov].remove(form)

                    elif self.lob == "IM" and ((form_pattern[:2] != "IM" or (form_pattern[:2] == "IM" and not form_pattern[2:4].isnumeric())) or "TC" in form[0]):
                            #If the coverage isn't already in the hashtable as a key, add it now
                            if cov not in common_forms:
                                common_forms[cov] = []
                            
                            #Add form to conditions dictionary and remove it from the parent form dictionary so that it only prints in the 'Common Forms' tab
                            common_forms[cov].append(form)
                            parent_forms[cov].remove(form)
                    
                    else:
                        pass
        
        qrg_sbt_forms = []
        qrg_sbt_exclusions = []
        qrg_sbt_conditions = []
        qrg_state_amendatory = []
        qrg_common = [] 

        #Gather exclusions, conditions and common forms from QRG Forms
        for form in self.QRG_forms:
            form_pattern = form[0].replace(" ","") + form[2].replace('/'," ").replace(" ","")
            
            if form_pattern in sbt_type:
                #Check the last 4 characters in the 'Type' column within SBT extract
                suffix = sbt_type[form_pattern][-4:]

                #If the last 4 characters are 'Excl' it's an exclusion
                if suffix == "Excl":
                    qrg_sbt_exclusions.append(form)
                if suffix == "Cond":
                    qrg_sbt_conditions.append(form)
                if suffix == "Cov":
                    qrg_sbt_forms.append(form)
            
            elif "Amendatory" in form[1]:
                if form not in state_amendatory.values():
                    qrg_state_amendatory.append(form)
            
            elif self.lob == "GL" and ((form_pattern[:2] != "CG" or (form_pattern[:2] == "CG" and not form_pattern[2:4].isnumeric())) or "TC" in form[0]):
                if form not in parent_forms.values() and form not in exclusions.values() and form not in conditions.values() and form not in common_forms.values():
                    qrg_common.append(form)
                    
            elif self.lob == "CP" and ((form_pattern[:2] != "CP" or (form_pattern[:2] == "CP" and not form_pattern[2:4].isnumeric())) or "TC" in form[0]):
                if form not in parent_forms.values() and form not in exclusions.values() and form not in conditions.values() and form not in common_forms.values():
                    qrg_common.append(form)

            elif self.lob == "CA":
                if (form_pattern[:2] != "CA" and form_pattern[:2] != "CC") or ("TC" in form[0]) or ((form_pattern[:2] == "CA" or form_pattern[:2] == "CC") and not form_pattern[2:4].isnumeric()):
                    if form not in parent_forms.values() and form not in exclusions.values() and form not in conditions.values() and form not in common_forms.values():
                        qrg_common.append(form)

            elif self.lob == "IM" and ((form_pattern[:2] != "IM" or (form_pattern[:2] == "IM" and not form_pattern[2:4].isnumeric())) or "TC" in form[0]):
                if form not in parent_forms.values() and form not in exclusions.values() and form not in conditions.values() and form not in common_forms.values():
                    qrg_common.append(form)

            else:
                pass
            
        #Begin writing to product model
        product_model = openpyxl.load_workbook(self.template)

        #first 2 rows in product model are headers
        coverages_and_forms_row = 3
        exclusions_and_forms_row = 3
        conditions_and_forms_row = 3
        common_forms_row = 3
        state_amendatory_row = 3
        
        for cov_code in coverage:
            num_coverage_rows = 0
            num_condition_rows = 0
            num_exclusion_rows = 0
            num_common_forms_row = 0
            num_state_amendatory_row = 0

            if cov_code in parent_forms:
                num_coverage_rows = len(parent_forms[cov_code])
            if cov_code in exclusions:
                num_exclusion_rows = len(exclusions[cov_code])
            if cov_code in conditions:
                num_condition_rows = len(conditions[cov_code])
            if cov_code in common_forms:
                num_common_forms_row = len(common_forms[cov_code])
            if cov_code in state_amendatory:
                num_state_amendatory_row = len(state_amendatory[cov_code])
            if num_coverage_rows == 0 and num_exclusion_rows == 0 and num_condition_rows == 0 and num_common_forms_row == 0 and num_state_amendatory_row == 0:
                num_coverage_rows = 1

            if num_coverage_rows > 0:
                cov_index = 0 
                sheet = product_model["Coverages & Forms"]
                
                while cov_index <= num_coverage_rows - 1:
                     #Check if this coverage has a form
                    if cov_code in parent_forms and cov_index < num_coverage_rows:
                        current_row = print_forms(sheet, coverages_and_forms_row, cov_index, "General")
                    
                    while coverages_and_forms_row < current_row:
                        print_coverages(sheet, coverages_and_forms_row)
                        coverages_and_forms_row+=1
                    
                    cov_index+=1

            if num_exclusion_rows > 0:
                exclusion_index = 0
                sheet = product_model["Exclusions & Forms"]
                
                while exclusion_index <= num_exclusion_rows - 1:
                    #Check if this coverage has a form
                    if cov_code in exclusions and exclusion_index < num_exclusion_rows:
                        current_row = print_forms(sheet, exclusions_and_forms_row, exclusion_index, "Exclusion")

                    while exclusions_and_forms_row < current_row:
                        print_coverages(sheet, exclusions_and_forms_row)
                        exclusions_and_forms_row+=1

                    exclusion_index+=1

            if num_condition_rows > 0:
                condition_index = 0
                sheet = product_model["Conditions & Forms"]
                
                while condition_index <= num_condition_rows - 1:
                    #Check if this coverage has a form
                    if cov_code in conditions and condition_index < num_condition_rows:
                        current_row = print_forms(sheet, conditions_and_forms_row, condition_index, "Condition")

                    while conditions_and_forms_row < current_row:
                        print_coverages(sheet, conditions_and_forms_row)
                        conditions_and_forms_row+=1

                    condition_index+=1
                    
            if num_common_forms_row > 0:
                common_form_index = 0
                sheet = product_model["Common Forms"]
                
                while common_form_index <= num_common_forms_row - 1:
                    print_common_coverages(sheet, common_forms_row)

                    #Check if this coverage has a form
                    if cov_code in common_forms and common_form_index < num_common_forms_row:
                        print_common_forms(sheet, common_forms_row, common_form_index)

                    common_form_index+=1
                    common_forms_row+=1

            if num_state_amendatory_row > 0:
                state_amendatory_index = 0
                sheet = product_model["State Amendatory Endorsements"]
                
                while state_amendatory_index <= num_state_amendatory_row - 1:
                    print_amendatory_coverages(sheet, state_amendatory_row)

                    #Check if this coverage has a form
                    if cov_code in state_amendatory and state_amendatory_index < num_state_amendatory_row:
                        print_amendatory_forms(sheet, state_amendatory_row, state_amendatory_index)

                    state_amendatory_index+=1
                    state_amendatory_row+=1

        for form in qrg_sbt_forms:
            sheet = product_model["Coverages & Forms"]
            coverages_and_forms_row = print_qrg_sbt_forms(sheet, coverages_and_forms_row, form)

        for form in qrg_sbt_exclusions:
            sheet = product_model["Exclusions & Forms"]
            exclusions_and_forms_row = print_qrg_sbt_forms(sheet, exclusions_and_forms_row, form)

        for form in qrg_sbt_conditions:
            sheet = product_model["Conditions & Forms"]
            conditions_and_forms_row = print_qrg_sbt_forms(sheet, conditions_and_forms_row, form)

        for form in qrg_state_amendatory:
            sheet = product_model["State Amendatory Endorsements"]
            state_amendatory_row = print_qrg_forms(sheet, state_amendatory_row, form)

        for form in qrg_common:
            sheet = product_model["Common Forms"]
            common_forms_row = print_qrg_forms(sheet, common_forms_row, form)

        #Begin writing to Coverage Terms worksheet
        coverage_terms_sheet = product_model["Coverage terms"]
        coverage_terms_row = 3

        #Begin writing to Coverage Term Options worksheet
        coverage_term_options_sheet = product_model["Coverage Term Options"]
        coverage_terms_options_row = 3

        direct_term = ['ENTERABLE', 'ENTER_INC', 'DEFAULT_NC', 'DEFAULT_EN', 'LABEL_NC']
        option_term = ['DROPDOWN', 'DEFAULT_DD', 'FILTER_DD', 'LABEL_DD']

        for coverage_code in covterms:
            items = {}

            if coverage_code in sbt_parent_coverages:
                for sbt_form in sbt_parent_coverages[coverage_code]:
                    if sbt_form in sbt_form_to_clause:
                        for clause in sbt_form_to_clause[sbt_form]:
                            if clause in self.SBT_model_covterms:
                                items = self.SBT_model_covterms[clause]
                                coverage_terms_row, coverage_terms_options_row = populate_sbt_covterms(items, coverage_code, sbt_form, coverage_terms_row, coverage_terms_options_row, clause)
            else:
                items = covterms[coverage_code]
                coverage_terms_row, coverage_terms_options_row = populate_normal_covterms(items, coverage_code, coverage_terms_row, coverage_terms_options_row)
            
        product_model.save(self.template)

        messagebox.showinfo('Processed', 'All Excel files have been consolidated!')


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelLoaderApp(root)
    root.mainloop()