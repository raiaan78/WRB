import os
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict
import pandas as pd
import openpyxl

class ExcelLoaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title('ProductModel WRB Accelerator')
        
        # Dataframes for each file type
        self.Coverages = None
        self.Forms = None
        self.Inference = None
        self.Transaction_types = None
        self.Limits = None
        self.SBT_model = None

        #set line of business
        self.lob = ""

        #Path for template file
        self.template = ""

        # Filepaths for display
        self.loaded_files = []

        # Buttons for each file type

        self.sbt_extract_btn = tk.Button(self.root, text='Load SBT Product Model Extract', command=lambda: self.load_file('SBT_extract'))
        self.sbt_extract_btn.pack(pady=10)

        self.coverage_btn = tk.Button(self.root, text='Load Coverage File', command=lambda: self.load_file('coverage'))
        self.coverage_btn.pack(pady=10)

        self.forms_btn = tk.Button(self.root, text='Load Forms File', command=lambda: self.load_file('forms'))
        self.forms_btn.pack(pady=10)

        self.inference_btn = tk.Button(self.root, text='Load Form Inference Steps File', command=lambda: self.load_file('inference'))
        self.inference_btn.pack(pady=10)

        self.qrg_btn = tk.Button(self.root, text='Load Forms QRG File', command=lambda: self.load_file('QRG'))
        self.qrg_btn.pack(pady=10)

        self.covterm_options_btn = tk.Button(self.root, text='Load Limit Deductible File', command=lambda: self.load_file('covterm_options'))
        self.covterm_options_btn.pack(pady=10)

        self.input_template_btn = tk.Button(self.root, text='Load Template File', command=lambda: self.load_file('input_template'))
        self.input_template_btn.pack(pady=10)

        #self.coverage_exclusions_btn = tk.Button(self.root, text='Load Coverage Exclusions File', command=lambda: self.load_file('coverage_exclusions'))
        #self.coverage_exclusions_btn.pack(pady=10)

        options = ["GL","CP"]
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

    def load_file(self, file_type):
        filepath = filedialog.askopenfilename(title=f'Select {file_type} Excel File', filetypes=(('Excel Files', '*.xls;*.xlsx;*.xlsm'), ('All Files', '*.*')))
        if not filepath:
            return

        filename = os.path.basename(filepath)

        if file_type == 'coverage' and "Coverage" in filename:
            self.Coverages = pd.read_excel(io=filepath, usecols = "A:B, E:I, L, R:S, W, X, AB, AN")
            self.coverage_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'forms' and "Forms To Coverages" in filename:
            self.Forms = pd.read_excel(io=filepath, usecols = "A:C, F, H:I")
            self.forms_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'inference' and "Steps" in filename:
            self.Inference = pd.read_excel(io=filepath)
            self.inference_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'QRG' and "QRG" in filename:
            self.Transaction_types = pd.read_excel(io=filepath, usecols = "B, H")
            self.Transaction_types.drop_duplicates(inplace=True)
            self.qrg_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'covterm_options' and "Limit" in filename:
            self.Limits = pd.read_excel(io=filepath, usecols = "C:D, F, H, L")
            self.covterm_options_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'SBT_extract' and "ProductModelExport" in filename:
            self.SBT_model = pd.read_excel(io=filepath, sheet_name = "Clause", usecols = "B, C, I")
            #Parse the SBT model since multiple form IDs are within one cell in some cases
            self.SBT_model = self.SBT_model.assign(Form_ID = self.SBT_model['Form(s)'].str.split(r'\n')).explode('Form(s)')
            self.SBT_model = self.SBT_model.explode('Form_ID')
            self.SBT_model = self.SBT_model[["Description", "Type", "Form_ID"]]
            self.SBT_model.drop_duplicates(inplace=True)
            self.sbt_extract_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        elif file_type == 'input_template' and "Product Model" in filename:
            self.template = filepath
            self.input_template_btn.config(state=tk.DISABLED)
            self.loaded_files.append(filename)

        else:
            messagebox.showerror("Error", f"Invalid file selected for {file_type}. Please select the correct file.")

        # Update the loaded files label
        self.loaded_label.config(text=', '.join(self.loaded_files))

        # Enable the process button if all files are loaded
        if len(self.loaded_files) == 7:  # Assuming you have 7 files to load
            self.process_btn.config(state=tk.NORMAL)
            self.set_lob()

    def process_files(self):
        def print_coverages(sheet, row):
            #Populate coverage name
            cell = "G" + str(row)
            sheet[cell] = cov_name

            #Populate operating units
            cell2 = "O" + str(row)
            operating_units = set()

            #Populate underwriting companies
            cell7 = "P" + str(row)
            uw_companies = ""
            for j in cov_opunit[cov_name].keys():
                uw_companies = uw_companies + j + "(" + ','.join(cov_opunit[cov_name][j]) + ")" + "\n"
                operating_units.update(cov_opunit[cov_name][j])

            sheet[cell2] = ','.join(operating_units)
            sheet[cell7] = uw_companies

            #Populate coverage states
            cell3 = "I" + str(row)

            if len(cov_states[cov_name]) == len(US_states) or "A1" in cov_states[cov_name]:
                sheet[cell3] = "All States"
            elif len(cov_states[cov_name]) <= 10:
                sheet[cell3] = ','.join(cov_states[cov_name])
            else:
                difference = US_states.difference(cov_states[cov_name])
                sheet[cell3] = "All states except: " + ','.join(difference)

            #Populate ASOLB/Major Peril Code
            cell4 = "Q" + str(row)
            sheet[cell4] = ','.join(major_peril[cov_name])

            #Populate Offering/Program
            cell5 = "J" + str(row)
            sheet[cell5] = program[cov_name]

            #Populate Premium Bearing
            cell6 = "M" + str(row)
            sheet[cell6] = premium[cov_name]

            #Populate existence of coverage
            cell8 = "L" + str(row)
            eoc = existence[cov_name]

            if eoc[0] == 'Y' and eoc[1] == 'N':
                sheet[cell8] = "Required"
            elif eoc[0] == 'N' and eoc[1] == 'N':
                sheet[cell8] = "Electable"
            else:
                sheet[cell8] = "Suggested"

            if self.lob == "GL":
                #Populate Subline C items
                code = subline[cov_name]

                if code == '          ':
                    pass
                elif code == 334 or code == 336:
                    cell15 = "R" + str(row)
                    cell16 = "S" + str(row)
                    sheet[cell15] = "x"
                    sheet[cell16] = "x"
                elif code == 332:
                    cell15 = "T" + str(row)
                    cell16 = "U" + str(row)
                    sheet[cell15] = "x"
                    sheet[cell16] = "x"
                elif code == 317:
                    cell15 = "V" + str(row)
                    cell16 = "W" + str(row)
                    sheet[cell15] = "x"
                    sheet[cell16] = "x"
                elif code == 325:
                    cell15 = "X" + str(row)
                    cell16 = "Y" + str(row)
                    sheet[cell15] = "x"
                    sheet[cell16] = "x"
                elif code == 360:
                    cell15 = "Z" + str(row)
                    cell16 = "AA" + str(row)
                    sheet[cell15] = "x"
                    sheet[cell16] = "x"
                else:
                    sheet[cell4] = str(code) + "/" + sheet[cell4].value
                
            #Populate scheduled field
            cell17 = "N" + str(row)
            if parent_scheduled[cov_name] == "Y":
                sheet[cell17] = "Y"
            else:
                answer = False
                for child in parent_child[cov_name]:
                    if child_scheduled[child] == "Y":
                        answer = True
                        break

                if answer == True:
                    sheet[cell17] = "Y"
                else:
                    sheet[cell17] = "N"

        def print_forms(sheet, row, index, coverage_type):
            #Populate form info
            if coverage_type == "General":
                form_number = parent_forms[cov_name][index][0]
                form_name = parent_forms[cov_name][index][1]
                form_edition = parent_forms[cov_name][index][2].replace('/'," ")
                form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")
            elif coverage_type == "Exclusion":
                form_number = exclusions[cov_name][index][0]
                form_name = exclusions[cov_name][index][1]
                form_edition = exclusions[cov_name][index][2].replace('/'," ")
                form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")
            else:
                form_number = conditions[cov_name][index][0]
                form_name = conditions[cov_name][index][1]
                form_edition = conditions[cov_name][index][2].replace('/'," ")
                form_pattern = form_number.replace(" ","") + form_edition.replace(" ","")

            if self.lob == "CP":
                cell9 = "AF" + str(row)
                cell10 = "AG" + str(row)
                cell11 = "AH" + str(row)
                cell12 = "AI" + str(row)
            else:
                cell9 = "T" + str(row)
                cell10 = "U" + str(row)
                cell11 = "V" + str(row)
                cell12 = "W" + str(row)

            sheet[cell9] = form_pattern
            sheet[cell10] = form_number
            sheet[cell11] = form_edition
            sheet[cell12] = form_name

            #Populate SBT/OOTB
            cell18 = "F" + str(row)

            if form_pattern[:2] == "CG" and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                sheet[cell18] = "New"
            elif form_number.replace(" ","") in sbt.keys():
                sheet[cell18] = "SBT"
                #Change coverage name to whatever is in the SBT model
                #sheet[cell] = sbt[form_number.replace(" ","")]
            else:
                pass

            #ISO/Proprietary
            cell13 = "H" + str(row)
            if (form_pattern[:2] == "CG" or form_pattern[:2] == "CL") and form_pattern[2:4].isnumeric() and int(form_pattern[2:4]) >= 83:
                sheet[cell13] = "Proprietary"
            else:
                sheet[cell13] = "ISO"

            #Populate form states
            if self.lob == "GL":
                cell14 = "AJ" + str(row)
            else:
                cell14 = "X" + str(row)

            state_set = set(form_states[form_number])

            if len(state_set) == len(US_states) or "A1" in state_set:
                sheet[cell14] = "All States"
            elif len(state_set) <= 10:
                sheet[cell14] = ','.join(state_set)
            else:
                difference = US_states.difference(state_set)
                sheet[cell14] = "All states except: " + ','.join(difference)

            #Populate Transaction Types
            if self.lob == "GL":
                cell19 = "AN" + str(row)
            else:
                cell19 = "AB" + str(row)

            if transactions[form_number] == "RETAIN":
                sheet[cell19] = "Submission, Policy, Change, Rewrite, Rewrite New Account, Renewal"
            else:
                sheet[cell19] = "Submission, Policy, Change, Rewrite, Rewrite New Account"

    # Your consolidation code goes here. Use self.coverage_df, self.covterm_df, and self.forms_df.
        
        '''
        def generate_text(group):
            logic_texts = []

            for _, row in group.iterrows():
                step_logic = f"If {row['STEP_NAME']}"
                if pd.notna(row['GOTO_STEP_ON_TRUE']):
                    step_logic += f", then go to step {row['GOTO_STEP_ON_TRUE']}."
                if pd.notna(row['GOTO_STEP_ON_FALSE']):
                    step_logic += f" If not, then go to step {row['GOTO_STEP_ON_FALSE']}."
                logic_texts.append(step_logic)

            return "\n- ".join(logic_texts)

        #Dataframe for inference logic
        Inference = pd.read_excel(io=file_name6)
        Inference = Inference[Inference["OPERATOR"] != 'Move']
        result = Inference.groupby('ROLL_ON_CND3_CODE').apply(generate_text).reset_index()
        result.columns = ['ROLL_ON_CND3_CODE', 'Inference Logic']
        result.to_excel("output.xlsx")
        '''
        transactions = self.Transaction_types.set_index("Form Number").to_dict()["RENEWAL_ACTION_C"]
        
        #Dictionary for SBT Forms ID->Coverage Description
        sbt = self.SBT_model.set_index("Form_ID").to_dict()["Description"]

        #Dictionary for SBT Form ID->Type of form (exclusion, condition)
        sbt_type = self.SBT_model.set_index("Form_ID").to_dict()["Type"]

        #List of valid US States
        US_states = {"AK","AL","AR","AZ","CA","CO","CT","DC","DE","FL","GA","HI","IA","ID","IL","IN","IZ","KS","KY","LA","MA","MD","ME","MI","MN","MO","MS","MT","NC","ND","NE","NH","NJ","NM","NV","NY","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VA","VT","WA","WI","WV","WY"}
            
        #Hashtable for parent coverage description->child coverage description
        parent_child = defaultdict(set)

        #Hashtable for coverage code->coverage description
        coverage = defaultdict()

        #Hashtable for parent coverage description->operating unit
        cov_opunit = defaultdict(lambda: defaultdict(set))

        #Hashtable for parent coverage description->states
        cov_states = defaultdict(set)

        #Hashtable for parent coverage description->ASLOB Code
        major_peril = defaultdict(set)

        #Hashtable for parent coverage description->offering/program
        program = defaultdict()

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

        for index, row in self.Coverages.iterrows():
            #Child / Covterm
            if row["COVERAGE_ID"] != row["PARENT_COVERAGE_ID"]:
                child_scheduled[row["COVERAGE_DESC"]] = row["SCHD_COVERAGE_F"]
                covterm_states[row["COVERAGE_DESC"]].add(row["STATE_CODE"])
            
            #Parent
            else:
                cov_opunit[row["COVERAGE_DESC"]][row["COMPANY_NAME"]].add(row["Operating Unit"].rstrip())
                coverage[row["COVERAGE_CODE"]] = row["COVERAGE_DESC"]
                cov_states[row["COVERAGE_DESC"]].add(row["STATE_CODE"])
                major_peril[row["COVERAGE_DESC"]].add(str(row["MAJOR_PERIL_C"]))
                program[row["COVERAGE_DESC"]] = row["PROGRAM_NAME"]
                premium[row["COVERAGE_DESC"]] = row["CNTRB_TO_PREMIUM_F"]
                existence[row["COVERAGE_DESC"]] = [row["REQUIRED_COV_F"], row["AUTO_ADD_COV_F"]]
                subline[row["COVERAGE_DESC"]] = row["SUBLINE_C"]
                parent_scheduled[row["COVERAGE_DESC"]] = row["SCHD_COVERAGE_F"]
                parent_id[row["PARENT_COVERAGE_ID"]] = row["COVERAGE_DESC"]

        for index, row in self.Coverages.iterrows():
            if row["COVERAGE_ID"] != row["PARENT_COVERAGE_ID"]:
                parent_child[parent_id[row["PARENT_COVERAGE_ID"]]].add(row["COVERAGE_DESC"])

        #Convert coverage dictionary to dataframe to be able to join with Forms dataframe
        df = pd.DataFrame.from_dict(coverage, orient= "index").reset_index()
        df.columns = ["COVERAGE_CODE", "COVERAGE_DESC"]

        #Remove trailing whitespaces from form number and convert edition to mm/yy format
        self.Forms['FORM_NBR'] = self.Forms['FORM_NBR'].apply(lambda x: x.rstrip())
        self.Forms['FORM_EDITION'] = self.Forms["FORM_EDITION"].dt.strftime('%m/%y')

        #Group form states by form #
        form_states = self.Forms.groupby('FORM_NBR')['STATE_CODE'].apply(lambda x: x.values.tolist()).to_dict()
        self.Forms.drop(columns = 'STATE_CODE', inplace = True)

        #Parent coverage dictionary joined with Forms file
        self.Forms.drop_duplicates(inplace=True)
        parent_forms = pd.merge(df, self.Forms, on=['COVERAGE_CODE', 'COVERAGE_DESC'], how = 'inner')

        #Dictionary of coverage description->[form #, form title, form edition]
        parent_forms = parent_forms.groupby('COVERAGE_DESC')[['FORM_NBR', 'Form Title', 'FORM_EDITION']].apply(lambda x: x.values.tolist()).to_dict()

        #Begin writing to product model
        #product_model = openpyxl.load_workbook("C:\\Users\\rvalli001\\Desktop\\WRB\\Coverages\\GL\\PC - SSP - GL Product Model & Forms Inference_Draft.xlsx")
        product_model = openpyxl.load_workbook(self.template)

        #Gather exclusions and conditions
        for cov in parent_forms.keys():
            for form in parent_forms[cov][:]:
                form_pattern = form[0].replace(" ","") + form[2].replace('/'," ").replace(" ","")

                if "Exclusion" in form[1]:
                    if cov not in exclusions:
                        exclusions[cov] = []

                    exclusions[cov].append(form)
                    parent_forms[cov].remove(form)

                elif form_pattern in sbt.keys() and form_pattern in sbt_type.keys():
                    suffix = sbt_type[form_pattern][-4:]

                    if suffix == "Excl":
                        if cov not in exclusions:
                            exclusions[cov] = []

                        exclusions[cov].append(form)
                        parent_forms[cov].remove(form)
                    elif suffix == "Cond":
                        if cov not in conditions:
                            conditions[cov] = []

                        conditions[cov].append(form)
                        parent_forms[cov].remove(form)
                    else:
                        pass

                else:
                    pass

        #first 2 rows in product model are headers
        coverages_and_forms_row = 3
        exclusions_and_forms_row = 3
        conditions_and_forms_row = 3
        
        for cov_name in coverage.values():
            num_coverage_rows = 0
            num_condition_rows = 0
            num_exclusion_rows = 0

            if cov_name in parent_forms:
                num_coverage_rows = len(parent_forms[cov_name])
            if cov_name in exclusions:
                num_exclusion_rows = len(exclusions[cov_name])
            if cov_name in conditions:
                num_condition_rows = len(conditions[cov_name])
            if num_coverage_rows == 0 and num_exclusion_rows == 0 and num_condition_rows == 0:
                num_coverage_rows = 1

            if num_coverage_rows > 0:
                cov_index = 0 
                sheet = product_model["Coverages & Forms"]    
                while cov_index <= num_coverage_rows - 1:
                    print_coverages(sheet, coverages_and_forms_row)

                    #Check if this coverage has a form
                    if cov_name in parent_forms and cov_index < num_coverage_rows:
                        print_forms(sheet, coverages_and_forms_row, cov_index, "General")
                    
                    cov_index+=1
                    coverages_and_forms_row+=1

            if num_exclusion_rows > 0:
                exclusion_index = 0
                sheet = product_model["Exclusions & Forms"]    
                while exclusion_index <= num_exclusion_rows - 1:
                    print_coverages(sheet, exclusions_and_forms_row)

                    #Check if this coverage has a form
                    if cov_name in exclusions and exclusion_index < num_exclusion_rows:
                        print_forms(sheet, exclusions_and_forms_row, exclusion_index, "Exclusion")

                    exclusion_index+=1
                    exclusions_and_forms_row+=1

            if num_condition_rows > 0:
                condition_index = 0
                sheet = product_model["Conditions & Forms"]    
                while condition_index <= num_condition_rows - 1:
                    print_coverages(sheet, conditions_and_forms_row)

                    #Check if this coverage has a form
                    if cov_name in conditions and condition_index < num_condition_rows:
                        print_forms(sheet, conditions_and_forms_row, condition_index, "Condition")

                    condition_index+=1
                    conditions_and_forms_row+=1

        #Hashtable for child description->child option states
        covterm_options_states = defaultdict(set)

        #Hashtable for child description->child options
        covterm_options_list = defaultdict(set)

        #Dictionary of coverage description->[form #, form title, form edition]
        covterm_default_values = self.Limits.groupby('COVERAGE_DESC')[['LIMIT_DED_DESC', 'DEFAULT_FLAG']].apply(lambda x: x.values.tolist()).to_dict()

        #Dictionary of coverage description->limit_ded_option
        covterm_term_value_type = self.Limits.groupby('COVERAGE_DESC')['LIMIT_DED_OPTION'].apply(lambda x: x.values.tolist()).to_dict()

        for index, row in self.Limits.iterrows():
            covterm_options_states[row["COVERAGE_DESC"]].add(row["STATE_CODE"])
            covterm_options_list[row["COVERAGE_DESC"]].add(row["LIMIT_DED_DESC"])
            
        #Begin writing to Coverage Terms worksheet
        coverage_terms_sheet = product_model["Coverage terms"]
        coverage_terms_row = 3

        #Begin writing to Coverage Term Options worksheet
        coverage_term_options_sheet = product_model["Coverage Term Options"]
        coverage_terms_options_row = 3

        direct_term = ["ENTERABLE", "ENTER_INC", "DEFAULT_NC", "DEFAULT_EN", "LABEL_NC"]
        option_term = ["DROPDOWN", "DEFAULT_DD", "FILTER_DD"]

        for parent in parent_child.keys():
            for child in parent_child[parent]:
                #Populate coverage terms sheet
                cov_term_parent_output = "G" + str(coverage_terms_row)
                cov_term_child_output = "H" + str(coverage_terms_row)
                
                coverage_terms_sheet[cov_term_parent_output] = parent
                coverage_terms_sheet[cov_term_child_output] = child

                term_type = "J" + str(coverage_terms_row)
                value_type = "K" + str(coverage_terms_row)

                if child in covterm_term_value_type:
                    option = covterm_term_value_type[child]

                    if option in direct_term:
                        coverage_terms_sheet[term_type] = "Direct"
                        coverage_terms_sheet[value_type] = "Other"
                    if option in option_term:
                        coverage_terms_sheet[term_type] = "Option"
                        coverage_terms_sheet[value_type] = "Other"

                child_states = "N" + str(coverage_terms_row)
                states = covterm_states[child]
                
                if len(states) == len(US_states) or "A1" in states:
                    coverage_terms_sheet[child_states] = "All States"
                elif len(states) <= 10:
                    coverage_terms_sheet[child_states] = ','.join(states)
                else:
                    difference = US_states.difference(states)
                    coverage_terms_sheet[child_states] = "All states except: " + ','.join(difference)

                #Populate default value for covterm
                default_value = "M" + str(coverage_terms_row)
                coverage_terms_sheet[default_value] = "<blank>"

                if child in covterm_default_values:
                    for val in covterm_default_values[child]:
                        if val[1] == "Y":
                            coverage_terms_sheet[default_value].value = val[0]
                            break

                coverage_terms_row+=1
                #sorted_covterm_options = sorted(covterm_options_list[child])
                
                for option in covterm_options_list[child]:
                    #Populate Coverage Term Options sheet
                    covterm_options_parent_output = "G" + str(coverage_terms_options_row)
                    covterm_options_child_output = "H" + str(coverage_terms_options_row)   
                    option_name = "J" + str(coverage_terms_options_row)

                    coverage_term_options_sheet[covterm_options_parent_output] = parent
                    coverage_term_options_sheet[covterm_options_child_output] = child
                    coverage_term_options_sheet[option_name] = option

                    child_states2 = "K" + str(coverage_terms_options_row)
                    states = covterm_options_states[child]

                    if len(states) == len(US_states) or "A1" in states:
                        coverage_term_options_sheet[child_states2] = "All States"
                    elif len(states) <= 10:
                        coverage_term_options_sheet[child_states2] = ','.join(states)
                    else:
                        difference = US_states.difference(states)
                        coverage_term_options_sheet[child_states2] = "All states except: " + ','.join(difference)
                    
                    coverage_terms_options_row+=1

        #product_model.save("C:\\Users\\rvalli001\\Desktop\\WRB\\Coverages\\GL\\PC - SSP - GL Product Model & Forms Inference_Draft.xlsx")
        product_model.save(self.template)

        messagebox.showinfo('Processed', 'All Excel files have been consolidated!')


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelLoaderApp(root)
    root.mainloop()