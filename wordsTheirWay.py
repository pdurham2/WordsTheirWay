import pandas as pd
from re import match
from numpy import nan

def check_spelling_stage(level_selection, fp):
    #Create Elementary Spelling Stages
    if level_selection in ["primary", "elementary"]:
        spelling_stages = [
            ("emergent late" , 5),
            ("letter name early" , 10),
            ("letter name middle" , 16),
            ("letter name late" , 23),
            ("within pattern early" , 28),
            ("within pattern middle" , 35),
            ("within pattern late" , 40), 
            ("syllables and affixes early" , 45),
            ("syllables and affixes middle" , 50),
            ("syllables and affixes late" , 55),
            ("derivational relations early" , 60),
            ("derivational relations middle" , 63)
        ]
    elif level_selection == "upper-level":
        spelling_stages = [
            ("within pattern early", 3),
            ("within pattern middle", 12),
            ("within pattern late", 19),
            ("syllables and affixes early", 27),
            ("syllables and affixes middle", 36),
            ("syllables and affixes late", 46),
            ("derivational relations early", 53),
            ("derivational relations middle", 60),
            ("derivational relations late", 66)
        ]
    else:
        print("Incorrect level selection")
        return None
    #Sort by threshold
    spelling_stages.sort(key = lambda x : x[1])
    #if threshold > fp
    stage = next(stage for stage, threshold in spelling_stages if threshold > fp)

    return stage



#Calculate feature points and return answer sheet for a given name
def calculate_feature_points(level_selection, feature_guide, test, name):
    #Get the word list for the given name
    words = test[name]
    #Overwrite the name of the student with test_words
    words.name = "test_words"
    #Join feature guide and test words together
    word_guide = feature_guide.merge(words, on = "feature", how = "left")
    #Check if the name is missing any words
    if word_guide["test_words"].isna().any():
        missing_words = word_guide.loc[word_guide["test_words"].isna(), "feature"]
        missing_words = list(missing_words)
        print(("The provided excel is missing the following features "
            "for {}: {}").format(name, missing_words))
        return False

    #Create answer sheet
    answer_sheet = feature_guide.set_index("feature").applymap(lambda x: nan)
    
    #Initialize feature points to 0
    feature_points = 0
    if level_selection in ["elementary", "primary"]:
        for i in range(len(word_guide)):
            #Initial Consonants Check
            if match("^%s"%word_guide["consonants.initial"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"consonants.initial"] = 1
            #Final Consonants Check
            if match(".*%s[aeiouy]*$"%word_guide["consonants.final"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"consonants.final"] = 1
            #Short Vowels Check
            if match(".*%s+.*"%word_guide["short vowels"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"short vowels"] = 1
            #Digraphs Check
            if match(".*%s+.*"%word_guide["digraphs"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"digraphs"] = 1
            #Blends Check
            if match(".*%s+.*"%word_guide["blends"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"blends"] = 1
            #Common Long Vowels Check
            if match(".*%s+.*"%str(word_guide["common long vowels"][i]).replace("-","."), word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"common long vowels"] = 1
            #Other Vowels Check
            if match(".*%s+.*"%word_guide["other vowels"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"other vowels"] = 1
            #Inflected Endings Check
            if match(".*%s$"%word_guide["inflected endings"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"inflected endings"] = 1
            #Syllable Junctures Check
            if match(".*%s.*"%word_guide["syllable junctures"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"syllable junctures"] = 1
            #Unaccented Final Syllables Check
            if match(".*%s$"%word_guide["unaccented final syllables"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"unaccented final syllables"] = 1
            #Harder Suffixes Check
            if match(".*%s$"%word_guide["harder suffixes"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"harder suffixes"] = 1
            #Bases Or Roots Check
            if match(".*%s.*"%word_guide["bases or roots"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"bases or roots"] = 1
    elif level_selection == "upper-level":
        for i in range(len(word_guide)):
            #Blends and Digraphs Check
            if match(".*%s+.*"%word_guide["blends and digraphs"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"blends and digraphs"] = 1
            #Vowels Check
            if match(".*%s+.*"%str(word_guide["vowels"][i]).replace("-","."), word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"vowels"] = 1
            #Complex Consonants Check
            if match(".*%s+.*"%word_guide["complex consonants"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"complex consonants"] = 1
            #Inflected Endings And Syllable Juncture Check
            if match(".*%s+.*"%word_guide["inflected endings and syllable juncture"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"inflected endings and syllable juncture"] = 1
            #Unaccented Final Syllables Check
            if match(".*%s$"%word_guide["unaccented final syllables"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"unaccented final syllables"] = 1
            #Affixes Check
            #CHECK IF AFFIXES HAVE TO BE AT START/END
            if match(".*%s+.*"%word_guide["affixes"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"affixes"] = 1
            #Reduced Vowels in Unaccented Syllables Check
            if match(".*%s+.*"%word_guide["reduced vowels in unaccented syllables"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"reduced vowels in unaccented syllables"] = 1
            #Greek and Latin Elements Check
            if match(".*%s.*"%word_guide["greek and latin elements"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"greek and latin elements"] = 1
            #Assimilated Prefixes Check
            if match(".*%s.*"%word_guide["assimilated prefixes"][i], word_guide["test_words"][i]):
                feature_points += 1
                answer_sheet.loc[word_guide["feature"][i],"assimilated prefixes"] = 1     
    else:
        print("Incorrect level selection")
        return name, None, None
    #Set 0 for tested characteristics that are not 1
    fg_indexed = feature_guide.set_index("feature")
    for col in answer_sheet.columns:
        for row in answer_sheet.index:
            if answer_sheet.loc[row, col] != 1 and fg_indexed.loc[row, col] is not nan:
                answer_sheet.loc[row, col] = 0

    #Sum Feature Points By Words        
    answer_sheet["feature points"] = answer_sheet.sum(axis = 1)
    #Correct Spelling Bonus Point
    correct_words = list(word_guide["feature"] == word_guide["test_words"])
    correct_words = [int(x) for x in correct_words]

    answer_sheet["words spelled correctly"] = correct_words

    #Calculate spelling stage
    stage = check_spelling_stage(level_selection, feature_points)
    
    answer_sheet.insert(0, "spelled words", value = list(word_guide["test_words"]))
                                                  
    return name, stage, answer_sheet

#Load the feature guide to use for scoring
def get_feature_guide(selection):
    s = selection.lower()

    #change selection to string equivalent
    if s.isnumeric():
        menu_dict = {
        "1" : "primary",
        "2" : "elementary",
        "3" : "upper-level"
        }

        s = menu_dict[s]

    if s not in ["primary", "elementary", "upper-level"]:
        print("Valid options are primary, elementary, or upper-level")
        return None
    #Read in feature guide csv
    feature_guide = pd.read_csv(s + ".csv")
    feature_guide.fillna(0, inplace=True)
    #Set column names to lowercase
    feature_guide.columns = [str.lower(col) for col in feature_guide.columns]   
    return s, feature_guide

########################
### Main
########################

#GET LEVEL primary, elementary, ...

#GET FILE NAME AND PATH
while True:
    #Read in appropriate feature guide
    while True:
        s = input("""Enter the feature guide you want to use.
Valid options are:
    1) primary
    2) elementary
    3) upper-level
Choice: """)
        level_selection, feature_guide = get_feature_guide(s)
        if feature_guide is not None:
            break           

    while True:
        test_path = input("Enter the name of the file you would like to score or q to quit: ")
        
        if test_path == "q":
            break

        try :
            #Read in test
            test = pd.read_excel(test_path)
            test = test.rename(columns = {"name" : "feature"}).set_index('feature')
        except FileNotFoundError:
            print("File not found")
            continue

        break   

    
    if test_path == "q":
        break

    #Calculate Answer Sheets
    results = [calculate_feature_points(level_selection, feature_guide, test, name) for name in test.columns]
    
    #Check if results were returned for all names
    if any([True for r in results if r == False]):
        break
    #CREATE EXCEL SHEETS BY name AND SUMMARY PAGE WITH name AND stage
    #Create output path
    out_path = test_path.split(".")[0] + "_output.xlsx"
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
    #Create Summary Table
    names_stages = [[r[0], r[1]] for r in results]

    names_stages_df = pd.DataFrame(names_stages, columns = ["Name", "Learning Stage"])
    names_stages_df.to_excel(writer, sheet_name = "Summary", index = False)


    #Create Student Sheets
    for r in results:
        r[2].to_excel(writer, sheet_name = r[0], startrow = 2)
        worksheet = writer.book.get_worksheet_by_name(r[0])
        #Add name and spelling stage above answer sheet
        worksheet.write(0, 0, "Name")
        worksheet.write(0, 1, r[0])
        worksheet.write(1, 0, "Spelling Stage")
        worksheet.write(1, 1, r[1])
        #Apply conditionaly formatting to the answer sheet
        worksheet.conditional_format('C4:N28', {'type': 'icon_set',
                                       'icon_style': '3_symbols'})  

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    writer.close()
    #DISPLAY CONCLUDING MESSAGE STATING WHERE THE OUTPUT FILE IS SAVED TO
    print("Output has been saved to the workbook: %s"%out_path)

    while True:
        user_input = input("Would you like to score another workbook? (y/n): ")
        if user_input != "y" and user_input != "n":
            print("Invalid Input")
            continue
        break

    if user_input == "n":
        break




    