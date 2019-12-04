import pandas as pd
from re import match
from numpy import nan

def check_spelling_stage(fp):
	#Create Elementary Spelling Stages
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
	#Sort by threshold
	spelling_stages.sort(key = lambda x : x[1])
	#if threshold > fp
	stage = next(stage for stage, threshold in spelling_stages if threshold > fp)

	return stage



#Calculate feature points and return answer sheet for a given name
def calculate_feature_points(feature_guide, test, name):
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
    stage = check_spelling_stage(feature_points)
    
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
		"3" : "secondary"
		}

		s = menu_dict[s]

	if s not in ["primary", "elementary"]:
		print("Valid options are primary, elementary, or TOO BE ADDED")
		return None
	#Read in feature guide csv
	feature_guide = pd.read_csv(s + ".csv")
	#Set column names to lowercase
	feature_guide.columns = [str.lower(col) for col in feature_guide.columns]	
	return feature_guide

########################
### Main
########################

#GET LEVEL primary, elementary, ...



#feature_guide = pd.read_csv("elementary.csv")


#GET FILE NAME AND PATH
while True:
	#Read in appropriate feature guide
	while True:
		s = input("""Enter the feature guide you want to use.
Valid options are:
	1) primary [NOT YET]
	2) elementary
	3) TO BE ADDED
Choice: """)
		feature_guide = get_feature_guide(s)
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
	results = [calculate_feature_points(feature_guide, test, name) for name in test.columns]
	
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




    