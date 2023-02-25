##import packages
import xlrd
import os
import openai
import xlsxwriter

##setup initial variables for storage and counting
finalQuestions = []
counter = 0
##load in openai api key for shared account
openai.api_key = "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"

##for statement to iterate through a range of 0 - 120
for i in range(0, 120):

    ##first 40 questions will include the argumentative ones
    if 0 <= i < 40:
        value = dataframe.cell_value(i, 0)
        value = value.lower()
        fullString = "Generate an argumentative fake news article up to 510 words."
        finalQuestions.append(fullString)
    ##second 40 questions will include the persuasive ones
    elif 40 <= i < 80:
        value = dataframe.cell_value(i, 0)
        value = value.lower()
        fullString = "Generate an persuasive fake news article up to 510 words."
        finalQuestions.append(fullString)
    ##third 40 questions will include the informative ones
    elif 80 <= i < 120:
        value = dataframe.cell_value(i, 0)
        value = value.lower()
        fullString = "Generate an informative fake news article up to 510 words."
        finalQuestions.append(fullString)

##add a new workbook to write the final articles into
workbook = xlsxwriter.Workbook('finalGeneratedArticles.xlsx')
worksheet = workbook.add_worksheet()

if __name__ == '__main__':
    ##for each question in the questions array 
    for x in finalQuestions:
        ##create a open ai completion response
        response = openai.Completion.create(
            ##using model davinci-003
            model="text-davinci-003",
            ##prompt being the question
            prompt = x,
            temperature=0.9,
            ##max tokens will be set to 2000 to not interrupt the ai creation
            max_tokens=2000,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        ##choose the text only in the response
        text = response.choices[0].text
        ##print the article number and done to keep track of progress
        print("Article done. " + str(counter))
        ##write to the worksheet including the proper counter number for row
        worksheet.write(counter, 0, text)
        ##increment counter
        counter = counter + 1

##close the workbook and commit changes
workbook.close()

