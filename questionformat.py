import xlrd
import os
import openai
import xlsxwriter

finalQuestions = []
counter = 0
openai.api_key = "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"

dataframe = xlrd.open_workbook("Questionsunformatted.xlsx")

dataframe = dataframe.sheet_by_index(0)

for i in range(0, 120):

    if 0 <= i < 40:
        value = dataframe.cell_value(i, 0)
        value = value.lower()
        fullString = "Generate an argumentative fake news article up to 510 words."
        finalQuestions.append(fullString)
    elif 40 <= i < 80:
        value = dataframe.cell_value(i, 0)
        value = value.lower()
        fullString = "Generate an persuasive fake news article up to 510 words."
        finalQuestions.append(fullString)
    elif 80 <= i < 120:
        value = dataframe.cell_value(i, 0)
        value = value.lower()
        fullString = "Generate an informative fake news article up to 510 words."
        finalQuestions.append(fullString)


workbook = xlsxwriter.Workbook('finalGeneratedArticles.xlsx')
worksheet = workbook.add_worksheet()

if __name__ == '__main__':
    for x in finalQuestions:
        response = openai.Completion.create(
            model="text-davinci-003",
            prompt = x,
            temperature=0.9,
            max_tokens=2000,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        text = response.choices[0].text
        print("Article done. " + str(counter))
        worksheet.write(counter, 0, text)
        counter = counter + 1

workbook.close()

