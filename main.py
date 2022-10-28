import asyncio
import panel as pn
import pandas as pd
import numpy as np
from panel.io.pyodide import show
from PyPDF2 import PdfReader
import re



file_input = pn.widgets.FileInput(accept='.pdf', multiple = True, width=180)
text_input = pn.widgets.TextInput(placeholder='Enter; your; keywords; seperated by; semicolon', value = "automation;consulting;ai;artificial intelligence;machine learning;strategy")
button_upload = pn.widgets.Button(name='Upload', button_type='primary', width=100)
row = pn.Row(file_input, text_input, button_upload, height=75)

table = pn.widgets.Tabulator(pagination='remote', page_size=25, header_filters = False, hierarchical  = False, editors = {},show_index=False, selectable = True)
document.getElementById('warning').style.display = 'none'

filename, button = table.download_menu(
    text_kwargs={'name': 'Enter filename', 'value': 'default.csv'},
    button_kwargs={'name': 'Download table'}
)

row2 = pn.Row(
    table
)

row3 = pn.Row(
    pn.Column(filename, button)
)


def process_file(event):
    document.getElementById('warning').style.display = 'none'
    # variables needed to build the pd.Series used to create the data frames at the end
    count = 0
    l_number_of_times_word_appeared = []
    l_keywords = []
    l_frequency = []
    l_page = []
    l_docName = []
    l_docname_filtered = []

    l_page_filtered = []
    l_keywords_filtered = []

    # List of terms that are filtered out
    daniel = ["automation","consulting","ai","artificial intelligence","machine learning","strategy"]
    dominik = ["consulting", "governance", "steering", "project management", "program management", "PMO"]
    #words_of_interest = pd.Series(dominik)

    if file_input.value is not None:
        words_of_interest = pd.Series(text_input.value.lower().split(";"))
        print(words_of_interest)
        file_n = 0
        for f in file_input.value:
            reader = PdfReader(io.BytesIO(f))
            count = 0
            for page_n in range(0,reader.getNumPages()):
                count += 1 #used to label the page number

                text_decoded = getDecodedText(reader, page_n) # extract text from page and decode it

                #extract each term out of the text
                keywords = re.findall(r'[a-zA-Z]\w+',text_decoded)

                keywords = parseMultiWordSearch(text_decoded, words_of_interest, keywords)

                #Create dataframe with the keywords
                df = pd.DataFrame(list(set(keywords)),columns=['keywords'])

                #Add the absolute frequency and the page number
                df['number_of_times_word_appeared'] = df['keywords'].apply(lambda x: weightage(x,text_decoded))
                df["page"] = str(count)
                df["docName"] = re.findall(r'[^\/]+?pdf$',file_input.filename[file_n])[0]

                #Extract the data into lists to create the overall extract
                l_number_of_times_word_appeared.extend(df['number_of_times_word_appeared'].tolist())
                l_keywords.extend(df['keywords'].tolist())
                l_page.extend(df['page'].tolist())
                l_docName.extend(df["docName"].tolist())

                #Create additional extract lists for the words of interest
                for word in words_of_interest:
                    if re.search(r"\b" + re.escape(word) + r"\b", text_decoded):
                        l_page_filtered.append(str(count))
                        l_keywords_filtered.append(word)
                        l_docname_filtered.append(re.findall(r'[^\/]+?pdf$',file_input.filename[file_n])[0])
            file_n += 1
                #table.value = df
                #document.getElementById('table').style.display = 'block'
        #Assemble lists to create the final data frame and save it as a csv
        df_final = pd.DataFrame(list(zip(l_docName, l_page,l_keywords,l_number_of_times_word_appeared)), columns = ["docName","page_number","keywords","frequency_abs"])

        #Create Pivot Table with words of interest
        df_final["page_number"].astype(int,copy=False)
        df_final["keywords"].astype(str,copy=False)

        if df_final["keywords"].isin(words_of_interest).sum() != 0:
            pivot = df_final[df_final["keywords"].isin(words_of_interest)].pivot_table(index=["docName","page_number"],columns="keywords",fill_value=0,sort=False,margins=[True,False],aggfunc="sum").iloc[:-1,:].sort_values(by=("frequency_abs","All"),ascending=False)
            pivot.columns = pivot.columns.droplevel(level=0)
            table.value = pivot.reset_index().astype(str)
            print(pivot)

        else:
            document.getElementById('warning').textContent = 'No matches found'
            document.getElementById('warning').style.display = 'block'

button_upload.on_click(process_file)

await show(row, 'fileinput')
await show(row2, 'table')
await show(row3, 'dl')


def getDecodedText(reader, page_n):
    text = ""
    text = reader.pages[page_n].extract_text()
    # encode and make text lowercase to enable matching
    text = text.encode('ascii','ignore').lower()
    return text.decode()

def parseMultiWordSearch(text, words_of_interest, keywords):
    for word in words_of_interest:
        if re.findall(r" ", word):
            if re.findall(r"\b" + re.escape(word) + r"\b", text):
                matches = re.findall(r"\b" + re.escape(word) + r"\b", text)
                for k in matches:
                    keywords.append(k)
    return keywords

def weightage(word,text,number_of_documents=1):
    word_list = re.findall(r"\b" +word+r"\b" ,text)
    number_of_times_word_appeared =len(word_list)
    tf = number_of_times_word_appeared/float(len(text))
    idf = 0 #np.log((number_of_documents)/float(number_of_times_word_appeared))
    tf_idf = tf*idf
    return number_of_times_word_appeared 