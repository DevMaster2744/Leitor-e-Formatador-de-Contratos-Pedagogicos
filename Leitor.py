from tabula import read_pdf
from tabula import convert_into
from io import StringIO 
import json 
from openpyxl import Workbook, load_workbook
import json

tabela = []

#convert_into("Contratos Pedagógicos\Contratos_Pedagógicos_82\CONTRATO I 82 MAT 2024.pdf", "data.json", output_format="json")

with open("main_data.json", "r+") as f:
    saved_data = json.load(f)
    save_file = saved_data['data_file']

def read_pdf_to(retry=False, lattice_mode=True, file=""):
    df = read_pdf(file, pages="all", silent=True, output_format="json", force_subprocess=True,
                  relative_columns=True, relative_area=True)

    print(df)
    subject_search_finished = False
    subject = "None"

    date_header = 0
    date_search_finished = False

    name_header = 0
    name_search_finished = False

    value_header = 0
    value_search_finished = False

    header_index = 0

    is_data = False

    for i_index, i in enumerate(df):
        # if i['extraction_method'] == "stream":
        size = len(i)
        data = i['data']

        for row_index, row in enumerate(data):
            for colum_index, colum in enumerate(row):
                data_text = colum['text']  # linha

                date_index = data_text.lower().find('data')
                if not subject_search_finished:
                    with open("main_data.json", "r+") as f:
                        data_json = json.load(f)
                        for subject_json in data_json['subjects']:
                            subject_index = unidecode.unidecode(data_text.lower()).find(subject_json['id'])
                            if subject_index >= 0:
                                print(f"Encontramos a matéria!")
                                subject = subject_json['name']
                                subject_search_finished = True

                if date_index >= 0 and not date_search_finished:
                    print(f"Encontramos a coluna de datas!")
                    date_header = colum_index
                    date_search_finished = True

                name_index1 = data_text.lower().find('conteúdos')
                name_index2 = data_text.lower().find('instrumento')
                name_index3 = data_text.lower().find('avaliação')
                if name_index1 >= 0 or name_index2 >= 0 or name_index3 >= 0 and not name_search_finished:
                    print(f"Encontramos a coluna de nomes dos trabalhos!")
                    name_header = colum_index
                    name_search_finished = True

                value_index1 = data_text.lower().find('nota')
                value_index2 = data_text.lower().find('peso')
                if value_index1 >= 0 or value_index2 >= 0 and not value_search_finished:
                    print(f"Encontramos a coluna de valores(pontos)!")
                    value_header = colum_index
                    value_search_finished = True
                if date_search_finished and name_search_finished and value_search_finished and subject_search_finished:
                    if is_data:
                        if row_index >= header_index - 1:
                            #if is_data:
                                if data_text != "" and colum_index == value_header:
                                    tabela.append({"Data": "-----", "Nome": "", "Valor": data_text, "Matéria": "-----"})

                                #name = str(i.iloc[row, name_header[0]])
                                if data_text != "" and len(tabela) > 0 and colum_index == name_header:
                                    tabela[-1]['Nome'] += data_text + "\n"

                                #date = str(i.iloc[row, date_header[0]])
                                if data_text != "" and len(tabela) > 0 and colum_index == date_header:
                                    tabela[-1]['Data'] = data_text
                                    if tabela[-1]['Data'] == "-----":
                                        tabela[-1]['Data'] = ""
                                    sep = tabela[-1]['Data'].find(" ")
                                    if sep >= 0:
                                        split = tabela[-1]['Data'].split(" ")
                                        tabela[-1]['Data'] = split[-1]
                                if len(tabela) > 0:
                                    tabela[-1]['Matéria'] = subject
                    else:
                        is_data = True
                elif date_search_finished and name_search_finished and value_search_finished:
                    if is_data:
                        if row_index >= header_index - 1:
                            #if is_data:
                                if data_text != "" and colum_index == value_header:
                                    tabela.append({"Data": "-----", "Nome": "", "Valor": data_text, "Matéria": "-----"})

                                # name = str(i.iloc[row, name_header[0]])
                                if data_text != "" and len(tabela) > 0 and colum_index == name_header:
                                    tabela[-1]['Nome'] += data_text + "\n"

                                # date = str(i.iloc[row, date_header[0]])
                                if data_text != "" and len(tabela) > 0 and colum_index == date_header:
                                    tabela[-1]['Data'] = data_text
                                    if tabela[-1]['Data'] == "-----":
                                        tabela[-1]['Data'] = ""
                                    sep = tabela[-1]['Data'].find(" ")
                                    if sep >= 0:
                                        split = tabela[-1]['Data'].split(" ")
                                        tabela[-1]['Data'] = split[-1]
                                if len(tabela) > 0:
                                    tabela[-1]['Matéria'] = subject
                    else:
                        is_data = True
    print("Terminando...")

    load_data()


def load_data():
    global tabela

    workbook = load_workbook("Lista de trabalhos.xlsx")

    page = workbook.active

    tabela = sorted(tabela, reverse=False, key= lambda x: (int(x['Data'].split("/")[0]) + (int(x['Data'].split("/")[1]) * 30) 
                                                if str.isnumeric(x['Data'].split("/")[0]) and str.isnumeric(x['Data'].split("/")[1]) else False) 
                                                if len(x['Data'].split("/")) > 1 and len(x['Data'].split("/")) < 3 else False)
    print(f"Tabela: {tabela}")
    if len(tabela) > 0:
        for i in range(len(tabela)):
            page[f"B{i + 4}"] = tabela[i]['Data']
            page[f"C{i + 4}"] = tabela[i]['Matéria']
            page[f"D{i + 4}"] = tabela[i]['Nome']
            page[f"E{i + 4}"] = tabela[i]['Valor']

        workbook.save("Lista de trabalhos.xlsx")

        with open(save_file, "r+") as f:
            saved_data = json.load(f)

            saved_data['tabela'] = tabela

            f.seek(0)        # <--- should reset file position to the beginning.
            json.dump(saved_data, f, indent=4)
            f.truncate() 

        print("Acabado!")

        print(tabela)
    else:
        print("Erro ao salvar!")

while True:
    input_text = str(input())

    if input_text.upper() == "READ":
        
        print("Qual é a matéria?")

        matéria = str(input())
       
        print("Qual é o local do arquivo(local relativo)?")

        file = str(input())

        df = read_pdf(file, output_format="json", pages="all", silent=True)

        for i in df:
            if i['extraction_method'] == "stream":
                size = i['data'].__len__()

                date_header = 0
                date_search_finished = False

                name_header = 0
                name_search_finished = False

                value_header = 0
                value_search_finished = False
                
                header_index = 0

                is_data = False

                for data_list_index in range(i['data'].__len__()):
                    data_list = i['data'][data_list_index] # linha
                    for data_index in range(data_list.__len__()):
                        if value_search_finished and name_search_finished and date_search_finished:
                            break

                        data = data_list[data_index]

                        date_index = str(data['text']).lower().find('data')

                        if date_index >= 0:
                            print(f"Encontramos a coluna de datas! index: {data_index}")
                            date_header = data_index
                            date_search_finished = True

                        name_index1 = str(data['text']).lower().find('conteúdos')
                        name_index2 = str(data['text']).lower().find('instrumento')
                        name_index3 = str(data['text']).lower().find('avaliação')
                        if name_index1 >= 0 or name_index2 >= 0 or name_index3 >= 0:
                            print(f"Encontramos a coluna de nomes dos trabalhos! index: {data_index}")
                            name_header = data_index
                            name_search_finished = True

                        value_index1 = str(data['text']).lower().find('nota')
                        value_index2 = str(data['text']).lower().find('peso')
                        if value_index1 >= 0 or value_index2 >= 0:
                            print(f"Encontramos a coluna de valores(pontos)! index: {data_index}")
                            value_header = data_index
                            value_search_finished = True

                        if value_search_finished and name_search_finished and date_search_finished:
                            header_index = data_index

                    if date_search_finished and name_search_finished and value_search_finished:
                        if is_data:
                            if data_list[value_header]['text'] != "":
                                tabela.append({"Data": "ERRO", "Nome": "", "Valor": data_list[value_header]['text'], "Matéria": "-----"})

                            if data_list[name_header]['text'] != "" and len(tabela) > 0:
                                tabela[-1]['Nome'] += data_list[name_header]['text'] + "\n"

                            if data_list[date_header]['text'] != "" and len(tabela) > 0:
                                tabela[-1]['Data'] = data_list[date_header]['text']
                                sep = tabela[-1]['Data'].find(" ")
                                if sep >= 0:
                                    split = tabela[-1]['Data'].split(" ")
                                    tabela[-1]['Data'] = split[-1]
                            if len(tabela) > 0:
                                tabela[-1]['Matéria'] = matéria

                            '''trabalhos.append({"Data": str(data_list[date_header]['text']),
                                            "Nome": str(data_list[name_header]['text']),
                                            "Valor": str(data_list[value_header]['text'])})'''
                        # depois que ele reconhece que as informações não são mais os "headers" ele começa a gravar os dados
                        is_data = True
                    else:
                        header_index += 1

        print("Terminando...")

        load_data()
        
    elif input_text.upper() == "LOAD":
        print(save_file)
        with open(save_file, "r+") as f:
            saved_data = json.load(f)
            tabela = saved_data['tabela']
            
        load_data()

    elif input_text.upper() == "DELETE DATA":
        with open(save_file, "r+") as f:
            saved_data = json.load(f)

            saved_data['tabela'] = []

            f.seek(0)        # <--- should reset file position to the beginning.
            json.dump(saved_data, f, indent=4)
            f.truncate()

    elif input_text.upper() == "MAKE SAVE":
        print("Qual será o nome do arquivo?")

        input_name = str(input())

        with open(input_name + ".json", "w") as f:
            f.write('{\n"tabela": []\n}')
    elif input_text.upper() == "LOAD SAVE":
        print("Qual é o caminho para este arquivo(caminho relativo)?")

        input_path = str(input())

        save_file = input_path

        with open("main_data.json", "r+") as f:
            saved_data = json.load(f)

            saved_data['data_file'] = input_path

            f.seek(0)        # <--- should reset file position to the beginning.
            json.dump(saved_data, f, indent=4)
            f.truncate() 
    elif input_text.upper() == "INSERT":
        print("Qual a data?")

        input_date = str(input())

        print("Qual a descrição?")

        input_desc = str(input())

        print("Qual o valor(pontos)?")

        input_value = str(input())

        print("Matéria?")

        input_class = str(input())

        tabela.append({"Data": input_date, "Nome": input_desc, "Valor": input_value, "Matéria": input_class})

        load_data()
