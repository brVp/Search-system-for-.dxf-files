# -*- coding: utf-8 -*-

import os
import pandas as pd
import ezdxf
import datetime

#Start

debug = False

file_name = "DB.xlsx"
path = os.path.dirname(os.path.abspath("Notebook.ipynb")) #Path folder

if debug == True: path = r"Fixed path" #Debug

#Functions for search

def search_value(val,df_ref,col,search_list):
    
    if not df_ref.loc[df_ref[col].str.contains(val, case=False,na=False)].empty:    
        search_list.append(col[:-4]) # [:-4] remove file extension

def list2text(list_):
    text = ''
    for item in list_:
        text = text + item + '\n' 
    return text

def search(value2search,df_ref):
    search_keys = []
    for column_name in list(df_ref.columns.values):
        #print(column_name)
        search_value(value2search,df_ref,column_name,search_keys)
   
    if not search_keys:
        search_keys.append('Não foi encontrado nenhum resultado para o valor buscado')
        
    return list2text(search_keys)

def create_database():
    
    key_words = []

    df = pd.DataFrame() #df with all data

    final_ml_text= ''
    #main_win['search_keys_str'].update('final_ml_text')
    #main_win.refresh()

    for file in os.listdir(path):
        if file[-4:] == '.dxf':
            #print(file)
            
            final_ml_text= final_ml_text + file + '\n'
            main_win['search_keys_str'].update(final_ml_text)
            main_win.refresh()
            
            get_keywords(file,path + "\\",key_words)
            #print(key_words)
            #textfile = open(key_words[0][:-4] + ".txt", "w")
            #for element in key_words:
            #    textfile.write(element + "\n")
            #textfile.close()
            #df[key_words[0]] = pd.Series(key_words[1:])
            df2 = pd.DataFrame(key_words[1:], columns=[key_words[0]])
            #print(df2)
            df = pd.concat([df, df2], axis=1)
            df = limpa_df(df)
            key_words = []

    #Save as xlsx

    #df = limpa_df(df)

    file_name = "DB.xlsx"

    df.to_excel(path + "//" + file_name, index=False, encoding='utf-8-sig')

    #print("Banco de dados salvo %s" % path + "\\" + file_name)
    #input("Execução finalizada!")

#functions for CAD / keywords
   
def get_block(ent,list):
    
    for entity in ent.virtual_entities():

        #print("Entity type: %s" % entity.dxftype()) #Entity name, print only for debug

        if entity.dxftype() == 'TEXT':
            #print("TEXT Text: %s" % entity.dxf.text)
            list.append(entity.dxf.text)

        if entity.dxftype() == 'MTEXT':
            #print("MTEXT Text: %s" % entity.text)
            list.append(entity.text)
            
        if entity.dxftype() == 'INSERT':
            get_block(entity, list)

    if ent.attribs:

        for att in ent.attribs:
            #print("Att Tag: %s" % att.dxf.tag)
            #print("Att Text: %s" % att.dxf.text) #When attribute is in block, is necessary to use text instead the TAG
            list.append(att.dxf.text)

    #else:
        #print("block has not attributes") #For debug only
        
def get_keywords(file_name,path,key_words):
    
    doc = ezdxf.readfile(path + file_name) #Read File
    # iterate over all entities in modelspace
    msp = doc.modelspace()

    key_words.append(file_name)

    for ent in msp:

        #print("Entity type: %s" % ent.dxftype()) #Entity name, print only for debug

        if ent.dxftype() == 'ATTDEF':
            #print(ent.dxf.prompt)
            #print(ent.dxf.text)
            #print("Att F Text: %s" % ent.dxf.tag) #When exists a attribute used as text is better use the TAG for extract text
            key_words.append(ent.dxf.tag)
            #print("\n")

        if ent.dxftype() == 'TEXT':
            #print("TEXT Text: %s" % ent.dxf.text)
            key_words.append(ent.dxf.text)
            #print("\n")

        if ent.dxftype() == 'MTEXT':
            #print("MTEXT Text: %s" % ent.text)
            key_words.append(ent.text)
            #print("\n")

        if ent.dxftype() == 'INSERT':
            get_block(ent,key_words)

def limpa_df(df_ref):
    df_ref  = df_ref.fillna("")
    df_ref  = df_ref.replace(r'\n',' ', regex=True)
    df_ref  = df_ref.replace('%%d',' ', regex=True)
    df_ref  = df_ref.replace('%%',' ', regex=True)
    df_ref  = df_ref[df_ref.apply(lambda col: col.map(lambda x: True if len(str(x)) > 2 else False))] #Remover valores com 2 chars
    df_ref  = df_ref.apply(lambda x: x.sort_values().values) #Mandar os nans para o final
    df_ref  = df_ref.apply(lambda col: col.drop_duplicates().reset_index(drop=True)) #Remover duplicados
    return df_ref

def att_database():
    main_win['title_mline'].update('Desenhos carregados para o banco de dados:') 
    main_win['search_keys_str'].update('')
    main_win['search_value'].update('')
    main_win.refresh()
    change_status('Atualizando banco de dados (esse processo pode demorar alguns minutos)...')
    create_database() #Create new data base
    change_status('Carregando banco de dados...')       
    df = pd.read_excel(path + "//" + file_name, index_col=None,na_values=" ")
    df = df.applymap(str) #Convert all df to str
    change_status('Banco de dados carregado...')
    att_date = modification_date(path + "//" + file_name)
    main_win['att_date_key'].update('Última atualização: ' + att_date)
    main_win.refresh()
    return df

def modification_date(filename):
    t = os.path.getmtime(filename)
    return str(datetime.datetime.fromtimestamp(t).strftime("%Y-%m-%d %H:%M:%S"))

def change_status(att_status):
    main_win['status'].update('Status: ' + att_status)
    #print(att_status)
    main_win.refresh()

#GUI

from PySimpleGUI import PySimpleGUI as sg

sg.theme('Reddit')

att_status = 'Iniciado...'

layout_search = [
    [sg.Text('Status: ' + att_status,key='status')],
    [sg.Text('')],
    [sg.Text('Buscar texto:'), sg.Input(key='search_value',size=(67,1)),sg.Button('Pesquisar')],
    [sg.Text('')],
    [sg.Text('Resultado da busca:',key='title_mline')],
    [sg.Multiline('', key='search_keys_str',size=(90,8))],
    [sg.Text('')],
    [sg.Button('Atualizar banco de dados')],
    [sg.Text('')],
    [sg.Text('Dados: ' + path + "\\" + file_name)],
    [sg.Text('Última atualização: ',key='att_date_key')]
]

main_win = sg.Window('Pesquisar',layout_search,finalize=True)

while True:
    try:
        change_status('Carregando banco de dados...')
        df = pd.read_excel(path + "//" + file_name, index_col=None,na_values=" ")
        df = df.applymap(str) #Convert all df to str
        att_date = modification_date(path + "//" + file_name)
        main_win['att_date_key'].update('Última atualização: ' + att_date)
        main_win.refresh()
        #df = df.fillna(" ")
        change_status('Banco de dados carregado...')
        break
    except:
        #print('Não encontrado')
        df = att_database()
        break

while True:
    eventos,valores = main_win.read()
    change_status('Disponível')
        
    if eventos == sg.WIN_CLOSED:
        break
                
    if eventos == 'Pesquisar':
        
        change_status('Pesquisando...')
        main_win['title_mline'].update('Resultado da busca:')
        main_win.refresh()
        
        if valores['search_value']:
            main_win['search_keys_str'].update(search(valores['search_value'],df))
        else:
            main_win['search_keys_str'].update('Busca vazia!')
            
        change_status('Aguardando nova busca...')
        
    if eventos == 'Atualizar banco de dados':
        df = att_database()
        
        
    #print(valores['search_keys_str'])
    #main_win['search_keys_str'].update('')
    #main_win.refresh()