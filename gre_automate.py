import pickle
from slacker import Slacker
import pandas as pd
import xlrd


saivishToken='<ENTER API TOKEN HERE>'
slackobj=Slacker(saivishToken)
channel_name='vocab-build'

def get_var(filename):    
    f=open(filename,'rb')
    var=pickle.load(f)
    f.close()
    return var

def dump_var(var,filename):
    f=open(filename,'wb')
    pickle.dump(var,f)
    f.close()

def get_sheet():
    current_index=get_var('index.p')
    current_sheet=get_var('sheet.p')
    if current_index>25:
        current_sheet+=1
        dump_var(current_index,'sheet.p')
    return current_sheet

def summary():
    word_list=get_var('word_list.p')
    msg='*WORDS FROM TODAY*\n'
    i=0
    for word in word_list:
        msg+='\n'+str(i+1)+'. '+word
        i+=1
    slackobj.chat.post_message('#' + channel_name,msg, as_user=False)
    word_list=[]
    current_word=get_var('cur_word.p')
    current_word=1
    dump_var(current_word,'cur_word.p')
    


def post_word(): 
    current_sheet=get_sheet()
    if current_sheet>48:
        slackobj.chat.post_message('#' + channel_name, "CONGRATS. YOU GUYS HAVE REACHED THE END.\n\n\n\n https://youtu.be/dQw4w9WgXcQ ", as_user=False)
    xlsheet=pd.read_excel('Words with mnemonics (1-48).xlsx',sheet_name='day '+str(current_sheet),index_col=0,dtype={'WORDS': str, 'DEFINITION': str,'MNEMONIC':str,'SYNONYMS':str,'PHONETIC':str})
    if xlsheet.isnull().values.any():
        xlsheet.dropna(inplace=True)
    col=['WORDS', 'DEFINITION','MNEMONIC','SYNONYMS','PHONETIC']
    xlsheet.columns=col
    xlsheet=xlsheet[1:]
    current_word=get_var('cur_word.p')
    current_index=get_var('index.p')
    msg=str(current_word)+'.\n\n *_'+xlsheet[col[0]][current_index].upper()+'_*\n\n  '+xlsheet[col[1]][current_index]+'\n\n *SYNONYMS* : '+xlsheet[col[3]][current_index]+'\n\n *MNEMONIC* : '+xlsheet[col[2]][current_index]
    slackobj.chat.post_message('#' + channel_name, msg, as_user=False)
    word_list=get_var('word_list.p')
    word_list.append(xlsheet[col[0]][current_index])
    current_index+=1
    current_word+=1
    dump_var(current_index,'index.p')
    dump_var(current_word,'cur_word.p')
    dump_var(word_list,'word_list.p')
    

if __name__ == "__main__":
    post_word()
