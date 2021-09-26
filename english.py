
import pyautogui,time
import random
import webbrowser,sys,pandas as pd
from webutil import clear
opened=False
#Add your own path
path=r"C:\Users\Welcome\Desktop\English.xlsx"


def words(opened):
    chrome=webbrowser.get('"C:\Program Files\Google\Chrome\Application\chrome.exe" %s')
    clear.clear()
    # Checking if chrome is already opened
    if opened == False:
        pyautogui.moveTo(717,1079)
        time.sleep(2)
        # Chrome Location(Adjust accordingly)
        pyautogui.click(287,1057)
        time.sleep(2)
    
    with open(r"words.txt","r") as f:
        words=f.readlines()
        words=[word.replace("\n","") for word in words]
        chosen_words=random.sample(words,5)
        links=["https://www.merriam-webster.com/dictionary/"+word for word in chosen_words]
        print(chosen_words)
        new_wordsdf=pd.DataFrame(zip(chosen_words,links),columns=["Words","Links"])
        df=pd.read_excel(path)
        df=df.append(new_wordsdf,ignore_index=True)
        df=df.drop_duplicates()
        df.to_excel(path,index=False)
        for word in chosen_words:
            chrome.open("https://www.merriam-webster.com/dictionary/"+word)
    GoON(df)


# Asking if the user wants to see another batch of 5 words
def GoON(df):
    try:
        choice=input(f"Y:-Another 5 Words, N:- Exit\n").upper()
    except:
        print("Enter Y/N")
        time.sleep(3)
        GoON()
    if choice =="Y" or choice =="YES":
        opened=True
        words(opened)
        
    elif choice =="N" or choice =="NO":
        
        print("Exiting")
        time.sleep(3)
        sys.exit()
        
if __name__="__main__":
    words()
        
        
