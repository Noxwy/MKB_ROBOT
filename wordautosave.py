from pywinauto import Application
import time
import os

#Версия pywinauto работает с версией python 3.7.7+ / 3.8.2+


path = "C:\\..\\"#Путь к папке, в которой находятся финальные отчеты

def word_automate(list_file):   
    #Запуск приложения MS Word
    count = 1
    for item in list_file:
        app = Application(backend="uia").start(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")

        #Получение UI и взаимодейсвие
        #           Открытие документа
        app.Dialog.ListItem3.click_input()
        app.Dialog.BrowseButton.click_input()
        
        app.Dialog.Open.Edit49.type_keys(item)
        app.Dialog.Open.OpenSplitButton.click_input()

        #           Новый документ
        ##    app.Dialog.ListItem7.click_input()

        #Переход на новый интерфейс
        app = Application(backend="uia").connect(path="WINWORD.exe", title="Document1 - Word")
        app.Dialog.set_focus()
        
        #Нажимаем на клавишу "File"
        app.Dialog.FileTab.click()
        
        #Нажимаем на клавишу "Save As"
        app.Dialog.SaveAsListItem.click_input()

        #Нажимаем на клавишу "Browse"
        app.Dialog.BrowseButton.click_input()

        #Проверяем на существование файла, если True, то его удаляем
        if os.path.isfile(path+'file{i}'.format(i = str(count))+'.docx'):
            os.remove(path+'file{i}'.format(i = str(count))+'.docx')
            
        #Сохранениие докуента   
        Sub=app.Dialog.child_window(title_re="Save As", class_name="#32770")
        Sub.FileNameCombo.type_keys(path+'file{i}'.format(i = str(count)).replace('.docx',''))
        Sub.Save.click()
        try:
            app.Dialog.MicrosoftWord.OK.click()
        except:
            print('Word cохранил проект в более нововой версии!!!')
        while 1==1:
         if os.path.isfile(path+'file{i}'.format(i = str(count))+'.docx'):
            #Закрытие приложенния MS Word
            app.kill()
            break
         else:
             time.sleep(1)
        count+=1            

    
#---------Получение UI элементов
#app.Dialog.print_control_identifiers()#

