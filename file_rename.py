import tkinter
import tkinter.messagebox
import os


#決定ボタンを押したとき
def click_btn():
    #テキスト入力値の取得
    input_path = entry_path.get()
    input_number = entry_number.get()
    input_subject = entry_subject.get()
    input_extention = entry_extention.get()
    #パスが存在しないときの処理
    if os.path.exists(input_path)==True:
        #アルゴリズムの実行
        rename_tool(input_path,input_number,input_subject,input_extention)
        #完了したらメッセージボックス表示
        tkinter.messagebox.showinfo("imformation","Excuted !")
        #path,subject,extentionの入力値初期化
        entry_path.delete(0,tkinter.END)
        entry_subject.delete(0,tkinter.END)
        entry_extention.delete(0,tkinter.END)
    else:
        tkinter.messagebox.showinfo("imformation","[win Error3] Path dosen't exist")


#ファイル名アルゴリズム
def rename_tool(path,number,subject,extention):
    list = os.listdir(path)
    name_num=1
    #拡張子が"ALL"のときはすべてのファイルをrename
    if extention =="all":
        for i in range(len(list)):
            unneccessary_path,read_extention=os.path.splitext(path+list[i])
            #フォルダ名は変更しない処理
            if read_extention!="":
                new_file_name = "{}_{}_第{}回{}" .format(number,subject,name_num,read_extention)
                os.rename(os.path.join(path,list[i]),os.path.join(path,new_file_name))
                name_num += 1
    #拡張子を指定するとき
    else:
        for i in range(len(list)):
            unneccessary_path,read_extention=os.path.splitext(path+list[i])
            if read_extention == extention:
                new_file_name = "{}_{}_第{}回{}" .format(number,subject,name_num,read_extention)
                os.rename(os.path.join(path,list[i]),os.path.join(path,new_file_name))
                name_num += 1




#################################################################
#TkinterToolの処理
#################################################################

TITLE_GRAPHIC_SIZE = 45

INPUT_GRAPHIC_SIZE = 23
INPUT_EX_GRAPHIC_SIZE = 10

ENTRY_WIDTH = 40
ENTRY_GRAPHIC_SIZE = 3


#ウィンドウの配置
root = tkinter.Tk()
root.title("File Rename tool")
root.geometry("800x600")


#文字の配置
label_title = tkinter.Label(root,text="File rename tool",font=("System",TITLE_GRAPHIC_SIZE))
label_title.place(x=240,y=40)

label_path = tkinter.Label(root,text="Enter the path",font=("System",INPUT_GRAPHIC_SIZE))
label_path.place(x=100,y=150)
label_path_ex = tkinter.Label(root,text="ex : C:\Program Files\Google",font=("System",INPUT_EX_GRAPHIC_SIZE))
label_path_ex.place(x=140,y=180)

label_number = tkinter.Label(root,text="Enter the student number",font=("System",INPUT_GRAPHIC_SIZE))
label_number.place(x=100,y=230)

label_subject = tkinter.Label(root,text="Enter the subject",font=("System",INPUT_GRAPHIC_SIZE))
label_subject.place(x=100,y=310)

label_extention = tkinter.Label(root,text="Enter the extention",font=("System",INPUT_GRAPHIC_SIZE))
label_extention.place(x=100,y=390)
label_extention_ex = tkinter.Label(root,text="ex : [.pdf] or ALLfile : [all]",font=("System",INPUT_EX_GRAPHIC_SIZE))
label_extention_ex.place(x=140,y=420)

label_warning = tkinter.Label(root,text="Backup and run",font=("System",8))
label_warning.place(x=340,y=480)


#テキスト入力欄の配置
entry_path = tkinter.Entry(width=ENTRY_WIDTH,font=("System",ENTRY_GRAPHIC_SIZE))
entry_path.place(x=400,y=150)

entry_number = tkinter.Entry(width=ENTRY_WIDTH,font=("System",ENTRY_GRAPHIC_SIZE))
entry_number.place(x=400,y=230)

entry_subject = tkinter.Entry(width=ENTRY_WIDTH,font=("System",ENTRY_GRAPHIC_SIZE))
entry_subject.place(x=400,y=310)

entry_extention = tkinter.Entry(width=ENTRY_WIDTH,font=("System",ENTRY_GRAPHIC_SIZE))
entry_extention.place(x=400,y=390)


#チェックボタンの配置
#button_check = tkinter.Checkbutton(text="Backups have been made.")
#button_check.place(x=320,y=480)


#ボタンの配置
button_decision = tkinter.Button(root,text="Rename",font=("System,40"),command=click_btn)
button_decision.place(x=350,y=510)


root.mainloop()
