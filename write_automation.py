import openpyxl
import logging
import os
from datetime import datetime, timedelta



#--------------------------------各自変更箇所------------------------------------
#ファイル名の設定
fileName = "月毎研究時間_月_Takamatsu.xlsx"

#xlsxファイルが存在する場所の絶対パス_先頭にrを付けないとエラーが出ます
filePath = r""

#開始時刻のdelay[minutes]
delayStart = -10
#終了時刻のdelay[minutes]
delayFinish = 10

#ログが不要の場合以下のコードを有効に
#logging.disable(logging.CRITICAL)
#----------------------------------ここまで-------------------------------------




#ファイルがある場所のパスの取得
#filePath = os.getcwd()

#ログの出力
logging.basicConfig(filename = (os.path.join(filePath, "write_automation_log.txt")), filemode = "a", level = logging.INFO, format = " %(asctime)s -%(levelname)s- %(message)s ")



def main():
    while True:
        month, day, hour, minutes = datetimeGet(0)
        print("ctrl+cでこのツールを終了することができます")
        print("固まったら任意のキーを押すことで動作するかも")
        print(fileName + "を開いています…\n")

        #xlsxファイルと当月シートが開けた
        if excelOpen(month):
            logging.info(fileName + " の " + sheetName + " シートがユーザによって開かれました")
            #-------------ここがメインの処理-----------------------
            while True:
                process()
            #---------------------------------------------------

        #xlsxファイルと当月シートが開けなかった
        else:
            logging.error("ファイルもしくはシートを開くことができませんでした")
            print("ERROR : ファイルを閉じて再度実行してください")
            input("        任意のキーを入力することで再実行できます ")
            print("\n\n")

#処理を行う
def process():
    contentsIn = input("1/開始時刻 2/終了時刻 3/休憩時間 4/就活時間 5/SAと授業時間 : ")
    #数値が入力されたとき
    try:
        contentsNum = int(contentsIn)
    #数値以外が入力されたとき
    except:
        contentsNum = 99
    
    #1:開始時刻の書き込み
    if (contentsNum == 1):
        month, day, hour, minutes = datetimeGet(delayStart)
        excelWrite_start(day, hour, minutes)
    #2:終了時刻の書き込み
    elif (contentsNum == 2):
        month, day, hour, minutes = datetimeGet(delayFinish)
        excelWrite_finish(day, hour, minutes)
    #3:休憩時間の書き込み
    elif (contentsNum == 3):
        month, day, hour, minutes = datetimeGet(0)
        excelWrite_relax(day)
    #4:就活時間の書き込み
    elif (contentsNum == 4):
        month, day, hour, minutes = datetimeGet(0)
        excelWrite_Jobhunt(day)
    #5:SA,授業時間の書き込み
    elif (contentsNum == 5):
        month, day, hour, minutes = datetimeGet(0)
        excelWrite_Teach(day)
    #1-5以外が入力されたとき
    else:
        print("WARNING : 1-5の数値で入力してください\n")


#現在の月,日,時,分を返す関数(delayも計算)
def datetimeGet(delay):
    if (delay > 0):
        now = datetime.now() + timedelta(minutes = delay)
    elif (delay < 0):
        now = datetime.now() - timedelta(minutes = (abs(delay)))
    else:
        now = datetime.now()
    return now.month, now.day, now.hour, now.minute

#excelを開いて当月のシートに移動
def excelOpen(sheetMonth):
    global excelBook, excelData, sheetName    
    try:
        sheetName = str(sheetMonth) + "月"
        excelData = openpyxl.load_workbook(filePath + '\\' + fileName)
        excelBook = excelData[sheetName]
        return True
    except:
        return False




#---------------------------書き込み処理----------------------------------------
#開始時刻の書き込み
def excelWrite_start(day, startHour, startMinutes):
    writeCell = "W" + str(day + 5)
    writeValue = str(startHour) + ":" + "{:02d}".format(startMinutes)
    #書き込みを実行
    if excelWrite_overwrite(writeCell):
        excelBook[writeCell] = writeValue
        excelData.save(os.path.join(filePath, fileName))
        print("開始時刻 " + str(writeValue) + " で書き込み完了\n")
        logging.info(writeCell + "のセルに開始時刻 " + str(writeValue) + " を書き込み成功")
    #書き込みを中断
    else:
        print("処理を中断しました\n")
        logging.info(writeCell + "のセルへの開始時間 " + str(writeValue) + " への書き込みを中断")

#終了時刻の書き込み
def excelWrite_finish(day, finishHour, finishMinutes):
    writeCell = "AG" + str(day + 5)
    writeValue = str(finishHour) + ":" + "{:02d}".format(finishMinutes)
    #書き込みを実行
    if excelWrite_overwrite(writeCell):
        excelBook[writeCell] = writeValue
        excelData.save(os.path.join(filePath, fileName))
        print("終了時刻 " + str(writeValue) + " で書き込み完了")
        print("おつかれさまでした！！\n")
        logging.info(writeCell + "のセルに終了時刻 " + str(writeValue) + " を書き込み成功")
    #書き込みを中断
    else:
        print("処理を中断しました\n")
        logging.info(writeCell + "のセルへの終了時間 " + str(writeValue) + " への書き込みを中断")

#休憩時間の書き込み
def excelWrite_relax(day):
    writeCell = "DK" + str(day + 5)
    while True:
        timeRerax = input("休憩時間を入力してください : ")
        try:
            writeValue = float(timeRerax)
            break
        except:
            print("WARNING : 数値で入力してください")
    #書き込みを実行
    if excelWrite_overwrite(writeCell):
        excelBook[writeCell] = writeValue
        excelData.save(os.path.join(filePath, fileName))
        print("休憩時間 " + str(writeValue) + "[h] で書き込み完了\n")
        logging.info(writeCell + "のセルに休憩時間 " + str(writeValue) + " を書き込み成功")
    #書き込みを中断
    else:
        print("処理を中断しました\n")
        logging.info(writeCell + "のセルへの休憩時間 " + str(writeValue) + " への書き込みを中断")

#就活時間の書き込み
def excelWrite_Jobhunt(day):
    writeCell = "BW" + str(day + 5)
    while True:
        timeJobhunt = input("就活した時間を入力してください : ")
        try:
            writeValue = float(timeJobhunt)
            break
        except:
            print("WARNING : 数値で入力してください")
    #書き込みを実行
    if excelWrite_overwrite(writeCell):
        excelBook[writeCell] = writeValue
        excelData.save(os.path.join(filePath, fileName))
        print("就活時間 " +str(writeValue) + "[h] で書き込み完了\n")
        logging.info(writeCell + "のセルに就活時間 " + str(writeValue) + " を書き込み成功")
    #書き込みを中断
    else:
        print("処理を中断しました\n")
        logging.info(writeCell + "のセルへの就活時間 " + str(writeValue) + " への書き込みを中断")

#SAと授業に参加した時間の書き込み
def excelWrite_Teach(day):
    writeCell = "CQ" + str(day + 5)
    while True:
        timeTeach = input("SAと授業に参加した時間を入力してください : ")
        try:
            writeValue = float(timeTeach)
            break
        except:
            print("WARNING : 数値で入力してください")
    #書き込みを実行
    if excelWrite_overwrite(writeCell):
        excelBook[writeCell] = writeValue
        excelData.save(os.path.join(filePath, fileName))
        print("SAと授業参加時間 " + str(writeValue) + "[h] で書き込み完了\n")
        logging.info(writeCell + "のセルにSA,授業時間 " + str(writeValue) + " を書き込み成功")
    #書き込みを中断
    else:
        print("処理を中断しました\n")
        logging.info(writeCell + "のセルへのSAと授業に参加した時間 " + str(writeValue) + " への書き込みを中断")


#セルに上書きするかの確認
def excelWrite_overwrite(cell):
    #書き込むセルに値が存在するとき
    if excelBook[cell].value:
        while True:
            print(f"セルには {excelBook[cell].value} が入っています ", end='')
            overwriteState = input("上書きしますか(y/n) : ")
            if (overwriteState == 'y'):
                logging.info(cell + "のセルの値 " + str(excelBook[cell].value) + " を削除")
                return True
            elif (overwriteState == 'n'):
                return False
            else:
                print("WARNING : y or nで入力してください")
    #書き込むセルに値が存在しないときはそのまま書き込み
    else:
        return True


if __name__=="__main__":
    main()
