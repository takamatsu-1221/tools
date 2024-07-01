# tools
ちょっとした作業効率化するツールの制作を行いました

## dile_rename.py
### コンセプト
フォルダにあるファイル名を全て変更したいとき，１つづつ選択して  
名前を変更するのは面倒であるため自動化を行った．  
今回は授業での課題を想定する．プログラムを実行し，フォルダのパスと  
ファイルの形式を指定することで学生番号_科目名_第n回と変更できる．
### システム設計
当該フォルダに指定した形式のファイルがあるか確認して  
名前を変更し，第n回の部分を順にカウントアップしていくだけの  
シンプルなコードである．


## schedule_make.js
### コンセプト
クレジットカードの引き落としや給与の振込は毎月同じ日に行われるが，  
Googleカレンダー上に1日1日手動で登録するのは面倒であるため自動化を行う．   
そこで，GoogleAppsScriptを用いて，プログラムを1度実行するだけで  
N年間のクレジットカードの引き落とし日をGoogleカレンダー上に登録するツールを作成した．  

### システム設計  
Y年M月の25日が平日か確かめる．土日祝の場合は日を遡り直前の平日を探す．  
予定を登録し，リマインダが届くようにGoogleカレンダーに設定を行う．  
これらの処理をY年×M=12ヶ月分行う．


## write_automation.py
### コンセプト
研究室ではExcelに研究時間を記録している．従来はファイルを開き時刻を手入力していたが，  
毎回ファイルを開くのは面倒と感じたためファイルへの記入の自動化を行った．
### システム設計
プログラムを実行すると，メニューが表示される．  
例えば，1を選択すると登校時刻の欄に，2を選択すると帰宅時刻の欄に，  
現在の時間が自動的に記入される． また，3を選択すると休憩時刻の記入ができる．  
休憩時刻は日によって異なるため，"1.5"(時間)のように手動で入力する．  
誤った操作を行い，入力しようとしているセルに既にデータが入っている場合は，  
上書きを行うか確認する画面が表示されるため，誤操作の心配は無用である．  
さらに，意図しない挙動によりデータが削除された時のバックアップとして  
ログを出力している．このログには何時にプログラムを実行し，どこのセルにどのような時刻を  
記入したかを記録している．
