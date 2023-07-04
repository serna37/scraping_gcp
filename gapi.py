# ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import os.path

class GApi:

    def getSpreadSheet(self, sheet_key):

        # 実行中のファイルのとこに移動, トークンを参照
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

        # 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]

        # 秘密鍵ファイル名
        secret_key = 'flowing-density.json'
        # 秘密鍵から認証情報設定
        credentials = ServiceAccountCredentials.from_json_keyfile_name(secret_key, scope)
        # OAuth2の資格情報を使用してGoogle APIにログイン
        gc = gspread.authorize(credentials)

        # スプレッドシートを開く
        try:
            workbook = gc.open_by_key(sheet_key)
            return workbook
        except:
            return None

    def getLastRow(self, worksheet):
        # A1:A100のうち, 最も上の空白の行数を取得
        chkRange = worksheet.range('A1:A100')
        target_row = 0
        for idx, val in enumerate(chkRange):
            if val.value == '':
                target_row = idx + 1
                break
        return target_row if target_row != 0 else None

    def sendMail(self, to, cc, subject, body, att):
        # common mail GAS sync
        key = "XXXXXXX"
        book = GApi().getSpreadSheet(key)
        sheet = book.worksheet("xxxx")
        last_row = GApi().getLastRow(sheet)
        # to, cc
        sheet.update_cell(last_row, 1, to)
        sheet.update_cell(last_row, 2, cc)
        # subject
        sheet.update_cell(last_row, 3, subject)
        # body
        sheet.update_cell(last_row, 4, body)
        # att
        sheet.update_cell(last_row, 5, att)
        # commit
        sheet.update_acell('F1', 'ok')

    def upMailAttFile(self, name, mime):
        # トークン参照のため、起動中フォルダに移動. 直下のsecretsとcredentialsを参照
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        gauth = GoogleAuth()
        gauth.LocalWebserverAuth()
        #gauth.CommandLineAuth()
        drive = GoogleDrive(gauth)

        folder_id = 'XXXXXXX'
        file = drive.CreateFile({'title': name, 'mimeType': mime, 'parents': [{'id': folder_id}]})
        file.SetContentFile(name)
        file.Upload()
