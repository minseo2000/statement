# -*- coding: utf-8 -*-
import datetime
import os
import statement
from flask import Flask, request, json, jsonify, send_file
import win32com.client
import threading
import pythoncom
import queue

def create_app():
    app = Flask(__name__)

    @app.route('/load_statement', methods=['POST'])
    def load_statement():

        statement_info = request.json

        print(statement_info)

        account = statement_info['account']
        item = statement_info['item']
        today = datetime.datetime.now()
        date = account['name']+" "+today.strftime("%H시 %M분 %S초".encode('unicode-escape').decode()).encode().decode('unicode-escape')  # 오늘 연월일 날짜
        new_statement = statement.Statement(
            date=date,
            date_str=date,
            path='',
            NAME= "타르트에오",
            CHAIRMAN="정다정",
            LOCATION="경주시 원효로87 타르트에오",
            CATEGORY="일반음식점",
            CATE="카페",
            NUM="275 09 01625"
        )
        new_statement.make_new_workbook()

        new_layout = new_statement.make_new_layout()

        new_statement.enter_item(
            account=account,
            item=item,
            intercell=new_layout
        )

        q = queue.Queue()

        file_name = date + ".xlsx"
        store_name = account['name']

        t1 = threading.Thread(target=makeJpg(file_name=file_name, store_name=store_name, q=q), args=(q,))
        t1.start()

        t1.join()

        result = q.get()
        print(result)

        # pdf_path = makeJpg(file_name=file_name, store_name=store_name)
        '''
        import fitz
        doc = fitz.open(pdf_path+'.pdf')
        for i, page in enumerate(doc):
            img = page.get_pixmap()
            img_path = f"./pdf/{i}.png"
            img.save(img_path)
        '''

        return send_file(result + '.pdf', mimetype='pdf')

    return app

def makeJpg(file_name, store_name, q):




    today = datetime.datetime.now()

    save_path = today.strftime('%Y년 %m월 %d일'.encode('unicode-escape').decode()).encode().decode('unicode-escape')


    print(save_path)
    if not os.path.exists(os.getcwd()+"/pdf"):
        # 폴더가 존재하지 않으면 새로운 폴더를 만듭니다.
        os.mkdir(os.getcwd()+"/pdf")

    else:
        # 폴더가 이미 존재하면 사용자에게 알립니다.
        print(f" 폴더가 이미 존재합니다.")

    if not os.path.exists(os.getcwd()+"/pdf/"+save_path):
        # 폴더가 존재하지 않으면 새로운 폴더를 만듭니다.
        os.mkdir(os.getcwd()+"/pdf/"+save_path)
    else:
        # 폴더가 이미 존재하면 사용자에게 알립니다.
        print(f" 폴더가 이미 존재합니다.")







    # 엑셀을 실행할 객체 생성
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    # pdf로 변환할 파일명 선택
    wb = excel.Workbooks.Open(os.getcwd() + "/"+save_path+"/" + file_name)

    # 워크북의 시트명 설정
    ws_sht = wb.Worksheets(store_name)
    # 설정한 시트 선택
    ws_sht.Select()
    ws_sht.PageSetup.Orientation = 1
    ws_sht.PageSetup.LeftMargin = 160
    ws_sht.PageSetup.TopMargin = 20
    ws_sht.PageSetup.BottomMargin = 20
    ws_sht.PageSetup.RightMargin = 20
    # PDF파일을 저장할 경로 및 파일명 지정
    savepath = os.getcwd() + "/pdf/" + save_path+"/"+ file_name

    # 활성화된 시트를 pdf 저장
    wb.ActiveSheet.ExportAsFixedFormat(0, savepath, IgnorePrintAreas=False)

    # 엑셀 워크북 및 프로그램 종료
    # 종료를 제대로 해주어야 다음 실행시 에러 안생김
    wb.Close(False)
    excel.Quit()
    pythoncom.CoUninitialize()
    q.put(savepath)
    return savepath


if __name__ == '__main__':
    host = "0.0.0.0"
    port = 50000


    create_app().run(host=host, port=port)
