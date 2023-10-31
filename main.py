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
        date = today.strftime("%Y%m%d%H%M%S")  # 오늘 연월일 날짜
        new_statement = statement.Statement(date=date, date_str=date, path='')
        new_statement.make_new_workbook()

        new_layout = new_statement.make_new_layout()

        new_statement.enter_item(
            account=account,
            item=item,
            intercell=new_layout
        )

        q = queue.Queue()

        file_name = "거래 명세서 " + date + ".xlsx"
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
            img_path = f"./이미지/{i}.png"
            img.save(img_path)
        '''
        return send_file(result + '.pdf', mimetype='pdf')

    return app


def makeJpg(file_name, store_name, q):
    # 엑셀을 실행할 객체 생성
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    # pdf로 변환할 파일명 선택
    wb = excel.Workbooks.Open(os.getcwd() + "/거래명세서/" + file_name)

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
    savepath = os.getcwd() + "/이미지/" + file_name

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
    host = input('호스트 IP를 입력하세요: ')
    port = int(input('포트번호를 입력하세요: '))

    create_app().run(host=host, port=port)
