import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:flutter/material.dart';
import 'package:dio/dio.dart';
import 'package:test_flutter/screen/test_screen.dart';

class LoadStatementScreen extends StatefulWidget {
  const LoadStatementScreen({Key? key}) : super(key: key);

  @override
  State<LoadStatementScreen> createState() => _LoadStatementScreenState();
}

class _LoadStatementScreenState extends State<LoadStatementScreen> {

  late Map<String, Map<String, dynamic>> items;

  @override
  void initState(){
    items = {
      "item" : {
        "name" : ["빵", "초코", "초코", "초코"],
        "cnt" : [1,2,2,2]
      },
      "account" : {
        "name" : "거래처 이름",
        "ph_n" : "01073394768",
        "loc" : "경기도 김포"
      }
    };
  }
  bool downloading = false;
  String progressString = '';

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      floatingActionButton: FloatingActionButton(
        onPressed: (){
          Navigator.push(context, MaterialPageRoute(builder: (_)=>TestScreen()));
        },
        child: Text('다운받은 목록'),
      ),
      body: Container(
        width: MediaQuery.of(context).size.width,
        height: MediaQuery.of(context).size.height,
        child: Center(
          child: ElevatedButton(
            onPressed: (){
              showDownloadStatement();
            },
            child: Text('거래 명세서 출력'),
          ),
        ),
      ),
    );
  }
  
  void showDownloadStatement(){
    progressString = '';
    showDialog(context: context, builder: (context){
      return StatefulBuilder(
        builder: (context, setState){
          return AlertDialog(
            title: ElevatedButton(
              onPressed: (){
                load_statement(items, setState);
              },
              child: Text('거래 명세서 다운'),
            ),
            content: Row(
              children: [
                Text("다운로드: "+progressString),
              ],
            ),
          );
        },
      );
    });
  }
  Future<void> load_statement(Map<String, Map<String, dynamic>> items, StateSetter setState) async{

    Dio dio = Dio();
    String url = "http://10.0.2.2:50000/load_statement";

    DateTime now = DateTime.now();
    String formattedDate = "${now.year}년${now.month}월${now.day}일${now.hour}시${now.minute}분${now.second}초";
    progressString = '';
    var tempDir = await getTemporaryDirectory();
    String fullPath = tempDir.path + "/"+formattedDate +".pdf";
    print('full path ${fullPath}');

    download2(dio, url, fullPath, items, setState);
  }
  Future download2(Dio dio, String url, String savePath, Map<String, Map<String, dynamic>> items, StateSetter setState) async {
    try {
      Response response = await dio.post(
          url,
          onReceiveProgress: (received, total){
            setState(() {
              progressString = (received / total * 100).toStringAsFixed(0) + "%";
            });
          },
          //Received data with List<int>
          options: Options(
              responseType: ResponseType.bytes,
              followRedirects: false,
              validateStatus: (status) {
                return status! < 500;
              }),
          data: items
      );
      print(response.headers);
      File file = File(savePath);
      var raf = file.openSync(mode: FileMode.write);
      // response.data is List<int> type
      raf.writeFromSync(response.data);
      await raf.close();
    } catch (e) {
      print(e);
    }

    downloading = false;
  }
}

