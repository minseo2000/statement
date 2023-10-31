import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';

class TestScreen extends StatefulWidget {
  const TestScreen({Key? key}) : super(key: key);

  @override
  State<TestScreen> createState() => _TestScreenState();
}

class _TestScreenState extends State<TestScreen> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        width: MediaQuery.of(context).size.width,
        height: MediaQuery.of(context).size.height,
        child: DownloadedFilesList()
      ),
    );
  }
}


class DownloadedFilesList extends StatefulWidget {
  @override
  _DownloadedFilesListState createState() => _DownloadedFilesListState();
}

class _DownloadedFilesListState extends State<DownloadedFilesList> {
  late List<FileSystemEntity> _files;

  @override
  void initState() {
    super.initState();
    _files = [];
    _listDownloadedFiles();
  }

  _listDownloadedFiles() async {
    Directory? downloadsDirectory;

    // 앱의 다운로드 폴더를 가져옵니다.
    try {
      downloadsDirectory = await getTemporaryDirectory();
    } catch (e) {
      print(e);
    }

    if (downloadsDirectory != null) {
      setState(() {
        _files = downloadsDirectory!.listSync();
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('다운로드한 파일 목록'),
      ),
      body: ListView.builder(
        itemCount: _files.length,
        itemBuilder: (context, index) {
          return ListTile(
            title: Text(_files[index].path.split('/').last),
            onTap: () async {
              // 파일을 열기 위한 로직
              await OpenFile.open(_files[index].path);
            },
          );
        },
      ),
    );
  }
}
