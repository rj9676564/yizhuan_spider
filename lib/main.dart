import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:html/dom.dart' as htmlDom;
import 'package:html/parser.dart' as htmlParser;
import 'package:http/http.dart' as http;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: '易撰文章获取',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: '易撰文章获取'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  int _counter = 0;

  void _incrementCounter() {
    setState(() {
      _counter++;
    });
  }

  selectFile() async {
    // Navigate to the settings page
    print("file selector");
    const XTypeGroup typeGroup = XTypeGroup(
      label: '易撰文档',
      extensions: <String>['xls', 'xlsx'],
    );
    final XFile? file =
        await openFile(acceptedTypeGroups: <XTypeGroup>[typeGroup]);

    if (file != null) {
      readXls(file!.path);
      print('file: ' + file!.path);
    } else {
      print('file: null');
    }
  }

  String sanitizeTitle(String title) {
    // 保留斜杠
    var sanitizedTitle = title.replaceAll('/', '-');
    return sanitizedTitle;
  }

  readXls(filePath) async {
    Uint8List bytes = File(filePath).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    // 打印工作表名称
    print(excel.tables.keys);

    // 获取第一个工作表
    var table = excel.tables.keys.first;
    var sheet = excel.tables[table];

    // 创建一个列表来保存解析后的数据
    List<Map<String, dynamic>> dataList = [];

    // 遍历每一行，跳过标题行
    for (var row in sheet!.rows.skip(1)) {
      // 创建一个Map来保存当前行的数据
      var data = <String, dynamic>{};

      // 将每列的数据添加到Map中
      data['title'] = row[0]?.value.toString();
      data['阅读量'] = row[1]?.value.toString();
      data['评论数'] = row[2]?.value.toString();
      data['领域'] = row[3]?.value.toString();
      data['作者'] = row[4]?.value.toString();
      data['发布时间'] = row[5]?.value.toString();
      data['url'] = row[6]?.value.toString();
      data['imageUrl'] = row[7]?.value.toString();

      // 将当前行的数据Map添加到列表中
      dataList.add(data);
    }
    print(dataList);
    // 根据标题创建文件夹
    for (var data in dataList) {
      var title = data['title'];
      title = sanitizeTitle(title);
      var folder = Directory("${Directory.systemTemp.path}/tmp/$title");
      folder.createSync(recursive: true);
      File file = File("${Directory.systemTemp.path}/tmp/$title/$title.txt");
      print('data: ${file.absolute}');
      var url = data['url'];
      var htmlContent = await fetchHtmlContent(url);
      // File("${Directory.systemTemp.path}/tmp/$title/news.txt")
      //     .writeAsStringSync(htmlContent);

      htmlContent = extractContent(htmlContent, folder);
      print('htmlContent: $htmlContent');
      file.writeAsStringSync(htmlContent);
      print(
          '========================================================================');
    }

    // var excel = Excel.decodeBytes(bytes);
    // for (var table in excel.tables.keys) {
    //   print(table);
    //   print(excel.tables[table]!.maxColumns);
    //   print(excel.tables[table]!.maxRows);
    //   for (var row in excel.tables[table]!.rows) {
    //     print("${row.map((e) => e?.value)}");
    //   }
    // }
  }

  String extractContent(String htmlContent, Directory directory) {
    var document = htmlParser.parse(htmlContent);
    // 查找所有符合条件的 <div> 元素
    var divElements =
        document.querySelectorAll('div[data-role="original-title"]');
    // 遍历所有符合条件的 <div> 元素
    String txt = "";
    int index = 0;
    for (var divElement in divElements) {
      // 遍历 <div> 元素的子节点
      divElement.nodes.forEach((node) async {
        // 如果是 <p> 元素且包含 <img> 元素，则打印图像链接

        if (node is htmlDom.Element &&
            node.localName == 'p' &&
            node.querySelector('img') != null) {
          var imgUrl = node.querySelector('img')?.attributes['src'];
          print('$imgUrl');
          if (imgUrl != null) {
            index++;
            downloadImage(imgUrl, '${directory.path}/$index.jpg');
            txt = "$txt${'这里是文章图片\\$index.jpg'}\n";
          }
        } else if (node is htmlDom.Element && node.localName == 'p') {
          // 如果是 <p> 元素且包含 <span> 元素，则打印文本内容

          // var textContent = node.querySelector('span')?.text.trim();
          var textContent = node.text.trim();
          print('文本： $textContent');
          if (textContent != null) {
            if (textContent.startsWith("出品") ||
                textContent.startsWith("来源") ||
                textContent.startsWith("Author") ||
                textContent.startsWith("#") ||
                textContent.startsWith("作者") ||
                textContent.startsWith("摄像") ||
                textContent.startsWith("编审") ||
                textContent.startsWith("特约作者") ||
                textContent.startsWith("监制") ||
                textContent.startsWith("总编") ||
                textContent.startsWith("主持人") ||
                textContent.startsWith("More") ||
                textContent.startsWith("more") ||
                textContent.startsWith("www") ||
                textContent.startsWith("编辑")) {
              return;
            }
            textContent = textContent.replaceAll('文章来源', '');
            txt = "$txt$textContent\n";
          }
        }
      });
    }

    print(
        '--------------------------------------------------------------------------------- ' +
            txt);
    return txt;
  }

  Future<void> downloadImage(String url, String savePath) async {
    if (url.startsWith('//')) {
      url = 'https:$url';
    }
    var client = HttpClient();
    var request = await client.getUrl(Uri.parse(url));
    request.headers.set(HttpHeaders.userAgentHeader,
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36');
    request.headers
        .set(HttpHeaders.contentTypeHeader, 'image/jpeg'); // 指定 Content-Type

    var response = await request.close();
    if (response.statusCode == HttpStatus.ok) {
      var file = File(savePath);
      await response.pipe(file.openWrite());
      print('图片已保存到：$savePath');
    } else {
      print('下载失败：HTTP ${response.statusCode}');
    }
  }

  Future<String> fetchHtmlContent(String url) async {
    var response = await http.get(Uri.parse(url));

    if (response.statusCode == 200) {
      return utf8.decode(response.bodyBytes);
    } else {
      throw Exception('Failed to load HTML content: ${response.statusCode}');
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
        actions: [
          IconButton(icon: const Icon(Icons.settings), onPressed: selectFile),
        ],
      ),
      body: Center(
        child: Column(
          children: <Widget>[
            const Text(
              'You have pushed the button this many times:',
            ),
            Text(
              '$_counter',
              style: Theme.of(context).textTheme.headlineMedium,
            ),
          ],
        ),
      ),
      floatingActionButton: FloatingActionButton(
        onPressed: _incrementCounter,
        tooltip: 'Increment',
        child: const Icon(Icons.add),
      ), // This trailing comma makes auto-formatting nicer for build methods.
    );
  }
}
