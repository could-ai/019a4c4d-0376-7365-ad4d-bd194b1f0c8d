import 'dart:io';
import 'package:flutter/material.dart';
import 'package:image_picker/image_picker.dart';
import 'package:google_ml_kit/google_ml_kit.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:share_plus/share_plus.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'تحويل صورة الجدول إلى Excel',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  File? _image;
  String _extractedText = '';
  bool _isProcessing = false;
  String _statusMessage = 'اختر صورة للبدء';

  final ImagePicker _picker = ImagePicker();
  final TextRecognizer _textRecognizer = GoogleMlKit.vision.textRecognizer();

  Future<void> _pickImage(ImageSource source) async {
    if (source == ImageSource.camera) {
      var cameraStatus = await Permission.camera.request();
      if (!cameraStatus.isGranted) {
        setState(() {
          _statusMessage = 'يجب منح إذن الكاميرا';
        });
        return;
      }
    }

    var storageStatus = await Permission.storage.request();
    if (!storageStatus.isGranted) {
      setState(() {
        _statusMessage = 'يجب منح إذن التخزين';
        _statusMessage = 'يجب منح إذن التخزين';
      });
      return;
    }

    final XFile? image = await _picker.pickImage(source: source);
    if (image != null) {
      setState(() {
        _image = File(image.path);
        _extractedText = '';
        _statusMessage = 'تم اختيار الصورة. اضغط على "استخراج النص"';
      });
    }
  }

  Future<void> _extractText() async {
    if (_image == null) return;

    setState(() {
      _isProcessing = true;
      _statusMessage = 'جاري استخراج النص...';
    });

    final inputImage = InputImage.fromFile(_image!);
    final RecognizedText recognizedText = await _textRecognizer.processImage(inputImage);

    setState(() {
      _extractedText = recognizedText.text;
      _isProcessing = false;
      _statusMessage = 'تم استخراج النص. اضغط على "تحويل إلى Excel"';
    });
  }

  Future<void> _convertToExcel() async {
    if (_extractedText.isEmpty) return;

    setState(() {
      _isProcessing = true;
      _statusMessage = 'جاري إنشاء ملف Excel...';
    });

    try {
      var excel = Excel.createExcel();
      Sheet sheetObject = excel['Sheet1'];

      // Simple parsing: split by lines and tabs/spaces
      List<String> lines = _extractedText.split('\n');
      for (int i = 0; i < lines.length; i++) {
        List<String> cells = lines[i].split(RegExp(r'\s+'));
        for (int j = 0; j < cells.length; j++) {
          sheetObject.cell(CellIndex.indexByColumnRow(columnIndex: j, rowIndex: i)).value = cells[j];
        }
      }

      Directory? directory;
      if (Platform.isAndroid) {
        directory = await getExternalStorageDirectory();
      } else {
        directory = await getApplicationDocumentsDirectory();
      }

      String filePath = '${directory!.path}/table.xlsx';
      File file = File(filePath);
      await file.writeAsBytes(excel.encode()!);

      setState(() {
        _isProcessing = false;
        _statusMessage = 'تم إنشاء ملف Excel: $filePath';
      });

      // Share the file
      await Share.shareXFiles([XFile(filePath)], text: 'ملف Excel من الصورة');
    } catch (e) {
      setState(() {
        _isProcessing = false;
        _statusMessage = 'خطأ في إنشاء Excel: $e';
      });
    }
  }

  @override
  void dispose() {
    _textRecognizer.close();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('تحويل صورة الجدول إلى Excel'),
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceEvenly,
              children: [
                ElevatedButton.icon(
                  onPressed: () => _pickImage(ImageSource.gallery),
                  icon: const Icon(Icons.photo),
                  label: const Text('من المعرض'),
                ),
                ElevatedButton.icon(
                  onPressed: () => _pickImage(ImageSource.camera),
                  icon: const Icon(Icons.camera),
                  label: const Text('الكاميرا'),
                ),
              ],
            ),n            if (_image != null) ...[
              const SizedBox(height: 16),
              Image.file(_image!, height: 200),
              const SizedBox(height: 16),
              ElevatedButton(
                onPressed: _isProcessing ? null : _extractText,
                child: const Text('استخراج النص'),
              ),
            ],
            const SizedBox(height: 16),
            if (_extractedText.isNotEmpty) ...[
              Expanded(
                child: SingleChildScrollView(
                  child: Text('النص المستخرج:\n$_extractedText'),
                ),
              ),
              const SizedBox(height: 16),
              ElevatedButton(
                onPressed: _isProcessing ? null : _convertToExcel,
                child: const Text('تحويل إلى Excel'),
              ),
            ],
            const SizedBox(height: 16),
            Text(_statusMessage, textAlign: TextAlign.center),
            if (_isProcessing) const CircularProgressIndicator(),
          ],
        ),
      ),
    );
  }
}