import 'dart:io';

import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:fluttertoast/fluttertoast.dart';
import 'package:path_provider/path_provider.dart';
import 'package:path/path.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Demo',
      home: ExcelDemo(),
    );
  }
}

class ExcelDemo extends StatefulWidget {
  @override
  _ExcelDemoState createState() => _ExcelDemoState();
}

class _ExcelDemoState extends State<ExcelDemo> {
  final TextEditingController _value1Controller = TextEditingController();
  final TextEditingController _value2Controller = TextEditingController();

  Future<double?> _createExcelFile(int value1, int value2) async {
    var excel = Excel.createExcel();
    var sheet = excel['Sheet1'];

    // Set headers
    sheet.cell(CellIndex.indexByString("A1")).value = TextCellValue('Value1');
    sheet.cell(CellIndex.indexByString("B1")).value = TextCellValue('Value2');
    sheet.cell(CellIndex.indexByString("C1")).value = TextCellValue('Total');



    // Set values
    sheet.cell(CellIndex.indexByString("A2")).value = IntCellValue(value1);
    sheet.cell(CellIndex.indexByString("B2")).value = IntCellValue(value2);

    // Use formula for total
    sheet.cell(CellIndex.indexByString("C2")).value = FormulaCellValue('SUM(A2:B2)');

    // Save the file
    Directory directory = await getApplicationDocumentsDirectory();
    String outputPath = join(directory.path, 'values.xlsx');
    List<int>? fileBytes = excel.save();
    if (fileBytes != null) {
      File(outputPath)
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
      print('Excel file saved to $outputPath');
    } else {
      print('Failed to save Excel file');
    }

    // Return the total value from the formula
    // Note: Excel formulas are not evaluated by the library in Dart, so this will return null
    return null;
  }

  void _onProceed() async {
    int value1 = int.tryParse(_value1Controller.text) ?? 0;
    int value2 = int.tryParse(_value2Controller.text) ?? 0;
     _createExcelFile(value1, value2);
     var total = value1 + value2;

    Fluttertoast.showToast(
      msg: 'Your Total: ${total ?? 'Calculation not available'}',
      backgroundColor: Colors.black,
      textColor: Colors.white,
      gravity: ToastGravity.BOTTOM,
      toastLength: Toast.LENGTH_SHORT,
    );

    _value1Controller.clear();
    _value2Controller.clear();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel Demo'),
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: [
            TextField(
              controller: _value1Controller,
              decoration: InputDecoration(
                labelText: 'Enter Value 1',
                border: OutlineInputBorder(),
              ),
              keyboardType: TextInputType.number,
            ),
            SizedBox(height: 16.0),
            TextField(
              controller: _value2Controller,
              decoration: InputDecoration(
                labelText: 'Enter Value 2',
                border: OutlineInputBorder(),
              ),
              keyboardType: TextInputType.number,
            ),
            SizedBox(height: 16.0),
            ElevatedButton(
              onPressed: _onProceed,
              child: Text('Proceed'),
            ),
          ],
        ),
      ),
    );
  }
}
