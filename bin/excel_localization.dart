import 'dart:io';
import 'package:excel/excel.dart';
import 'package:quartet/quartet.dart';

void main(List<String> arguments) async {
  const filePathInput = 'excel/input.xlsx';
  final excel = getExcel(filePathInput);
  final sheet = getFirstSheet(excel);
  final firstRow = getFirstRow(sheet);
  final fileBytes = createKeysRow(firstRow, excel);
  const filePathOutput = 'excel/output.xlsx';
  saveExcelfile(filePathOutput, fileBytes);
  final keys = createKeysText(excel);
  saveKeysFile(keys);
}

void saveKeysFile(String keys) {
  const fileKeysPath = 'excel/keys.text';
  File(fileKeysPath)
    ..createSync(recursive: true)
    ..writeAsString(keys);
}

String createKeysText(Excel excel) {
  var keys = '';
  for (final e in excel.sheets.entries.first.value.rows.first
      .map((e) => e.value)
      .toSet()) {
    keys += 'const $e = "$e";' '\n';
  }
  return keys;
}

void saveExcelfile(String filePathOutput, List<int> fileBytes) {
  File(filePathOutput)
    ..createSync(recursive: true)
    ..writeAsBytesSync(fileBytes);
}

List<int> createKeysRow(List<Data> firstRow, Excel excel) {
  final newValues = <String>[];
  for (var cell in firstRow) {
    var key = cell.value.toString().trim();
    key = formatKeyText(key);
    newValues.add(key);
  }
  appendKeys(excel, newValues);
  final fileBytes = excel.save();
  return fileBytes;
}

void appendKeys(Excel excel, List<String> newValues) {
  excel.sheets.entries.first.value.insertRow(0);
  excel.sheets.entries.first.value.insertRowIterables(newValues, 0);
}

String formatKeyText(String key) {
  key = camelCase(key);
  final regex =
      r'[^\p{Alphabetic}\p{Mark}\p{Decimal_Number}\p{Connector_Punctuation}\p{Join_Control}\s]+';
  key = key.replaceAll(RegExp(regex, unicode: true), '');
  return key;
}

List<Data> getFirstRow(Sheet sheet) {
  final rows = sheet.rows;
  final firstRow = rows.first;
  return firstRow;
}

Sheet getFirstSheet(Excel excel) {
  final table = excel.tables.entries.first;
  final sheet = table.value;
  return sheet;
}

Excel getExcel(String filePathInput) {
  final bytes = File(filePathInput).readAsBytesSync();
  final excel = Excel.decodeBytes(bytes);
  return excel;
}
