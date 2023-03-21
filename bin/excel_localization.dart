import 'dart:io';
import 'package:excel/excel.dart';
import 'package:quartet/quartet.dart';

void main(List<String> arguments) async {
  const filePathInput = 'excel/input.xlsx';
  const filePathOutput = 'excel/output.xlsx';
  const fileKeysPath = 'excel/keys.text';
  final bytes = File(filePathInput).readAsBytesSync();
  final excel = Excel.decodeBytes(bytes);

  for (var table in excel.tables.entries) {
    final sheet = table.value;
    final rows = sheet.rows;
    final newValues = <String>[];
    var firstRow = rows.first;
    for (var cell in firstRow) {
      var key = cell.value.toString().trim();
      key = camelCase(key);
      final regex =
          r'[^\p{Alphabetic}\p{Mark}\p{Decimal_Number}\p{Connector_Punctuation}\p{Join_Control}\s]+';
      key = key.replaceAll(RegExp(regex, unicode: true), '');
      newValues.add(key);
    }
    print(newValues);
    excel.sheets.entries.first.value.insertRow(0);
    excel.sheets.entries.first.value.insertRowIterables(newValues, 0);
  }
  var fileBytes = excel.save();
  File(filePathOutput)
    ..createSync(recursive: true)
    ..writeAsBytesSync(fileBytes);
  var keys = '';
  for (final e in excel.sheets.entries.first.value.rows.first
      .map((e) => e.value)
      .toSet()) {
    keys += 'const $e = "$e";' '\n';
  }
  File(fileKeysPath)
    ..createSync(recursive: true)
    ..writeAsString(keys);
}
