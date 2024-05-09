import 'package:flutter/material.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'package:excel/excel.dart';
import 'dart:convert';
import 'package:fl_chart/fl_chart.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel to JSON',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  @override
  _MyHomePageState createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  List<Map<String, dynamic>> jsonData = [];

  @override
  void initState() {
    super.initState();
    loadDataFromExcel();
  }

  Future<void> loadDataFromExcel() async {
    try {
      var data = await rootBundle.load('assets/data.xlsx');
      var bytes = data.buffer.asUint8List();
      var excel = Excel.decodeBytes(bytes);
      var sheet = excel.tables['Sheet1'];

      // Clear previous data
      setState(() {
        jsonData.clear();
      });

      // Extract data and convert to JSON
      for (var row in sheet!.rows) {
        var rowData = {
          "Country": row[0]?.value.toString(),
          "Supplier": row[1]?.value.toString(),
        };
        setState(() {
          jsonData.add(rowData);
        });
      }
    } catch (e) {
      print("Error loading Excel sheet: $e");
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel to JSON'),
      ), //Highlight
      body: jsonData.isEmpty
          ? Center(child: CircularProgressIndicator())
          : FutureBuilder<Map<String, int>>(
              future: getCountByCountry(),
              builder: (context, snapshot) {
                if (snapshot.connectionState == ConnectionState.waiting) {
                  return Center(child: CircularProgressIndicator());
                } else if (snapshot.hasError) {
                  return Center(child: Text('Error: ${snapshot.error}'));
                } else if (!snapshot.hasData || snapshot.data!.isEmpty) {
                  return Center(child: Text('No data available'));
                } else {
                  Map<String, int> countByCountry = snapshot.data!;
                  List<PieChartSectionData> pieChartSections =
                      countByCountry.entries.map((entry) {
                    return PieChartSectionData(
                      color: Color(
                          (entry.key.hashCode & 0xFFFFFF).toUnsigned(8) |
                              0xFF000000), // Random color
                      value: entry.value.toDouble(),
                      title: entry.key,
                      radius: 100,
                    );
                  }).toList();
                  return Center(
                    child: PieChart(
                      PieChartData(
                        sections: pieChartSections,
                        borderData: FlBorderData(show: false),
                        centerSpaceRadius: 40,
                        sectionsSpace: 0,
                      ),
                    ),
                  );
                }
              },
            ),
    );
  }

  Future<Map<String, int>> getCountByCountry() async {
    Map<String, int> countByCountry = {};
    for (var item in jsonData.skip(1)) {
      String country = item['Country'];
      countByCountry[country] = (countByCountry[country] ?? 0) + 1;
    }
    return countByCountry;
  }
}
