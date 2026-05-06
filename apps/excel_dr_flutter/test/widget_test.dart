import 'package:excel_dr_flutter/main.dart';
import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  testWidgets('shows the Excel-Dr workflow', (tester) async {
    tester.view.physicalSize = const Size(1280, 900);
    tester.view.devicePixelRatio = 1;
    addTearDown(tester.view.resetPhysicalSize);
    addTearDown(tester.view.resetDevicePixelRatio);

    await tester.pumpWidget(const ExcelDrApp());

    expect(find.text('Excel-Dr'), findsOneWidget);
    expect(find.text('让卡顿的 Excel 变顺畅'), findsOneWidget);
    expect(find.text('仅检测'), findsOneWidget);
    expect(find.text('清理并新建'), findsOneWidget);
  });
}
