import 'dart:async';
import 'dart:io';

import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:rinf/rinf.dart';
import 'package:window_manager/window_manager.dart';

import 'src/bindings/bindings.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await windowManager.ensureInitialized();
  await windowManager.waitUntilReadyToShow(
    const WindowOptions(
      size: Size(1280, 900),
      minimumSize: Size(1280, 900),
      center: true,
      title: 'Excel-Dr',
    ),
    () async {
      await windowManager.show();
      await windowManager.focus();
    },
  );
  await initializeRust(assignRustSignal);
  runApp(const ExcelDrApp());
}

class ExcelDrApp extends StatelessWidget {
  const ExcelDrApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Excel-Dr',
      theme: ThemeData(
        fontFamily: 'Microsoft YaHei UI',
        scaffoldBackgroundColor: const Color(0xfff4f8f8),
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xff08735f),
          brightness: Brightness.light,
        ),
        useMaterial3: true,
        textSelectionTheme: const TextSelectionThemeData(cursorColor: Color(0xff08735f)),
      ),
      home: const ExcelDrHome(),
    );
  }
}

class ExcelDrHome extends StatefulWidget {
  const ExcelDrHome({super.key});

  @override
  State<ExcelDrHome> createState() => _ExcelDrHomeState();
}

class _ExcelDrHomeState extends State<ExcelDrHome> {
  static const _xlsxType = XTypeGroup(label: 'Excel 工作簿', extensions: ['xlsx']);

  StreamSubscription<RustSignalPack<TaskResult>>? _resultSub;
  StreamSubscription<RustSignalPack<TaskProgress>>? _progressSub;
  String _targetKind = 'file';
  String _selectedPath = '';
  int _requestId = 0;
  bool _busy = false;
  int _progress = 0;
  String _progressTitle = '等待选择文件';
  String _progressDetail = '选择一个 .xlsx 文件或文件夹后开始检测';
  TaskResult? _result;

  @override
  void initState() {
    super.initState();
    _resultSub = TaskResult.rustSignalStream.listen((pack) {
      final message = pack.message;
      if (message.requestId != _requestId) return;
      setState(() {
        _busy = false;
        _progress = 100;
        _progressTitle = message.message;
        _progressDetail = message.outputPath.isEmpty ? '任务已完成' : message.outputPath;
        _result = message;
      });
    });
    _progressSub = TaskProgress.rustSignalStream.listen((pack) {
      final message = pack.message;
      if (message.requestId != _requestId) return;
      setState(() {
        _busy = true;
        _progress = message.percent.clamp(0, 99);
        _progressTitle = message.detail;
        _progressDetail = message.current;
      });
    });
  }

  @override
  void dispose() {
    _resultSub?.cancel();
    _progressSub?.cancel();
    super.dispose();
  }

  Future<void> _selectTarget() async {
    if (_busy) return;
    if (_targetKind == 'file') {
      final file = await openFile(acceptedTypeGroups: [_xlsxType]);
      if (file == null) return;
      await _setSelectedPath(file.path);
      return;
    }
    final folder = await getDirectoryPath(confirmButtonText: '选择文件夹');
    if (folder == null) return;
    await _setSelectedPath(folder);
  }

  Future<void> _setSelectedPath(String path) async {
    int? size;
    if (_targetKind == 'file') {
      try {
        size = await File(path).length();
      } catch (_) {
        size = null;
      }
    }
    setState(() {
      _selectedPath = path;
      _result = null;
      _progress = 0;
      _progressTitle = '已选择${_targetKind == 'file' ? '文件' : '文件夹'}';
      _progressDetail = _targetKind == 'file' && size != null
          ? '${_formatSize(size)} · 将在原目录生成清理版'
          : '检测和清理都不会修改原文件';
    });
  }

  void _sendTask(String operation) {
    if (_busy || _selectedPath.isEmpty) return;
    setState(() {
      _requestId += 1;
      _busy = true;
      _progress = 1;
      _progressTitle = operation == 'analyze' ? '正在检测' : '正在清理';
      _progressDetail = _fileName(_selectedPath);
      _result = null;
    });
    TaskRequest(
      requestId: _requestId,
      operation: operation,
      targetKind: _targetKind,
      path: _selectedPath,
    ).sendSignalToRust();
  }

  Future<void> _openOutput() async {
    final output = _result?.outputPath;
    final path = (output != null && output.isNotEmpty) ? output : _selectedPath;
    if (path.isEmpty) return;
    if (Platform.isWindows && output != null && output.isNotEmpty) {
      await Process.start('explorer.exe', ['/select,', path]);
      return;
    }
    if (Platform.isWindows) {
      await Process.start('explorer.exe', [path]);
    }
  }

  void _showHelp() {
    showDialog<void>(
      context: context,
      barrierColor: const Color(0x6b102027),
      builder: (context) => AlertDialog(
        backgroundColor: Colors.transparent,
        elevation: 0,
        contentPadding: EdgeInsets.zero,
        content: Container(
          width: 560,
          decoration: BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.circular(18),
            border: Border.all(color: const Color(0xffd9e7e8)),
            boxShadow: const [BoxShadow(color: Color(0x38102027), blurRadius: 70, offset: Offset(0, 24))],
          ),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Container(
                height: 60,
                padding: const EdgeInsets.symmetric(horizontal: 20),
                decoration: const BoxDecoration(border: Border(bottom: BorderSide(color: Color(0xffd9e7e8)))),
                child: Row(
                  children: [
                    const Text('使用说明', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w900)),
                    const Spacer(),
                    _AppButton(label: '关闭', onPressed: () => Navigator.of(context).pop()),
                  ],
                ),
              ),
              const Padding(
                padding: EdgeInsets.all(20),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    _HelpItem(title: '1. 先选文件', text: '选择一个 Excel 文件，或选择包含多个 Excel 文件的文件夹。'),
                    SizedBox(height: 14),
                    _HelpItem(title: '2. 先检测', text: '检测只查看问题，不会修改你的文件。'),
                    SizedBox(height: 14),
                    _HelpItem(title: '3. 再清理', text: '发现问题后，点击“清理并新建”，软件会生成一个新的清理版文件。'),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        width: double.infinity,
        height: double.infinity,
        decoration: const BoxDecoration(
          gradient: LinearGradient(
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
            colors: [Color(0xfff7fbfb), Color(0xffedf5f6)],
          ),
        ),
        child: Center(
          child: SizedBox(
            width: 1224,
            height: 844,
          child: Container(
            decoration: BoxDecoration(
              color: Colors.white.withValues(alpha: 0.94),
              borderRadius: BorderRadius.circular(22),
              border: Border.all(color: const Color(0xffd9e7e8)),
              boxShadow: const [
                BoxShadow(
                  color: Color(0x19102027),
                  blurRadius: 48,
                  offset: Offset(0, 18),
                ),
              ],
            ),
            child: Column(
              children: [
                _TopBar(onHelp: _showHelp),
                Expanded(
                  child: Padding(
                    padding: const EdgeInsets.fromLTRB(28, 16, 28, 22),
                    child: Row(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      children: [
                        SizedBox(
                          width: 370,
                          child: _LeftPane(
                            targetKind: _targetKind,
                            selectedPath: _selectedPath,
                            busy: _busy,
                            onKindChanged: (kind) {
                              if (_busy) return;
                              setState(() {
                                _targetKind = kind;
                                _selectedPath = '';
                                _result = null;
                                _progress = 0;
                                _progressTitle = '等待选择文件';
                                _progressDetail = '选择一个 .xlsx 文件或文件夹后开始检测';
                              });
                            },
                            onSelect: _selectTarget,
                            onAnalyze: () => _sendTask('analyze'),
                            onClean: () => _sendTask('clean'),
                          ),
                        ),
                        const SizedBox(width: 24),
                        Expanded(
                          child: _RightPane(
                            progress: _progress,
                            progressTitle: _progressTitle,
                            progressDetail: _progressDetail,
                            result: _result,
                            onOpenOutput: _openOutput,
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
      ),
    );
  }
}

class _TopBar extends StatelessWidget {
  const _TopBar({required this.onHelp});

  final VoidCallback onHelp;

  @override
  Widget build(BuildContext context) {
    return Container(
      height: 64,
      padding: const EdgeInsets.symmetric(horizontal: 26),
      decoration: const BoxDecoration(
        color: Color(0xfff8fbfb),
        border: Border(bottom: BorderSide(color: Color(0xffd9e7e8))),
        borderRadius: BorderRadius.vertical(top: Radius.circular(22)),
      ),
      child: Row(
        children: [
          Container(
            width: 34,
            height: 34,
            alignment: Alignment.center,
            decoration: BoxDecoration(
              color: const Color(0xff08735f),
              borderRadius: BorderRadius.circular(10),
              boxShadow: const [BoxShadow(color: Color(0x3808735f), blurRadius: 18, offset: Offset(0, 8))],
            ),
            child: const Text('Dr', style: TextStyle(color: Colors.white, fontWeight: FontWeight.w800)),
          ),
          const SizedBox(width: 12),
          const Text('Excel-Dr', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w800)),
          const Spacer(),
          _AppButton(label: '使用说明', onPressed: onHelp),
        ],
      ),
    );
  }
}

class _LeftPane extends StatelessWidget {
  const _LeftPane({
    required this.targetKind,
    required this.selectedPath,
    required this.busy,
    required this.onKindChanged,
    required this.onSelect,
    required this.onAnalyze,
    required this.onClean,
  });

  final String targetKind;
  final String selectedPath;
  final bool busy;
  final ValueChanged<String> onKindChanged;
  final VoidCallback onSelect;
  final VoidCallback onAnalyze;
  final VoidCallback onClean;

  @override
  Widget build(BuildContext context) {
    final hasTarget = selectedPath.isNotEmpty;
    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        const Text(
          '让卡顿的 Excel 变顺畅',
          style: TextStyle(fontSize: 25, height: 1.18, fontWeight: FontWeight.w900, color: Color(0xff102027)),
        ),
        const SizedBox(height: 4),
        const Text('选择文件，先检测问题，再新建清理版。', style: TextStyle(color: Color(0xff60757d))),
        const SizedBox(height: 12),
        _Surface(
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              const _SectionTitle(title: '处理对象', tag: '第 1 步'),
              const SizedBox(height: 10),
              _Segmented(
                value: targetKind,
                onChanged: onKindChanged,
                enabled: !busy,
              ),
              const SizedBox(height: 10),
              _FileBox(path: selectedPath, targetKind: targetKind),
              const SizedBox(height: 8),
              _AppButton(
                label: hasTarget ? '重新选择${targetKind == 'file' ? '文件' : '文件夹'}' : '选择${targetKind == 'file' ? '文件' : '文件夹'}',
                onPressed: onSelect,
                enabled: !busy,
                width: double.infinity,
              ),
            ],
          ),
        ),
        const SizedBox(height: 12),
        _Surface(
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              const _SectionTitle(title: '操作', tag: '按顺序执行'),
              const SizedBox(height: 12),
              _ActionBlock(
                tag: '第 2 步',
                title: '只检测问题，不修改文件',
                label: '仅检测',
                enabled: hasTarget && !busy,
                primary: false,
                onPressed: onAnalyze,
              ),
              const SizedBox(height: 12),
              _ActionBlock(
                tag: '第 3 步',
                title: '只新建文件，不修改原文件',
                label: '清理并新建',
                enabled: hasTarget && !busy,
                primary: true,
                onPressed: onClean,
              ),
            ],
          ),
        ),
        const SizedBox(height: 12),
        Container(
          padding: const EdgeInsets.all(14),
          decoration: BoxDecoration(color: const Color(0xffe7f5f1), borderRadius: BorderRadius.circular(16)),
          child: const Text('原文件保持不变，清理结果会生成一个新文件。', style: TextStyle(color: Color(0xff055f50), height: 1.45)),
        ),
      ],
    );
  }
}

class _RightPane extends StatelessWidget {
  const _RightPane({
    required this.progress,
    required this.progressTitle,
    required this.progressDetail,
    required this.result,
    required this.onOpenOutput,
  });

  final int progress;
  final String progressTitle;
  final String progressDetail;
  final TaskResult? result;
  final VoidCallback onOpenOutput;

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        _Surface(
          padding: const EdgeInsets.fromLTRB(20, 16, 20, 16),
          child: Column(
            children: [
              Row(
                children: [
                  const _StatusDot(),
                  const SizedBox(width: 10),
                  Expanded(child: Text(progressTitle, style: const TextStyle(fontWeight: FontWeight.w800))),
                  Text('$progress%', style: const TextStyle(color: Color(0xff08735f), fontSize: 16, fontWeight: FontWeight.w900)),
                ],
              ),
              const SizedBox(height: 14),
              ClipRRect(
                borderRadius: BorderRadius.circular(999),
                child: LinearProgressIndicator(
                  minHeight: 12,
                  value: progress / 100,
                  backgroundColor: const Color(0xffdfe9ec),
                  color: const Color(0xff0d806c),
                ),
              ),
              const SizedBox(height: 12),
              Row(
                children: [
                  Expanded(child: Text('当前：$progressDetail', overflow: TextOverflow.ellipsis, style: const TextStyle(color: Color(0xff60757d), fontSize: 12))),
                  const Text('请稍候', style: TextStyle(color: Color(0xff60757d), fontSize: 12)),
                ],
              ),
            ],
          ),
        ),
        const SizedBox(height: 12),
        Expanded(child: _ResultPanel(result: result, onOpenOutput: onOpenOutput)),
      ],
    );
  }
}

class _ResultPanel extends StatelessWidget {
  const _ResultPanel({required this.result, required this.onOpenOutput});

  final TaskResult? result;
  final VoidCallback onOpenOutput;

  @override
  Widget build(BuildContext context) {
    final rows = result?.rows ?? const <ResultRow>[];
    return _Surface(
      padding: const EdgeInsets.all(16),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          Row(
            children: [
              const Text('检测结果', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w900)),
              const Spacer(),
              _Tag(result?.message ?? '等待检测'),
            ],
          ),
          const SizedBox(height: 14),
          _Metrics(result: result),
          const SizedBox(height: 14),
          Expanded(
            child: _ResultTable(rows: rows, errors: result?.errors ?? const []),
          ),
          const SizedBox(height: 14),
          _OutputRow(path: result?.outputPath ?? '', onOpenOutput: onOpenOutput),
        ],
      ),
    );
  }
}

class _Metrics extends StatelessWidget {
  const _Metrics({required this.result});

  final TaskResult? result;

  @override
  Widget build(BuildContext context) {
    return Row(
      children: [
        _Metric(label: '扫描文件', value: result?.scannedFiles ?? 0, good: true),
        const SizedBox(width: 14),
        _Metric(label: '可清理对象', value: result?.suspiciousObjects ?? 0),
        const SizedBox(width: 14),
        _Metric(label: '损坏规则', value: result?.brokenRules ?? 0),
        const SizedBox(width: 14),
        _Metric(label: '输出文件', value: result?.outputFiles ?? 0),
      ],
    );
  }
}

class _ResultTable extends StatelessWidget {
  const _ResultTable({required this.rows, required this.errors});

  final List<ResultRow> rows;
  final List<String> errors;

  @override
  Widget build(BuildContext context) {
    final visibleRows = rows.take(7).toList();
    final visibleErrors = errors.take(4).toList();
    final hasEntries = visibleRows.isNotEmpty || visibleErrors.isNotEmpty;
    return Container(
      decoration: BoxDecoration(border: Border.all(color: const Color(0xffd9e7e8)), borderRadius: BorderRadius.circular(14)),
      clipBehavior: Clip.antiAlias,
      child: Column(
        children: [
          const _TableRow(location: '位置', issue: '问题', count: '数量', action: '处理方式', head: true),
          if (!hasEntries)
            const Expanded(child: Center(child: Text('选择文件后开始检测', style: TextStyle(color: Color(0xff60757d)))))
          else
            Expanded(
              child: ListView(
                padding: EdgeInsets.zero,
                children: [
                  for (final row in visibleRows)
                    _TableRow(
                      location: row.location,
                      issue: row.issue,
                      count: _formatNumber(row.count),
                      action: row.action,
                    ),
                  for (final error in visibleErrors)
                    _TableRow(location: '失败', issue: error, count: '1', action: '已隔离'),
                ],
              ),
            ),
        ],
      ),
    );
  }
}

class _OutputRow extends StatelessWidget {
  const _OutputRow({required this.path, required this.onOpenOutput});

  final String path;
  final VoidCallback onOpenOutput;

  @override
  Widget build(BuildContext context) {
    final hasOutput = path.isNotEmpty;
    return Container(
      height: 66,
      padding: const EdgeInsets.symmetric(horizontal: 16),
      decoration: BoxDecoration(color: const Color(0xff102027), borderRadius: BorderRadius.circular(14)),
      child: Row(
        children: [
          Expanded(
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                const Text('输出位置', style: TextStyle(color: Colors.white, fontWeight: FontWeight.w800)),
                const SizedBox(height: 8),
                Text(
                  hasOutput ? path : '完成清理后显示新文件位置',
                  maxLines: 1,
                  overflow: TextOverflow.ellipsis,
                  style: const TextStyle(color: Color(0xffc7d5d8), fontSize: 12),
                ),
              ],
            ),
          ),
          const SizedBox(width: 16),
          _AppButton(label: '打开输出文件夹', onPressed: onOpenOutput, enabled: hasOutput),
        ],
      ),
    );
  }
}

class _Surface extends StatelessWidget {
  const _Surface({required this.child, this.padding = EdgeInsets.zero});

  final Widget child;
  final EdgeInsets padding;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: padding,
      decoration: BoxDecoration(
        color: Colors.white,
        border: Border.all(color: const Color(0xffd9e7e8)),
        borderRadius: BorderRadius.circular(18),
        boxShadow: const [BoxShadow(color: Color(0x0a102027), blurRadius: 24, offset: Offset(0, 8))],
      ),
      child: child,
    );
  }
}

class _AppButton extends StatelessWidget {
  const _AppButton({
    required this.label,
    required this.onPressed,
    this.enabled = true,
    this.primary = false,
    this.width,
  });

  final String label;
  final VoidCallback onPressed;
  final bool enabled;
  final bool primary;
  final double? width;

  @override
  Widget build(BuildContext context) {
    final textColor = !enabled
        ? const Color(0xff93a5aa)
        : primary
            ? Colors.white
            : const Color(0xff27454d);
    final decoration = BoxDecoration(
      color: !enabled ? const Color(0xffedf2f3) : primary ? null : const Color(0xffeef5f6),
      gradient: enabled && primary
          ? const LinearGradient(
              begin: Alignment.topLeft,
              end: Alignment.bottomRight,
              colors: [Color(0xff08735f), Color(0xff055f50)],
            )
          : null,
      borderRadius: BorderRadius.circular(11),
      border: primary ? null : Border.all(color: enabled ? const Color(0xffabc6ca) : const Color(0xffd7e0e2)),
      boxShadow: enabled && primary
          ? const [BoxShadow(color: Color(0x3308735f), blurRadius: 20, offset: Offset(0, 10))]
          : null,
    );

    return SizedBox(
      width: width,
      height: primary ? 42 : 40,
      child: Material(
        color: Colors.transparent,
        child: InkWell(
          onTap: enabled ? onPressed : null,
          borderRadius: BorderRadius.circular(11),
          child: Ink(
            decoration: decoration,
            child: Center(
              child: Text(
                label,
                maxLines: 1,
                overflow: TextOverflow.ellipsis,
                style: TextStyle(color: textColor, fontWeight: FontWeight.w800, height: 1),
              ),
            ),
          ),
        ),
      ),
    );
  }
}

class _SectionTitle extends StatelessWidget {
  const _SectionTitle({required this.title, required this.tag});

  final String title;
  final String tag;

  @override
  Widget build(BuildContext context) {
    return Row(children: [Text(title, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w900)), const Spacer(), _Tag(tag)]);
  }
}

class _Segmented extends StatelessWidget {
  const _Segmented({required this.value, required this.onChanged, required this.enabled});

  final String value;
  final ValueChanged<String> onChanged;
  final bool enabled;

  @override
  Widget build(BuildContext context) {
    return Container(
      height: 40,
      padding: const EdgeInsets.all(4),
      decoration: BoxDecoration(color: const Color(0xffe9f1f2), borderRadius: BorderRadius.circular(13), border: Border.all(color: const Color(0xffabc6ca))),
      child: Row(
        children: [
          _Segment(label: '单文件', active: value == 'file', enabled: enabled, onTap: () => onChanged('file')),
          _Segment(label: '文件夹批量', active: value == 'folder', enabled: enabled, onTap: () => onChanged('folder')),
        ],
      ),
    );
  }
}

class _Segment extends StatelessWidget {
  const _Segment({required this.label, required this.active, required this.enabled, required this.onTap});

  final String label;
  final bool active;
  final bool enabled;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Expanded(
      child: InkWell(
        onTap: enabled ? onTap : null,
        borderRadius: BorderRadius.circular(10),
        child: Container(
          alignment: Alignment.center,
          decoration: BoxDecoration(color: active ? const Color(0xff08735f) : Colors.transparent, borderRadius: BorderRadius.circular(10)),
          child: Text(label, style: TextStyle(color: active ? Colors.white : const Color(0xff102027), fontWeight: FontWeight.w800)),
        ),
      ),
    );
  }
}

class _FileBox extends StatelessWidget {
  const _FileBox({required this.path, required this.targetKind});

  final String path;
  final String targetKind;

  @override
  Widget build(BuildContext context) {
    final empty = path.isEmpty;
    return Container(
      constraints: const BoxConstraints(minHeight: 74),
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(color: const Color(0xfff7fbfb), borderRadius: BorderRadius.circular(16), border: Border.all(color: const Color(0xffabc6ca))),
      child: Row(
        children: [
          Container(
            width: 34,
            height: 34,
            alignment: Alignment.center,
            decoration: BoxDecoration(color: const Color(0xffe7f5f1), borderRadius: BorderRadius.circular(10)),
            child: Text(targetKind == 'file' ? 'XLSX' : 'DIR', style: const TextStyle(color: Color(0xff08735f), fontSize: 11, fontWeight: FontWeight.w900)),
          ),
          const SizedBox(width: 10),
          Expanded(
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(empty ? '尚未选择' : _fileName(path), maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(fontWeight: FontWeight.w800)),
                const SizedBox(height: 6),
                Text(empty ? '请选择处理对象' : '将在原目录生成清理版', style: const TextStyle(color: Color(0xff60757d), fontSize: 12)),
              ],
            ),
          ),
        ],
      ),
    );
  }
}

class _ActionBlock extends StatelessWidget {
  const _ActionBlock({
    required this.tag,
    required this.title,
    required this.label,
    required this.enabled,
    required this.primary,
    required this.onPressed,
  });

  final String tag;
  final String title;
  final String label;
  final bool enabled;
  final bool primary;
  final VoidCallback onPressed;

  @override
  Widget build(BuildContext context) {
    return Container(
      constraints: const BoxConstraints(minHeight: 108),
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(color: const Color(0xfff7fbfb), borderRadius: BorderRadius.circular(14), border: Border.all(color: const Color(0xffd9e7e8))),
      child: Column(
        mainAxisAlignment: MainAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              SizedBox(width: 66, child: Align(alignment: Alignment.centerLeft, child: _Tag(tag))),
              const SizedBox(width: 12),
              Expanded(
                child: ConstrainedBox(
                  constraints: const BoxConstraints(minHeight: 26),
                  child: Align(
                    alignment: Alignment.centerLeft,
                    child: Text(
                      title,
                      maxLines: 2,
                      overflow: TextOverflow.ellipsis,
                      style: const TextStyle(fontSize: 14, height: 1.25, fontWeight: FontWeight.w800),
                    ),
                  ),
                ),
              ),
            ],
          ),
          const SizedBox(height: 10),
          Row(
            children: [
              const SizedBox(width: 66),
              const SizedBox(width: 12),
              Expanded(
                child: _AppButton(
                  label: label,
                  onPressed: onPressed,
                  enabled: enabled,
                  primary: primary,
                  width: double.infinity,
                ),
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _Metric extends StatelessWidget {
  const _Metric({required this.label, required this.value, this.good = false});

  final String label;
  final int value;
  final bool good;

  @override
  Widget build(BuildContext context) {
    return Expanded(
      child: Container(
        height: 88,
        padding: const EdgeInsets.all(16),
        decoration: BoxDecoration(
          color: good ? const Color(0xffe7f5f1) : const Color(0xfff7fbfb),
          borderRadius: BorderRadius.circular(16),
          border: Border.all(color: good ? const Color(0xffbce2d9) : const Color(0xffd9e7e8)),
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Text(label, style: const TextStyle(color: Color(0xff60757d), height: 1)),
            const SizedBox(height: 8),
            Text(_formatNumber(value), style: const TextStyle(fontSize: 28, height: 1, fontWeight: FontWeight.w900)),
          ],
        ),
      ),
    );
  }
}

class _TableRow extends StatelessWidget {
  const _TableRow({required this.location, required this.issue, required this.count, required this.action, this.head = false});

  final String location;
  final String issue;
  final String count;
  final String action;
  final bool head;

  @override
  Widget build(BuildContext context) {
    return Container(
      height: head ? 40 : 44,
      padding: const EdgeInsets.symmetric(horizontal: 14),
      decoration: BoxDecoration(color: head ? const Color(0xfff4f8f9) : Colors.white, border: const Border(bottom: BorderSide(color: Color(0xffd9e7e8)))),
      child: Row(
        children: [
          Expanded(flex: 11, child: Text(location, maxLines: 1, overflow: TextOverflow.ellipsis, style: _rowStyle(head))),
          Expanded(flex: 12, child: Text(issue, maxLines: 1, overflow: TextOverflow.ellipsis, style: _rowStyle(head))),
          Expanded(flex: 7, child: Text(count, maxLines: 1, overflow: TextOverflow.ellipsis, style: _rowStyle(head).copyWith(fontWeight: FontWeight.w900))),
          Expanded(flex: 12, child: Align(alignment: Alignment.centerLeft, child: _Tag(action))),
        ],
      ),
    );
  }

  TextStyle _rowStyle(bool head) => TextStyle(color: head ? const Color(0xff60757d) : const Color(0xff344a52), fontSize: head ? 12 : 14, fontWeight: head ? FontWeight.w800 : FontWeight.w500);
}

class _Tag extends StatelessWidget {
  const _Tag(this.text);

  final String text;

  @override
  Widget build(BuildContext context) {
    return Container(
      height: 26,
      padding: const EdgeInsets.symmetric(horizontal: 10),
      alignment: Alignment.center,
      decoration: BoxDecoration(color: const Color(0xffe7f5f1), borderRadius: BorderRadius.circular(999)),
      child: Text(text, maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(color: Color(0xff08735f), fontSize: 12, fontWeight: FontWeight.w900)),
    );
  }
}

class _StatusDot extends StatelessWidget {
  const _StatusDot();

  @override
  Widget build(BuildContext context) {
    return Container(
      width: 10,
      height: 10,
      decoration: BoxDecoration(
        color: const Color(0xff08735f),
        shape: BoxShape.circle,
        boxShadow: [BoxShadow(color: const Color(0xff08735f).withValues(alpha: 0.16), blurRadius: 0, spreadRadius: 6)],
      ),
    );
  }
}

class _HelpItem extends StatelessWidget {
  const _HelpItem({required this.title, required this.text});

  final String title;
  final String text;

  @override
  Widget build(BuildContext context) {
    return Container(
      width: 520,
      padding: const EdgeInsets.all(14),
      decoration: BoxDecoration(color: const Color(0xfff7fbfb), borderRadius: BorderRadius.circular(14), border: Border.all(color: const Color(0xffd9e7e8))),
      child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [Text(title, style: const TextStyle(fontWeight: FontWeight.w900)), const SizedBox(height: 8), Text(text, style: const TextStyle(color: Color(0xff60757d), height: 1.55))]),
    );
  }
}

String _fileName(String path) {
  final normalized = path.replaceAll('\\', '/');
  final parts = normalized.split('/').where((part) => part.isNotEmpty).toList();
  return parts.isEmpty ? path : parts.last;
}

String _formatNumber(int value) {
  final text = value.toString();
  final buffer = StringBuffer();
  for (var i = 0; i < text.length; i += 1) {
    final remaining = text.length - i;
    buffer.write(text[i]);
    if (remaining > 1 && remaining % 3 == 1) buffer.write(',');
  }
  return buffer.toString();
}

String _formatSize(int bytes) {
  if (bytes >= 1024 * 1024) return '${(bytes / 1024 / 1024).toStringAsFixed(1)} MB';
  if (bytes >= 1024) return '${(bytes / 1024).toStringAsFixed(1)} KB';
  return '$bytes B';
}
