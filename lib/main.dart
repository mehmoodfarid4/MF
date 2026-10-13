// ==============================================================================
// CampDesk Mobile - Cross-Platform Camp Accommodation & Manpower System
// Built with Flutter 3.x, Material 3, and Excel (.xlsx) Engine
// Compatible with Android (SAF Storage), iOS, and Web
// ==============================================================================

import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:url_launcher/url_launcher.dart';
import 'package:share_plus/share_plus.dart';
import 'package:path_provider/path_provider.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  SystemChrome.setSystemUIOverlayStyle(
    const SystemUiOverlayStyle(
      statusBarColor: Colors.transparent,
      statusBarIconBrightness: Brightness.light,
      systemNavigationBarColor: Color(0xFF0F172A),
    ),
  );
  runApp(const CampDeskApp());
}

// ------------------------------------------------------------------------------
// 1. DATA MODEL (16 SPREADSHEET COLUMNS)
// ------------------------------------------------------------------------------

enum WorkerStatus { present, vacation, transferred, exit }

extension WorkerStatusExt on WorkerStatus {
  String get label {
    switch (this) {
      case WorkerStatus.present:
        return 'Present';
      case WorkerStatus.vacation:
        return 'Vacation';
      case WorkerStatus.transferred:
        return 'Transferred';
      case WorkerStatus.exit:
        return 'Exit';
    }
  }

  Color get color {
    switch (this) {
      case WorkerStatus.present:
        return const Color(0xFF10B981); // Emerald
      case WorkerStatus.vacation:
        return const Color(0xFFF59E0B); // Amber / Yellow Highlight
      case WorkerStatus.transferred:
        return const Color(0xFF3B82F6); // Blue
      case WorkerStatus.exit:
        return const Color(0xFFEF4444); // Rose
    }
  }
}

class CampPersonnel {
  String sl;
  String roomNo;
  String ffId;
  String position;
  String iqamaNo;
  String employeeName;
  String sponsorName;
  String nationality;
  String mobileNo;
  String religion;
  String company;
  WorkerStatus status;
  String checkIn;
  String vacant;
  String project;
  String camp;

  CampPersonnel({
    required this.sl,
    required this.roomNo,
    required this.ffId,
    required this.position,
    required this.iqamaNo,
    required this.employeeName,
    required this.sponsorName,
    required this.nationality,
    required this.mobileNo,
    required this.religion,
    required this.company,
    required this.status,
    required this.checkIn,
    required this.vacant,
    required this.project,
    required this.camp,
  });

  CampPersonnel copyWith({
    String? sl,
    String? roomNo,
    String? ffId,
    String? position,
    String? iqamaNo,
    String? employeeName,
    String? sponsorName,
    String? nationality,
    String? mobileNo,
    String? religion,
    String? company,
    WorkerStatus? status,
    String? checkIn,
    String? vacant,
    String? project,
    String? camp,
  }) {
    return CampPersonnel(
      sl: sl ?? this.sl,
      roomNo: roomNo ?? this.roomNo,
      ffId: ffId ?? this.ffId,
      position: position ?? this.position,
      iqamaNo: iqamaNo ?? this.iqamaNo,
      employeeName: employeeName ?? this.employeeName,
      sponsorName: sponsorName ?? this.sponsorName,
      nationality: nationality ?? this.nationality,
      mobileNo: mobileNo ?? this.mobileNo,
      religion: religion ?? this.religion,
      company: company ?? this.company,
      status: status ?? this.status,
      checkIn: checkIn ?? this.checkIn,
      vacant: vacant ?? this.vacant,
      project: project ?? this.project,
      camp: camp ?? this.camp,
    );
  }
}

// ------------------------------------------------------------------------------
// 2. EXCEL REPOSITORY & RESILIENT INGESTION SERVICE
// ------------------------------------------------------------------------------

class ExcelRosterService {
  static const List<String> canonicalHeaders = [
    'SL', 'Room No', 'FF ID', 'Position', 'Iqama No.',
    'Employee Name', 'Sponsor Name', 'Nationality', 'Mobile No.',
    'Religion', 'Company', 'Status', 'Check in', 'Vacant',
    'Project', 'Camp'
  ];

  static String _cleanKey(String s) =>
      s.toLowerCase().replaceAll(RegExp(r'[^a-z0-9]'), '');

  /// Resiliently parses Excel bytes with fuzzy column mapping
  static List<CampPersonnel> parseExcelBytes(Uint8List bytes) {
    final excel = Excel.decodeBytes(bytes);
    if (excel.tables.isEmpty) return [];

    final sheet = excel.tables.values.first;
    if (sheet.maxRows < 2) return [];

    // Header extraction & indexing
    final headerRow = sheet.rows.first;
    final Map<int, String> colMap = {};
    for (int col = 0; col < headerRow.length; col++) {
      final cellVal = headerRow[col]?.value?.toString().trim() ?? '';
      if (cellVal.isNotEmpty) colMap[col] = cellVal;
    }

    final List<CampPersonnel> result = [];

    for (int r = 1; r < sheet.rows.length; r++) {
      final row = sheet.rows[r];
      if (row.isEmpty) continue;

      final Map<String, String> rowDict = {};
      for (int c = 0; c < row.length; c++) {
        final colHeader = colMap[c];
        if (colHeader != null) {
          rowDict[_cleanKey(colHeader)] = row[c]?.value?.toString().trim() ?? '';
        }
      }

      String findVal(List<String> variants, String fallback) {
        for (final v in variants) {
          final cleanV = _cleanKey(v);
          for (final entry in rowDict.entries) {
            if (entry.key == cleanV || entry.key.contains(cleanV)) {
              if (entry.value.isNotEmpty) return entry.value;
            }
          }
        }
        return fallback;
      }

      final name = findVal(['employeename', 'name', 'worker'], '');
      final ffId = findVal(['ffid', 'id', 'emp_id'], '');
      // Skip empty ghost rows
      if (name.isEmpty && ffId.isEmpty) continue;

      final rawStatus = findVal(['status', 'state'], 'Present').toLowerCase();
      WorkerStatus status = WorkerStatus.present;
      if (rawStatus.contains('vacation') || rawStatus.contains('leave')) {
        status = WorkerStatus.vacation;
      } else if (rawStatus.contains('transfer')) {
        status = WorkerStatus.transferred;
      } else if (rawStatus.contains('exit')) {
        status = WorkerStatus.exit;
      }

      result.add(
        CampPersonnel(
          sl: findVal(['sl', 'sno', 'no'], '${result.length + 1}'),
          roomNo: findVal(['roomno', 'room', 'bed'], 'Unassigned'),
          ffId: ffId.isEmpty ? 'FF-${2000 + result.length}' : ffId,
          position: findVal(['position', 'trade', 'role'], 'General Worker'),
          iqamaNo: findVal(['iqamano', 'iqama', 'nationalid'], ''),
          employeeName: name.isEmpty ? 'Unnamed Worker' : name,
          sponsorName: findVal(['sponsorname', 'sponsor'], 'Al-Rashid Contracting'),
          nationality: findVal(['nationality', 'country'], 'Unspecified'),
          mobileNo: findVal(['mobileno', 'mobile', 'phone'], ''),
          religion: findVal(['religion'], 'Muslim'),
          company: findVal(['company', 'contractor'], 'Prime Infra'),
          status: status,
          checkIn: findVal(['checkin', 'date'], '2026-01-01'),
          vacant: findVal(['vacant'], 'Occupied'),
          project: findVal(['project', 'site'], 'Metro Line 4'),
          camp: findVal(['camp'], 'East Valley Camp'),
        ),
      );
    }

    return result;
  }

  /// Encodes the current camp roster into a formatted .xlsx byte stream
  static Uint8List encodeToExcelBytes(List<CampPersonnel> roster) {
    final excel = Excel.createExcel();
    final sheet = excel['Camp Roster'];
    excel.setDefaultSheet('Camp Roster');

    // Header Row with Styling
    sheet.appendRow(canonicalHeaders.map((h) => TextCellValue(h)).toList());

    for (final p in roster) {
      sheet.appendRow([
        TextCellValue(p.sl),
        TextCellValue(p.roomNo),
        TextCellValue(p.ffId),
        TextCellValue(p.position),
        TextCellValue(p.iqamaNo),
        TextCellValue(p.employeeName),
        TextCellValue(p.sponsorName),
        TextCellValue(p.nationality),
        TextCellValue(p.mobileNo),
        TextCellValue(p.religion),
        TextCellValue(p.company),
        TextCellValue(p.status.label),
        TextCellValue(p.checkIn),
        TextCellValue(p.vacant),
        TextCellValue(p.project),
        TextCellValue(p.camp),
      ]);
    }

    return Uint8List.fromList(excel.encode() ?? []);
  }
}

// ------------------------------------------------------------------------------
// 3. MAIN APPLICATION & THEME
// ------------------------------------------------------------------------------

class CampDeskApp extends StatelessWidget {
  const CampDeskApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'CampDesk Mobile',
      debugShowCheckedModeBanner: false,
      themeMode: ThemeMode.dark,
      darkTheme: ThemeData(
        useMaterial3: true,
        brightness: Brightness.dark,
        scaffoldBackgroundColor: const Color(0xFF090D16),
        colorScheme: const ColorScheme.dark(
          primary: Color(0xFF38BDF8), // Cyan / Sky blue
          secondary: Color(0xFF6366F1), // Indigo
          surface: Color(0xFF131B2E),
          surfaceContainer: Color(0xFF1E293B),
          onSurface: Color(0xFFF1F5F9),
          error: Color(0xFFF43F5E),
        ),
        appBarTheme: const AppBarTheme(
          backgroundColor: Color(0xFF0F172A),
          elevation: 0,
          scrolledUnderElevation: 0,
        ),
        cardTheme: CardTheme(
          color: const Color(0xFF131C31),
          elevation: 2,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(14),
            side: const BorderSide(color: Color(0xFF26334D), width: 1),
          ),
        ),
      ),
      home: const CampHomeScreen(),
    );
  }
}

// ------------------------------------------------------------------------------
// 4. HOME SCREEN (ROSTER, CARDS, FILTERS, FAB)
// ------------------------------------------------------------------------------

class CampHomeScreen extends StatefulWidget {
  const CampHomeScreen({super.key});

  @override
  State<CampHomeScreen> createState() => _CampHomeScreenState();
}

class _CampHomeScreenState extends State<CampHomeScreen> {
  List<CampPersonnel> _roster = [];
  String _activeFilePath = 'Camp_Accommodation_Roster.xlsx';
  bool _hasUnsavedChanges = false;
  String _searchQuery = '';
  String _selectedFilter = 'All'; // All, Present, Vacation, Transferred, Exit
  int _selectedTabIndex = 0; // 0: Roster, 1: Room Occupancy, 2: Analytics

  @override
  void initState() {
    super.initState();
    _loadSampleRoster();
  }

  void _loadSampleRoster() {
    setState(() {
      _roster = [
        CampPersonnel(
          sl: '1',
          roomNo: 'B-101',
          ffId: 'FF-2041',
          position: 'Civil Foreman',
          iqamaNo: '2498112034',
          employeeName: 'Tariq Mahmood',
          sponsorName: 'Al-Rashid Contracting',
          nationality: 'Pakistan',
          mobileNo: '055-1234567',
          religion: 'Muslim',
          company: 'Prime Infra',
          status: WorkerStatus.present,
          checkIn: '2026-01-15',
          vacant: 'Occupied',
          project: 'Metro Line 4',
          camp: 'East Valley Camp',
        ),
        CampPersonnel(
          sl: '2',
          roomNo: 'B-101',
          ffId: 'FF-2042',
          position: 'Safety Officer',
          iqamaNo: '2389145601',
          employeeName: 'Suresh Kumar',
          sponsorName: 'Global Logistics',
          nationality: 'India',
          mobileNo: '054-9876543',
          religion: 'Hindu',
          company: 'Apex Contracting',
          status: WorkerStatus.vacation, // Vacation highlight
          checkIn: '2025-11-20',
          vacant: 'Occupied',
          project: 'Airport Terminal 3',
          camp: 'East Valley Camp',
        ),
        CampPersonnel(
          sl: '3',
          roomNo: 'A-204',
          ffId: 'FF-2088',
          position: 'Heavy Equipment Operator',
          iqamaNo: '2501984421',
          employeeName: 'Abdul Rahman Al-Masri',
          sponsorName: 'Al-Rashid Contracting',
          nationality: 'Egypt',
          mobileNo: '056-5544332',
          religion: 'Muslim',
          company: 'Prime Infra',
          status: WorkerStatus.present,
          checkIn: '2026-03-01',
          vacant: 'Occupied',
          project: 'Metro Line 4',
          camp: 'East Valley Camp',
        ),
        CampPersonnel(
          sl: '4',
          roomNo: 'C-302',
          ffId: 'FF-2110',
          position: 'Master Electrician',
          iqamaNo: '2412896503',
          employeeName: 'Danilo Santos',
          sponsorName: 'Prime ElectroMech',
          nationality: 'Philippines',
          mobileNo: '053-8877665',
          religion: 'Christian',
          company: 'Delta Power',
          status: WorkerStatus.present,
          checkIn: '2025-08-10',
          vacant: 'Occupied',
          project: 'Solar Substation B',
          camp: 'East Valley Camp',
        ),
        CampPersonnel(
          sl: '5',
          roomNo: 'A-204',
          ffId: 'FF-2145',
          position: 'Steel Fixer',
          iqamaNo: '2398457712',
          employeeName: 'Mohammad Shahidul',
          sponsorName: 'Gulf Builders',
          nationality: 'Bangladesh',
          mobileNo: '057-2233441',
          religion: 'Muslim',
          company: 'BuildCorp Global',
          status: WorkerStatus.vacation,
          checkIn: '2025-12-05',
          vacant: 'Occupied',
          project: 'Metro Line 4',
          camp: 'East Valley Camp',
        ),
      ];
      _hasUnsavedChanges = false;
    });
  }

  // ---------------- Excel File Picking & Ingestion ----------------
  Future<void> _pickExcelFile() async {
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
        withData: true,
      );

      if (result != null && result.files.isNotEmpty) {
        final file = result.files.first;
        final bytes = file.bytes;
        if (bytes != null) {
          final parsed = ExcelRosterService.parseExcelBytes(bytes);
          if (parsed.isEmpty) {
            _showSnackBar('No valid rows found in selected Excel file', isError: true);
            return;
          }
          setState(() {
            _roster = parsed;
            _activeFilePath = file.name;
            _hasUnsavedChanges = false;
          });
          _showSnackBar('Successfully loaded ${parsed.length} personnel from ${file.name}');
        }
      }
    } catch (e) {
      _showSnackBar('Permission error or file locked: ${e.toString()}', isError: true);
    }
  }

  // ---------------- Export & Share Excel Roster ----------------
  Future<void> _exportExcel() async {
    try {
      final bytes = ExcelRosterService.encodeToExcelBytes(_roster);
      final tempDir = await getTemporaryDirectory();
      final file = File('${tempDir.path}/$_activeFilePath');
      await file.writeAsBytes(bytes);

      setState(() => _hasUnsavedChanges = false);
      await Share.shareXFiles([XFile(file.path)], text: 'Camp Accommodation Roster - ${DateTime.now().toString().split(' ')[0]}');
      _showSnackBar('Excel exported and ready to share!');
    } catch (e) {
      _showSnackBar('Export failed: $e', isError: true);
    }
  }

  // ---------------- Communication Intents ----------------
  Future<void> _makePhoneCall(String phone) async {
    final cleanPhone = phone.replaceAll(RegExp(r'[^0-9+]'), '');
    final uri = Uri.parse('tel:$cleanPhone');
    if (await canLaunchUrl(uri)) {
      await launchUrl(uri);
    } else {
      _showSnackBar('Cannot dial $phone on this device');
    }
  }

  Future<void> _sendWhatsApp(String phone, String name, String room) async {
    final cleanPhone = phone.replaceAll(RegExp(r'[^0-9]'), '');
    final msg = Uri.encodeComponent(
      'Hello $name (Room $room),\nThis is Camp Administration office regarding your accommodation.',
    );
    final uri = Uri.parse('https://wa.me/$cleanPhone?text=$msg');
    if (await canLaunchUrl(uri)) {
      await launchUrl(uri, mode: LaunchMode.externalApplication);
    } else {
      _showSnackBar('Cannot open WhatsApp for $phone');
    }
  }

  void _showSnackBar(String text, {bool isError = false}) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(text),
        backgroundColor: isError ? const Color(0xFFEF4444) : const Color(0xFF10B981),
        behavior: SnackBarBehavior.floating,
      ),
    );
  }

  // ---------------- Filtered List ----------------
  List<CampPersonnel> get _filteredRoster {
    return _roster.where((p) {
      final matchesFilter = _selectedFilter == 'All' || p.status.label == _selectedFilter;
      if (!matchesFilter) return false;

      if (_searchQuery.isEmpty) return true;
      final q = _searchQuery.toLowerCase();
      return p.employeeName.toLowerCase().contains(q) ||
          p.ffId.toLowerCase().contains(q) ||
          p.roomNo.toLowerCase().contains(q) ||
          p.position.toLowerCase().contains(q) ||
          p.iqamaNo.toLowerCase().contains(q) ||
          p.nationality.toLowerCase().contains(q);
    }).toList();
  }

  // ---------------- KPI Computations ----------------
  int get totalPersonnel => _roster.length;
  int get presentCount => _roster.where((p) => p.status == WorkerStatus.present).length;
  int get vacationCount => _roster.where((p) => p.status == WorkerStatus.vacation).length;
  int get occupiedRoomsCount => _roster.map((p) => p.roomNo).toSet().length;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              'CampDesk Mobile',
              style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Colors.white),
            ),
            Text(
              _activeFilePath,
              style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8)),
              overflow: TextOverflow.ellipsis,
            ),
          ],
        ),
        actions: [
          IconButton(
            tooltip: 'Open Excel File',
            icon: const Icon(Icons.folder_open_rounded, color: Color(0xFF38BDF8)),
            onPressed: _pickExcelFile,
          ),
          IconButton(
            tooltip: 'Share / Export Excel',
            icon: const Icon(Icons.share_rounded, color: Color(0xFF10B981)),
            onPressed: _exportExcel,
          ),
          IconButton(
            tooltip: 'Reset to Sample Data',
            icon: const Icon(Icons.restore_rounded, color: Color(0xFF94A3B8)),
            onPressed: _loadSampleRoster,
          ),
        ],
      ),
      body: _buildCurrentTabBody(),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () => _openAddWorkerSheet(),
        backgroundColor: const Color(0xFF2563EB),
        foregroundColor: Colors.white,
        icon: const Icon(Icons.person_add_rounded),
        label: const Text('Add Worker'),
      ),
      bottomNavigationBar: NavigationBar(
        selectedIndex: _selectedTabIndex,
        onDestinationSelected: (idx) => setState(() => _selectedTabIndex = idx),
        backgroundColor: const Color(0xFF0F172A),
        indicatorColor: const Color(0xFF1E3A8A),
        destinations: const [
          NavigationDestination(
            icon: Icon(Icons.people_alt_outlined),
            selectedIcon: Icon(Icons.people_alt_rounded, color: Color(0xFF60A5FA)),
            label: 'Roster',
          ),
          NavigationDestination(
            icon: Icon(Icons.meeting_room_outlined),
            selectedIcon: Icon(Icons.meeting_room_rounded, color: Color(0xFF60A5FA)),
            label: 'Rooms',
          ),
          NavigationDestination(
            icon: Icon(Icons.analytics_outlined),
            selectedIcon: Icon(Icons.analytics_rounded, color: Color(0xFF60A5FA)),
            label: 'Muster Roll',
          ),
        ],
      ),
    );
  }

  Widget _buildCurrentTabBody() {
    switch (_selectedTabIndex) {
      case 1:
        return _buildRoomsView();
      case 2:
        return _buildMusterRollView();
      default:
        return _buildRosterView();
    }
  }

  // ---------------- Tab 1: Roster View ----------------
  Widget _buildRosterView() {
    return Column(
      children: [
        // KPI Overview Cards
        _buildKpiSection(),

        // Search & Filter Bar
        _buildSearchFilterSection(),

        // Personnel List View
        Expanded(
          child: _filteredRoster.isEmpty
              ? _buildEmptyState()
              : ListView.builder(
                  padding: const EdgeInsets.fromLTRB(14, 4, 14, 88),
                  itemCount: _filteredRoster.length,
                  itemBuilder: (context, index) {
                    final item = _filteredRoster[index];
                    return _buildWorkerCard(item);
                  },
                ),
        ),
      ],
    );
  }

  Widget _buildKpiSection() {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
      child: Row(
        children: [
          _buildKpiCard('Total Personnel', '$totalPersonnel', const Color(0xFF38BDF8), Icons.groups_rounded),
          const SizedBox(width: 8),
          _buildKpiCard('Present', '$presentCount', const Color(0xFF10B981), Icons.check_circle_rounded),
          const SizedBox(width: 8),
          _buildKpiCard('Vacation', '$vacationCount', const Color(0xFFF59E0B), Icons.beach_access_rounded),
          const SizedBox(width: 8),
          _buildKpiCard('Rooms', '$occupiedRoomsCount', const Color(0xFF818CF8), Icons.bed_rounded),
        ],
      ),
    );
  }

  Widget _buildKpiCard(String label, String value, Color color, IconData icon) {
    return Expanded(
      child: Container(
        padding: const EdgeInsets.symmetric(vertical: 10, horizontal: 8),
        decoration: BoxDecoration(
          color: const Color(0xFF131C31),
          borderRadius: BorderRadius.circular(10),
          border: Border.all(color: color.withOpacity(0.2)),
        ),
        child: Column(
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                Icon(icon, size: 14, color: color),
                const SizedBox(width: 4),
                Text(
                  value,
                  style: TextStyle(
                    fontSize: 16,
                    fontWeight: FontWeight.bold,
                    color: color,
                    fontFeatures: const [FontFeature.tabularFigures()],
                  ),
                ),
              ],
            ),
            const SizedBox(height: 2),
            Text(
              label,
              style: const TextStyle(fontSize: 10, color: Color(0xFF94A3B8)),
              maxLines: 1,
              overflow: TextOverflow.ellipsis,
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildSearchFilterSection() {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 4),
      child: Column(
        children: [
          // Search Input
          TextField(
            onChanged: (val) => setState(() => _searchQuery = val),
            decoration: InputDecoration(
              hintText: 'Search Name, FF ID, Room, Iqama, Trade...',
              hintStyle: const TextStyle(fontSize: 13, color: Color(0xFF64748B)),
              prefixIcon: const Icon(Icons.search_rounded, size: 20, color: Color(0xFF94A3B8)),
              suffixIcon: _searchQuery.isNotEmpty
                  ? IconButton(
                      icon: const Icon(Icons.clear, size: 18),
                      onPressed: () => setState(() => _searchQuery = ''),
                    )
                  : null,
              contentPadding: const EdgeInsets.symmetric(vertical: 10),
              filled: true,
              fillColor: const Color(0xFF131C31),
              border: OutlineInputBorder(
                borderRadius: BorderRadius.circular(10),
                borderSide: const BorderSide(color: Color(0xFF26334D)),
              ),
              enabledBorder: OutlineInputBorder(
                borderRadius: BorderRadius.circular(10),
                borderSide: const BorderSide(color: Color(0xFF26334D)),
              ),
            ),
          ),
          const SizedBox(height: 8),

          // Filter Chips
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            child: Row(
              children: ['All', 'Present', 'Vacation', 'Transferred', 'Exit'].map((status) {
                final isSelected = _selectedFilter == status;
                return Padding(
                  padding: const EdgeInsets.only(right: 6),
                  child: FilterChip(
                    label: Text(status),
                    selected: isSelected,
                    onSelected: (selected) {
                      setState(() => _selectedFilter = status);
                    },
                    selectedColor: status == 'Vacation'
                        ? const Color(0xFF78350F)
                        : const Color(0xFF1E3A8A),
                    checkmarkColor: Colors.white,
                    labelStyle: TextStyle(
                      fontSize: 12,
                      color: isSelected
                          ? Colors.white
                          : (status == 'Vacation' ? const Color(0xFFFBBF24) : const Color(0xFF94A3B8)),
                      fontWeight: isSelected ? FontWeight.bold : FontWeight.normal,
                    ),
                    side: BorderSide(
                      color: status == 'Vacation' && isSelected
                          ? const Color(0xFFF59E0B)
                          : const Color(0xFF334155),
                    ),
                  ),
                );
              }).toList(),
            ),
          ),
        ],
      ),
    );
  }

  // ---------------- Worker Card with Vacation Highlight ----------------
  Widget _buildWorkerCard(CampPersonnel p) {
    final isVacation = p.status == WorkerStatus.vacation;

    return Card(
      margin: const EdgeInsets.only(bottom: 10),
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(12),
        // Distinct amber outline if Vacation
        side: BorderSide(
          color: isVacation ? const Color(0xFFF59E0B) : const Color(0xFF26334D),
          width: isVacation ? 1.8 : 1,
        ),
      ),
      color: isVacation ? const Color(0xFF1C1917) : const Color(0xFF131C31),
      child: InkWell(
        borderRadius: BorderRadius.circular(12),
        onTap: () => _openEditWorkerSheet(p),
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // Header Row: Room Badge, Name, Status Badge
              Row(
                children: [
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                    decoration: BoxDecoration(
                      color: const Color(0xFF1E293B),
                      borderRadius: BorderRadius.circular(6),
                      border: Border.all(color: const Color(0xFF334155)),
                    ),
                    child: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        const Icon(Icons.meeting_room_outlined, size: 12, color: Color(0xFF38BDF8)),
                        const SizedBox(width: 4),
                        Text(
                          p.roomNo,
                          style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Colors.white),
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: Text(
                      p.employeeName,
                      style: const TextStyle(fontSize: 15, fontWeight: FontWeight.w600, color: Colors.white),
                      overflow: TextOverflow.ellipsis,
                    ),
                  ),
                  // Status Badge
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                    decoration: BoxDecoration(
                      color: p.status.color.withOpacity(0.15),
                      borderRadius: BorderRadius.circular(6),
                      border: Border.all(color: p.status.color.withOpacity(0.4)),
                    ),
                    child: Text(
                      p.status.label,
                      style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: p.status.color),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 8),

              // Details: FF ID, Position, Company
              Row(
                children: [
                  Text('ID: ${p.ffId}', style: const TextStyle(fontSize: 12, color: Color(0xFF38BDF8), fontWeight: FontWeight.w500)),
                  const Text(' • ', style: TextStyle(color: Color(0xFF64748B))),
                  Expanded(
                    child: Text(
                      p.position,
                      style: const TextStyle(fontSize: 12, color: Color(0xFFCBD5E1)),
                      overflow: TextOverflow.ellipsis,
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 4),
              Row(
                children: [
                  Text(p.company, style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8))),
                  const Text(' | ', style: TextStyle(color: Color(0xFF475569))),
                  Text(p.nationality, style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8))),
                  if (p.iqamaNo.isNotEmpty) ...[
                    const Text(' | ', style: TextStyle(color: Color(0xFF475569))),
                    Text('Iqama: ${p.iqamaNo}', style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8))),
                  ]
                ],
              ),
              const Divider(color: Color(0xFF1E293B), height: 16),

              // Action Buttons Row: Direct Call, WhatsApp, Room change
              Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Text(
                    p.mobileNo.isNotEmpty ? p.mobileNo : 'No mobile registered',
                    style: TextStyle(
                      fontSize: 12,
                      color: p.mobileNo.isNotEmpty ? const Color(0xFF60A5FA) : const Color(0xFF64748B),
                    ),
                  ),
                  Row(
                    children: [
                      if (p.mobileNo.isNotEmpty) ...[
                        IconButton.filledTonal(
                          iconSize: 16,
                          constraints: const BoxConstraints(minWidth: 32, minHeight: 32),
                          padding: EdgeInsets.zero,
                          tooltip: 'Dial Call',
                          icon: const Icon(Icons.call_rounded, color: Color(0xFF38BDF8)),
                          onPressed: () => _makePhoneCall(p.mobileNo),
                        ),
                        const SizedBox(width: 6),
                        IconButton.filledTonal(
                          iconSize: 16,
                          constraints: const BoxConstraints(minWidth: 32, minHeight: 32),
                          padding: EdgeInsets.zero,
                          tooltip: 'Send WhatsApp',
                          icon: const Icon(Icons.chat_bubble_rounded, color: Color(0xFF10B981)),
                          onPressed: () => _sendWhatsApp(p.mobileNo, p.employeeName, p.roomNo),
                        ),
                        const SizedBox(width: 6),
                      ],
                      IconButton.filledTonal(
                        iconSize: 16,
                        constraints: const BoxConstraints(minWidth: 32, minHeight: 32),
                        padding: EdgeInsets.zero,
                        tooltip: 'Edit Record',
                        icon: const Icon(Icons.edit_note_rounded, color: Color(0xFF94A3B8)),
                        onPressed: () => _openEditWorkerSheet(p),
                      ),
                    ],
                  ),
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildEmptyState() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          const Icon(Icons.search_off_rounded, size: 48, color: Color(0xFF64748B)),
          const SizedBox(height: 12),
          const Text('No personnel matching criteria', style: TextStyle(fontSize: 15, color: Color(0xFF94A3B8))),
          const SizedBox(height: 8),
          TextButton.icon(
            onPressed: () => setState(() {
              _searchQuery = '';
              _selectedFilter = 'All';
            }),
            icon: const Icon(Icons.refresh_rounded),
            label: const Text('Clear Filters'),
          ),
        ],
      ),
    );
  }

  // ---------------- Tab 2: Room Occupancy Breakdown ----------------
  Widget _buildRoomsView() {
    final Map<String, List<CampPersonnel>> roomGroups = {};
    for (final p in _roster) {
      roomGroups.putIfAbsent(p.roomNo, () => []).add(p);
    }

    final sortedRooms = roomGroups.keys.toList()..sort();

    return ListView.builder(
      padding: const EdgeInsets.all(14),
      itemCount: sortedRooms.length,
      itemBuilder: (context, index) {
        final room = sortedRooms[index];
        final occupants = roomGroups[room]!;
        final vacationInRoom = occupants.where((p) => p.status == WorkerStatus.vacation).length;

        return Card(
          margin: const EdgeInsets.only(bottom: 12),
          child: Padding(
            padding: const EdgeInsets.all(14),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Row(
                      children: [
                        const Icon(Icons.meeting_room_rounded, color: Color(0xFF38BDF8)),
                        const SizedBox(width: 8),
                        Text('Room $room', style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold, color: Colors.white)),
                      ],
                    ),
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                      decoration: BoxDecoration(
                        color: const Color(0xFF1E293B),
                        borderRadius: BorderRadius.circular(6),
                      ),
                      child: Text(
                        '${occupants.length} Occupants' + (vacationInRoom > 0 ? ' ($vacationInRoom away)' : ''),
                        style: const TextStyle(fontSize: 12, color: Color(0xFF94A3B8)),
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 10),
                ...occupants.map((occ) => Padding(
                      padding: const EdgeInsets.symmetric(vertical: 4),
                      child: Row(
                        children: [
                          Icon(
                            occ.status == WorkerStatus.vacation ? Icons.beach_access_rounded : Icons.person_rounded,
                            size: 14,
                            color: occ.status.color,
                          ),
                          const SizedBox(width: 6),
                          Expanded(
                            child: Text(
                              '${occ.employeeName} (${occ.ffId})',
                              style: TextStyle(
                                fontSize: 13,
                                color: occ.status == WorkerStatus.vacation ? const Color(0xFFFBBF24) : Colors.white,
                              ),
                            ),
                          ),
                          Text(occ.position, style: const TextStyle(fontSize: 12, color: Color(0xFF64748B))),
                        ],
                      ),
                    )),
              ],
            ),
          ),
        );
      },
    );
  }

  // ---------------- Tab 3: Daily Muster Roll Report ----------------
  Widget _buildMusterRollView() {
    return Padding(
      padding: const EdgeInsets.all(16),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            'Daily Muster Roll & Site Summary',
            style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Colors.white),
          ),
          const SizedBox(height: 4),
          const Text(
            'Instant headcount report for camp safety, catering rosters, and manpower audits.',
            style: TextStyle(fontSize: 12, color: Color(0xFF94A3B8)),
          ),
          const SizedBox(height: 16),
          Card(
            child: Padding(
              padding: const EdgeInsets.all(14),
              child: Column(
                children: [
                  _buildMusterRow('Total Camp Roster', '$totalPersonnel', const Color(0xFF38BDF8)),
                  const Divider(color: Color(0xFF1E293B)),
                  _buildMusterRow('Present On Site', '$presentCount', const Color(0xFF10B981)),
                  const Divider(color: Color(0xFF1E293B)),
                  _buildMusterRow('On Vacation Leave', '$vacationCount', const Color(0xFFF59E0B)),
                  const Divider(color: Color(0xFF1E293B)),
                  _buildMusterRow(
                    'Transferred / Exit',
                    '${_roster.where((p) => p.status == WorkerStatus.transferred || p.status == WorkerStatus.exit).length}',
                    const Color(0xFFEF4444),
                  ),
                  const Divider(color: Color(0xFF1E293B)),
                  _buildMusterRow('Occupied Rooms', '$occupiedRoomsCount', const Color(0xFF818CF8)),
                ],
              ),
            ),
          ),
          const SizedBox(height: 20),
          SizedBox(
            width: double.infinity,
            child: ElevatedButton.icon(
              onPressed: _exportExcel,
              icon: const Icon(Icons.file_download_rounded),
              label: const Text('Export Daily Muster Roll (.xlsx)'),
              style: ElevatedButton.styleFrom(
                backgroundColor: const Color(0xFF10B981),
                foregroundColor: Colors.white,
                padding: const EdgeInsets.symmetric(vertical: 14),
                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildMusterRow(String label, String value, Color color) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 14, color: Color(0xFFCBD5E1))),
          Text(
            value,
            style: TextStyle(
              fontSize: 16,
              fontWeight: FontWeight.bold,
              color: color,
              fontFeatures: const [FontFeature.tabularFigures()],
            ),
          ),
        ],
      ),
    );
  }

  // ---------------- Worker Bottom Sheet (Add / Edit) ----------------
  void _openAddWorkerSheet() {
    _showWorkerFormModal(null);
  }

  void _openEditWorkerSheet(CampPersonnel personnel) {
    _showWorkerFormModal(personnel);
  }

  void _showWorkerFormModal(CampPersonnel? worker) {
    final isEditing = worker != null;
    final nameCtrl = TextEditingController(text: worker?.employeeName ?? '');
    final ffIdCtrl = TextEditingController(text: worker?.ffId ?? 'FF-${2150 + _roster.length}');
    final roomCtrl = TextEditingController(text: worker?.roomNo ?? 'A-101');
    final positionCtrl = TextEditingController(text: worker?.position ?? 'Civil Worker');
    final iqamaCtrl = TextEditingController(text: worker?.iqamaNo ?? '');
    final mobileCtrl = TextEditingController(text: worker?.mobileNo ?? '');
    final companyCtrl = TextEditingController(text: worker?.company ?? 'Prime Infra');
    final sponsorCtrl = TextEditingController(text: worker?.sponsorName ?? 'Al-Rashid Contracting');
    final natCtrl = TextEditingController(text: worker?.nationality ?? 'India');
    final projectCtrl = TextEditingController(text: worker?.project ?? 'Metro Line 4');
    final campCtrl = TextEditingController(text: worker?.camp ?? 'East Valley Camp');
    WorkerStatus currentStatus = worker?.status ?? WorkerStatus.present;

    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: const Color(0xFF0F172A),
      shape: const RoundedRectangleBorder(
        borderRadius: BorderRadius.vertical(top: Radius.circular(20)),
      ),
      builder: (ctx) {
        return StatefulBuilder(
          builder: (context, setModalState) {
            return Padding(
              padding: EdgeInsets.only(
                top: 20,
                left: 18,
                right: 18,
                bottom: MediaQuery.of(ctx).viewInsets.bottom + 20,
              ),
              child: SingleChildScrollView(
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(
                      mainAxisAlignment: MainAxisAlignment.spaceBetween,
                      children: [
                        Text(
                          isEditing ? 'Edit Camp Record' : 'Add New Manpower Bed',
                          style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Colors.white),
                        ),
                        if (isEditing)
                          IconButton(
                            icon: const Icon(Icons.delete_forever_rounded, color: Color(0xFFEF4444)),
                            onPressed: () {
                              Navigator.pop(ctx);
                              _confirmDelete(worker);
                            },
                          ),
                      ],
                    ),
                    const SizedBox(height: 12),

                    // Status Dropdown selector
                    const Text('Status Assignment', style: TextStyle(fontSize: 12, color: Color(0xFF94A3B8))),
                    const SizedBox(height: 6),
                    Row(
                      children: WorkerStatus.values.map((st) {
                        final isSel = currentStatus == st;
                        return Expanded(
                          child: GestureDetector(
                            onTap: () => setModalState(() => currentStatus = st),
                            child: Container(
                              margin: const EdgeInsets.symmetric(horizontal: 2),
                              padding: const EdgeInsets.symmetric(vertical: 8),
                              decoration: BoxDecoration(
                                color: isSel ? st.color.withOpacity(0.2) : const Color(0xFF1E293B),
                                borderRadius: BorderRadius.circular(8),
                                border: Border.all(
                                  color: isSel ? st.color : const Color(0xFF334155),
                                ),
                              ),
                              child: Center(
                                child: Text(
                                  st.label,
                                  style: TextStyle(
                                    fontSize: 11,
                                    fontWeight: isSel ? FontWeight.bold : FontWeight.normal,
                                    color: isSel ? st.color : const Color(0xFF94A3B8),
                                  ),
                                ),
                              ),
                            ),
                          ),
                        );
                      }).toList(),
                    ),
                    const SizedBox(height: 14),

                    // Inputs: Name & FF ID
                    Row(
                      children: [
                        Expanded(
                          flex: 3,
                          child: _modalInput('Employee Name *', nameCtrl),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          flex: 2,
                          child: _modalInput('FF ID *', ffIdCtrl),
                        ),
                      ],
                    ),
                    const SizedBox(height: 10),

                    // Room & Position
                    Row(
                      children: [
                        Expanded(
                          child: _modalInput('Room No *', roomCtrl),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: _modalInput('Trade / Position', positionCtrl),
                        ),
                      ],
                    ),
                    const SizedBox(height: 10),

                    // Mobile & Iqama
                    Row(
                      children: [
                        Expanded(
                          child: _modalInput('Mobile No.', mobileCtrl, keyboardType: TextInputType.phone),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: _modalInput('Iqama No.', iqamaCtrl, keyboardType: TextInputType.number),
                        ),
                      ],
                    ),
                    const SizedBox(height: 10),

                    // Company & Nationality
                    Row(
                      children: [
                        Expanded(
                          child: _modalInput('Company', companyCtrl),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: _modalInput('Nationality', natCtrl),
                        ),
                      ],
                    ),
                    const SizedBox(height: 10),

                    // Project & Camp
                    Row(
                      children: [
                        Expanded(
                          child: _modalInput('Project', projectCtrl),
                        ),
                        const SizedBox(width: 10),
                        Expanded(
                          child: _modalInput('Camp Name', campCtrl),
                        ),
                      ],
                    ),
                    const SizedBox(height: 18),

                    // Save Button
                    SizedBox(
                      width: double.infinity,
                      child: ElevatedButton(
                        style: ElevatedButton.styleFrom(
                          backgroundColor: const Color(0xFF2563EB),
                          foregroundColor: Colors.white,
                          padding: const EdgeInsets.symmetric(vertical: 14),
                          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
                        ),
                        onPressed: () {
                          if (nameCtrl.text.trim().isEmpty || ffIdCtrl.text.trim().isEmpty) {
                            _showSnackBar('Name and FF ID are required fields', isError: true);
                            return;
                          }

                          setState(() {
                            if (isEditing) {
                              final idx = _roster.indexWhere((p) => p.ffId == worker.ffId);
                              if (idx != -1) {
                                _roster[idx] = worker.copyWith(
                                  employeeName: nameCtrl.text.trim(),
                                  ffId: ffIdCtrl.text.trim(),
                                  roomNo: roomCtrl.text.trim(),
                                  position: positionCtrl.text.trim(),
                                  iqamaNo: iqamaCtrl.text.trim(),
                                  mobileNo: mobileCtrl.text.trim(),
                                  company: companyCtrl.text.trim(),
                                  sponsorName: sponsorCtrl.text.trim(),
                                  nationality: natCtrl.text.trim(),
                                  project: projectCtrl.text.trim(),
                                  camp: campCtrl.text.trim(),
                                  status: currentStatus,
                                );
                              }
                            } else {
                              _roster.add(
                                CampPersonnel(
                                  sl: '${_roster.length + 1}',
                                  roomNo: roomCtrl.text.trim(),
                                  ffId: ffIdCtrl.text.trim(),
                                  position: positionCtrl.text.trim(),
                                  iqamaNo: iqamaCtrl.text.trim(),
                                  employeeName: nameCtrl.text.trim(),
                                  sponsorName: sponsorCtrl.text.trim(),
                                  nationality: natCtrl.text.trim(),
                                  mobileNo: mobileCtrl.text.trim(),
                                  religion: 'Muslim',
                                  company: companyCtrl.text.trim(),
                                  status: currentStatus,
                                  checkIn: DateTime.now().toString().split(' ')[0],
                                  vacant: 'Occupied',
                                  project: projectCtrl.text.trim(),
                                  camp: campCtrl.text.trim(),
                                ),
                              );
                            }
                            _hasUnsavedChanges = true;
                          });

                          Navigator.pop(ctx);
                          _showSnackBar(isEditing ? 'Record updated!' : 'Worker added to roster!');
                        },
                        child: Text(isEditing ? 'Save Changes' : 'Create Record'),
                      ),
                    ),
                  ],
                ),
              ),
            );
          },
        );
      },
    );
  }

  Widget _modalInput(String label, TextEditingController ctrl, {TextInputType? keyboardType}) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(label, style: const TextStyle(fontSize: 11, color: Color(0xFF94A3B8))),
        const SizedBox(height: 4),
        TextField(
          controller: ctrl,
          keyboardType: keyboardType,
          style: const TextStyle(fontSize: 13, color: Colors.white),
          decoration: InputDecoration(
            isDense: true,
            contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10),
            filled: true,
            fillColor: const Color(0xFF1E293B),
            border: OutlineInputBorder(
              borderRadius: BorderRadius.circular(8),
              borderSide: const BorderSide(color: Color(0xFF334155)),
            ),
          ),
        ),
      ],
    );
  }

  void _confirmDelete(CampPersonnel p) {
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        backgroundColor: const Color(0xFF0F172A),
        title: const Text('Delete Worker Record?'),
        content: Text('Are you sure you want to remove ${p.employeeName} (${p.ffId}) from Room ${p.roomNo}? This cannot be undone.'),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('Cancel')),
          ElevatedButton(
            style: ElevatedButton.styleFrom(backgroundColor: const Color(0xFFEF4444)),
            onPressed: () {
              setState(() {
                _roster.removeWhere((item) => item.ffId == p.ffId);
                _hasUnsavedChanges = true;
              });
              Navigator.pop(ctx);
              _showSnackBar('Removed ${p.employeeName} from roster');
            },
            child: const Text('Delete', style: TextStyle(color: Colors.white)),
          ),
        ],
      ),
    );
  }
}
