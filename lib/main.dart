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
    String?
