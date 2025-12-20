import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:printing/printing.dart';
import 'package:csv/csv.dart';
import 'package:intl/intl.dart';
import 'package:share_plus/share_plus.dart';
import 'package:path_provider/path_provider.dart';
import '../models/committee.dart';
import '../models/member.dart';
import '../models/payment.dart';
import 'database_service.dart';

class ExportService {
  final DatabaseService _dbService = DatabaseService();

  // Theme colors
  static const _primaryColor = PdfColor.fromInt(0xFF1A1A2E);
  static const _accentColor = PdfColor.fromInt(0xFF00C853);
  static const _lightGreen = PdfColor.fromInt(0xFFE8F5E9);
  static const _lightRed = PdfColor.fromInt(0xFFFFEBEE);
  static const _lightBg = PdfColor.fromInt(0xFFF8F9FA);

  // ============ PDF EXPORT ============

  Future<void> exportToPdf(
    Committee committee, {
    DateTime? startDate,
    DateTime? endDate,
  }) async {
    final members = _dbService.getMembersByCommittee(committee.id);
    members.sort((a, b) => a.payoutOrder.compareTo(b.payoutOrder));

    final payments = _dbService.getPaymentsByCommittee(committee.id);
    final dates = _generateDates(committee, startDate: startDate, endDate: endDate);

    // Calculate totals
    final totalPayments = payments.where((p) => p.isPaid).length;
    final totalCollected = totalPayments * committee.contributionAmount;
    final paidMembersCount = members.where((m) => m.hasReceivedPayout).length;
    final totalExpectedPayments = members.length * dates.length;
    final collectionRate = totalExpectedPayments > 0 
        ? (totalPayments / totalExpectedPayments * 100).toStringAsFixed(1) 
        : '0';

    final pdf = pw.Document();

    pdf.addPage(
      pw.MultiPage(
        pageFormat: PdfPageFormat.a4,
        margin: const pw.EdgeInsets.all(32),
        header: (context) => _buildHeader(committee, context),
        footer: (context) => _buildFooter(context),
        build: (context) => [
          // Summary Section
          pw.Container(
            padding: const pw.EdgeInsets.all(20),
            decoration: pw.BoxDecoration(
              color: _lightBg,
              borderRadius: pw.BorderRadius.circular(8),
              border: pw.Border.all(color: PdfColors.grey300),
            ),
            child: pw.Column(
              crossAxisAlignment: pw.CrossAxisAlignment.start,
              children: [
                pw.Text('Summary', style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
                pw.SizedBox(height: 12),
                pw.Row(
                  mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
                  children: [
                    _statBox('Contribution', 'Rs. ${committee.contributionAmount.toInt()}'),
                    _statBox('Frequency', committee.frequency.toUpperCase()),
                    _statBox('Members', '${members.length}'),
                    _statBox('Collection Rate', '$collectionRate%'),
                  ],
                ),
                pw.SizedBox(height: 12),
                pw.Divider(color: PdfColors.grey300),
                pw.SizedBox(height: 12),
                pw.Row(
                  mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
                  children: [
                    _statBox('Total Collected', 'Rs. ${totalCollected.toInt()}', highlight: true),
                    _statBox('Payouts Completed', '$paidMembersCount / ${members.length}', highlight: true),
                    _statBox('Period', '${dates.length} cycles'),
                    _statBox('Pending', 'Rs. ${((totalExpectedPayments - totalPayments) * committee.contributionAmount).toInt()}'),
                  ],
                ),
              ],
            ),
          ),
          pw.SizedBox(height: 24),

          // Members Table
          pw.Text('Member Details', style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
          pw.SizedBox(height: 12),
          pw.Table(
            border: pw.TableBorder.all(color: PdfColors.grey400),
            columnWidths: {
              0: const pw.FlexColumnWidth(0.8),  // #
              1: const pw.FlexColumnWidth(3),    // Name
              2: const pw.FlexColumnWidth(2),    // Phone
              3: const pw.FlexColumnWidth(1.5),  // Paid
              4: const pw.FlexColumnWidth(1.2),  // %
              5: const pw.FlexColumnWidth(2),    // Amount
              6: const pw.FlexColumnWidth(1.5),  // Payout
            },
            children: [
              // Header
              pw.TableRow(
                decoration: const pw.BoxDecoration(color: _primaryColor),
                children: [
                  _tableHeader('#'),
                  _tableHeader('Member Name'),
                  _tableHeader('Phone'),
                  _tableHeader('Payments'),
                  _tableHeader('%'),
                  _tableHeader('Total Paid'),
                  _tableHeader('Payout'),
                ],
              ),
              // Data rows
              ...members.asMap().entries.map((entry) {
                final index = entry.key;
                final m = entry.value;
                int paidCount = 0;
                for (var date in dates) {
                  if (_isPaymentMarked(payments, m.id, date)) paidCount++;
                }
                final memberTotal = paidCount * committee.contributionAmount;
                final percentage = dates.isNotEmpty ? (paidCount / dates.length * 100).toInt() : 0;

                return pw.TableRow(
                  decoration: pw.BoxDecoration(
                    color: index % 2 == 0 ? PdfColors.white : _lightBg,
                  ),
                  children: [
                    _tableCell('${m.payoutOrder}', center: true),
                    _tableCell(m.name, bold: true),
                    _tableCell(m.phone),
                    _tableCell('$paidCount / ${dates.length}', center: true),
                    _tableCell('$percentage%', center: true, bgColor: percentage >= 80 ? _lightGreen : (percentage < 50 ? _lightRed : null)),
                    _tableCell('Rs. ${memberTotal.toInt()}'),
                    _tableCell(
                      m.hasReceivedPayout ? 'DONE' : 'Pending',
                      center: true,
                      bgColor: m.hasReceivedPayout ? _lightGreen : null,
                      bold: m.hasReceivedPayout,
                    ),
                  ],
                );
              }),
              // Totals row
              pw.TableRow(
                decoration: const pw.BoxDecoration(color: _accentColor),
                children: [
                  _tableFooter(''),
                  _tableFooter('TOTAL'),
                  _tableFooter(''),
                  _tableFooter('$totalPayments / $totalExpectedPayments'),
                  _tableFooter('$collectionRate%'),
                  _tableFooter('Rs. ${totalCollected.toInt()}'),
                  _tableFooter('$paidMembersCount done'),
                ],
              ),
            ],
          ),
          pw.SizedBox(height: 24),

          // Payout Schedule
          pw.Text('Payout Schedule', style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
          pw.SizedBox(height: 12),
          pw.Wrap(
            spacing: 8,
            runSpacing: 8,
            children: members.map((m) {
              final isDone = m.hasReceivedPayout;
              return pw.Container(
                width: 160,
                padding: const pw.EdgeInsets.all(10),
                decoration: pw.BoxDecoration(
                  color: isDone ? _lightGreen : PdfColors.white,
                  border: pw.Border.all(color: isDone ? PdfColors.green400 : PdfColors.grey300),
                  borderRadius: pw.BorderRadius.circular(6),
                ),
                child: pw.Column(
                  crossAxisAlignment: pw.CrossAxisAlignment.start,
                  children: [
                    pw.Row(
                      mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
                      children: [
                        pw.Text('#${m.payoutOrder}', style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold)),
                        pw.Container(
                          padding: const pw.EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                          decoration: pw.BoxDecoration(
                            color: isDone ? PdfColors.green : PdfColors.grey400,
                            borderRadius: pw.BorderRadius.circular(4),
                          ),
                          child: pw.Text(isDone ? 'DONE' : 'PENDING', style: pw.TextStyle(fontSize: 7, color: PdfColors.white, fontWeight: pw.FontWeight.bold)),
                        ),
                      ],
                    ),
                    pw.SizedBox(height: 4),
                    pw.Text(m.name, style: pw.TextStyle(fontSize: 11, fontWeight: pw.FontWeight.bold)),
                    if (m.payoutDate != null)
                      pw.Text('Received: ${DateFormat('dd/MM/yy').format(m.payoutDate!)}', style: const pw.TextStyle(fontSize: 8, color: PdfColors.grey600)),
                  ],
                ),
              );
            }).toList(),
          ),
        ],
      ),
    );

    await Printing.layoutPdf(
      onLayout: (PdfPageFormat format) async => pdf.save(),
      name: '${committee.name}_report.pdf',
    );
  }

  pw.Widget _buildHeader(Committee committee, pw.Context context) {
    return pw.Container(
      margin: const pw.EdgeInsets.only(bottom: 20),
      child: pw.Row(
        mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
        crossAxisAlignment: pw.CrossAxisAlignment.start,
        children: [
          pw.Column(
            crossAxisAlignment: pw.CrossAxisAlignment.start,
            children: [
              pw.Text(committee.name, style: pw.TextStyle(fontSize: 24, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
              pw.SizedBox(height: 4),
              pw.Text('Payment Report', style: const pw.TextStyle(fontSize: 12, color: PdfColors.grey600)),
            ],
          ),
          pw.Container(
            padding: const pw.EdgeInsets.symmetric(horizontal: 16, vertical: 10),
            decoration: pw.BoxDecoration(
              color: _accentColor,
              borderRadius: pw.BorderRadius.circular(8),
            ),
            child: pw.Column(
              children: [
                pw.Text('Code', style: const pw.TextStyle(fontSize: 9, color: PdfColors.white)),
                pw.Text(committee.code, style: pw.TextStyle(fontSize: 18, fontWeight: pw.FontWeight.bold, color: PdfColors.white)),
              ],
            ),
          ),
        ],
      ),
    );
  }

  pw.Widget _buildFooter(pw.Context context) {
    return pw.Container(
      margin: const pw.EdgeInsets.only(top: 16),
      padding: const pw.EdgeInsets.only(top: 8),
      decoration: const pw.BoxDecoration(border: pw.Border(top: pw.BorderSide(color: PdfColors.grey300))),
      child: pw.Row(
        mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
        children: [
          pw.Text('Generated: ${DateFormat('dd MMM yyyy, hh:mm a').format(DateTime.now())}', style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey600)),
          pw.Text('Page ${context.pageNumber} of ${context.pagesCount}', style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey600)),
        ],
      ),
    );
  }

  pw.Widget _statBox(String label, String value, {bool highlight = false}) {
    return pw.Column(
      crossAxisAlignment: pw.CrossAxisAlignment.start,
      children: [
        pw.Text(label, style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey600)),
        pw.SizedBox(height: 2),
        pw.Text(value, style: pw.TextStyle(fontSize: 14, fontWeight: pw.FontWeight.bold, color: highlight ? _accentColor : _primaryColor)),
      ],
    );
  }

  pw.Widget _tableHeader(String text) {
    return pw.Container(
      padding: const pw.EdgeInsets.symmetric(horizontal: 8, vertical: 10),
      child: pw.Text(text, style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold, color: PdfColors.white), textAlign: pw.TextAlign.center),
    );
  }

  pw.Widget _tableCell(String text, {bool bold = false, bool center = false, PdfColor? bgColor}) {
    return pw.Container(
      padding: const pw.EdgeInsets.symmetric(horizontal: 8, vertical: 8),
      color: bgColor,
      child: pw.Text(
        text,
        style: pw.TextStyle(fontSize: 10, fontWeight: bold ? pw.FontWeight.bold : pw.FontWeight.normal),
        textAlign: center ? pw.TextAlign.center : pw.TextAlign.left,
      ),
    );
  }

  pw.Widget _tableFooter(String text) {
    return pw.Container(
      padding: const pw.EdgeInsets.symmetric(horizontal: 8, vertical: 10),
      child: pw.Text(text, style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold, color: PdfColors.white), textAlign: pw.TextAlign.center),
    );
  }

  // ============ CSV EXPORT ============

  Future<void> exportToCsv(
    Committee committee, {
    DateTime? startDate,
    DateTime? endDate,
  }) async {
    final members = _dbService.getMembersByCommittee(committee.id);
    members.sort((a, b) => a.payoutOrder.compareTo(b.payoutOrder));

    final payments = _dbService.getPaymentsByCommittee(committee.id);
    final dates = _generateDates(committee, startDate: startDate, endDate: endDate);

    List<List<dynamic>> rows = [];

    // Title
    rows.add(['${committee.name} - Payment Report']);
    rows.add(['Code: ${committee.code}']);
    rows.add(['Generated: ${DateFormat('dd MMMM yyyy, hh:mm a').format(DateTime.now())}']);
    rows.add([]);

    // Summary
    rows.add(['SUMMARY']);
    rows.add(['Contribution', 'Rs. ${committee.contributionAmount.toInt()}']);
    rows.add(['Frequency', committee.frequency.toUpperCase()]);
    rows.add(['Members', members.length]);
    rows.add(['Total Collected', 'Rs. ${(payments.where((p) => p.isPaid).length * committee.contributionAmount).toInt()}']);
    rows.add([]);

    // Member Details
    rows.add(['MEMBER DETAILS']);
    rows.add(['Order', 'Name', 'Phone', 'Payments Made', 'Percentage', 'Total Paid', 'Payout Status', 'Payout Date']);
    for (var m in members) {
      int paidCount = 0;
      for (var date in dates) {
        if (_isPaymentMarked(payments, m.id, date)) paidCount++;
      }
      final percentage = dates.isNotEmpty ? (paidCount / dates.length * 100).toInt() : 0;
      rows.add([
        m.payoutOrder,
        m.name,
        m.phone,
        '$paidCount / ${dates.length}',
        '$percentage%',
        'Rs. ${(paidCount * committee.contributionAmount).toInt()}',
        m.hasReceivedPayout ? 'DONE' : 'Pending',
        m.payoutDate != null ? DateFormat('dd/MM/yyyy').format(m.payoutDate!) : '-',
      ]);
    }

    String csv = const ListToCsvConverter().convert(rows);

    if (kIsWeb) {
      await Printing.sharePdf(
        bytes: Uint8List.fromList(csv.codeUnits),
        filename: '${committee.name}_report.csv',
      );
    } else {
      final directory = await getApplicationDocumentsDirectory();
      final file = File('${directory.path}/${committee.name}_report.csv');
      await file.writeAsString(csv);
      await Share.shareXFiles([XFile(file.path)], subject: '${committee.name} Report');
    }
  }

  // ============ HELPERS ============

  List<DateTime> _generateDates(Committee committee, {DateTime? startDate, DateTime? endDate}) {
    List<DateTime> dates = [];
    final now = DateTime.now();
    final start = startDate ?? committee.startDate;
    final end = endDate ?? now;

    DateTime current = DateTime(start.year, start.month, start.day);

    while (!current.isAfter(end)) {
      dates.add(current);
      if (committee.frequency == 'monthly') {
        current = DateTime(current.year, current.month + 1, current.day);
      } else if (committee.frequency == 'weekly') {
        current = current.add(const Duration(days: 7));
      } else {
        current = current.add(const Duration(days: 1));
      }
    }
    return dates;
  }

  bool _isPaymentMarked(List<Payment> payments, String memberId, DateTime date) {
    try {
      final payment = payments.firstWhere(
        (p) => p.memberId == memberId && p.date.year == date.year && p.date.month == date.month && p.date.day == date.day,
      );
      return payment.isPaid;
    } catch (e) {
      return false;
    }
  }
}
