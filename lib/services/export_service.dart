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

  static const _primaryColor = PdfColor.fromInt(0xFF1A1A2E);
  static const _accentColor = PdfColor.fromInt(0xFF00C853);
  static const _lightGreen = PdfColor.fromInt(0xFFE8F5E9);
  static const _lightRed = PdfColor.fromInt(0xFFFFEBEE);
  static const _lightBg = PdfColor.fromInt(0xFFF8F9FA);

  // ============ PDF EXPORT ============

  Future<void> exportToPdf(Committee committee, {DateTime? startDate, DateTime? endDate}) async {
    final members = _dbService.getMembersByCommittee(committee.id);
    members.sort((a, b) => a.payoutOrder.compareTo(b.payoutOrder));
    final payments = _dbService.getPaymentsByCommittee(committee.id);
    final dates = _generateDates(committee, startDate: startDate, endDate: endDate);

    final totalPayments = payments.where((p) => p.isPaid).length;
    final totalCollected = totalPayments * committee.contributionAmount;
    final paidMembersCount = members.where((m) => m.hasReceivedPayout).length;
    final totalExpected = members.length * dates.length;
    final collectionRate = totalExpected > 0 ? (totalPayments / totalExpected * 100).toStringAsFixed(1) : '0';

    final pdf = pw.Document();
    pdf.addPage(
      pw.MultiPage(
        pageFormat: PdfPageFormat.a4,
        margin: const pw.EdgeInsets.all(32),
        header: (ctx) => _buildHeader(committee),
        footer: (ctx) => _buildFooter(ctx),
        build: (ctx) => [
          pw.Container(
            padding: const pw.EdgeInsets.all(20),
            decoration: pw.BoxDecoration(color: _lightBg, borderRadius: pw.BorderRadius.circular(8), border: pw.Border.all(color: PdfColors.grey300)),
            child: pw.Column(crossAxisAlignment: pw.CrossAxisAlignment.start, children: [
              pw.Text('Summary', style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
              pw.SizedBox(height: 12),
              pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
                _statBox('Contribution', 'Rs. ${committee.contributionAmount.toInt()}'),
                _statBox('Frequency', committee.frequency.toUpperCase()),
                _statBox('Members', '${members.length}'),
                _statBox('Collection', '$collectionRate%'),
              ]),
              pw.Divider(color: PdfColors.grey300),
              pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
                _statBox('Collected', 'Rs. ${totalCollected.toInt()}', highlight: true),
                _statBox('Payouts', '$paidMembersCount/${members.length}', highlight: true),
                _statBox('Cycles', '${dates.length}'),
                _statBox('Pending', 'Rs. ${((totalExpected - totalPayments) * committee.contributionAmount).toInt()}'),
              ]),
            ]),
          ),
          pw.SizedBox(height: 24),
          pw.Text('Member Details', style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
          pw.SizedBox(height: 12),
          pw.Table(
            border: pw.TableBorder.all(color: PdfColors.grey400),
            columnWidths: {0: const pw.FlexColumnWidth(0.8), 1: const pw.FlexColumnWidth(3), 2: const pw.FlexColumnWidth(2), 3: const pw.FlexColumnWidth(1.5), 4: const pw.FlexColumnWidth(1.2), 5: const pw.FlexColumnWidth(2), 6: const pw.FlexColumnWidth(1.5)},
            children: [
              pw.TableRow(decoration: const pw.BoxDecoration(color: _primaryColor), children: [_th('#'), _th('Name'), _th('Phone'), _th('Paid'), _th('%'), _th('Amount'), _th('Payout')]),
              ...members.asMap().entries.map((e) {
                int paid = 0;
                for (var d in dates) if (_isPaymentMarked(payments, e.value.id, d)) paid++;
                final pct = dates.isNotEmpty ? (paid / dates.length * 100).toInt() : 0;
                return pw.TableRow(decoration: pw.BoxDecoration(color: e.key % 2 == 0 ? PdfColors.white : _lightBg), children: [
                  _td('${e.value.payoutOrder}', center: true),
                  _td(e.value.name, bold: true),
                  _td(e.value.phone),
                  _td('$paid/${dates.length}', center: true),
                  _td('$pct%', center: true, bg: pct >= 80 ? _lightGreen : (pct < 50 ? _lightRed : null)),
                  _td('Rs. ${(paid * committee.contributionAmount).toInt()}'),
                  _td(e.value.hasReceivedPayout ? 'DONE' : 'Pending', center: true, bg: e.value.hasReceivedPayout ? _lightGreen : null, bold: e.value.hasReceivedPayout),
                ]);
              }),
              pw.TableRow(decoration: const pw.BoxDecoration(color: _accentColor), children: [_tf(''), _tf('TOTAL'), _tf(''), _tf('$totalPayments/$totalExpected'), _tf('$collectionRate%'), _tf('Rs. ${totalCollected.toInt()}'), _tf('$paidMembersCount')]),
            ],
          ),
          pw.SizedBox(height: 24),
          pw.Text('Payout Schedule', style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold, color: _primaryColor)),
          pw.SizedBox(height: 12),
          pw.Wrap(spacing: 8, runSpacing: 8, children: members.map((m) => pw.Container(
            width: 160, padding: const pw.EdgeInsets.all(10),
            decoration: pw.BoxDecoration(color: m.hasReceivedPayout ? _lightGreen : PdfColors.white, border: pw.Border.all(color: m.hasReceivedPayout ? PdfColors.green400 : PdfColors.grey300), borderRadius: pw.BorderRadius.circular(6)),
            child: pw.Column(crossAxisAlignment: pw.CrossAxisAlignment.start, children: [
              pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
                pw.Text('#${m.payoutOrder}', style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold)),
                pw.Container(padding: const pw.EdgeInsets.symmetric(horizontal: 6, vertical: 2), decoration: pw.BoxDecoration(color: m.hasReceivedPayout ? PdfColors.green : PdfColors.grey400, borderRadius: pw.BorderRadius.circular(4)), child: pw.Text(m.hasReceivedPayout ? 'DONE' : 'PENDING', style: pw.TextStyle(fontSize: 7, color: PdfColors.white, fontWeight: pw.FontWeight.bold))),
              ]),
              pw.SizedBox(height: 4),
              pw.Text(m.name, style: pw.TextStyle(fontSize: 11, fontWeight: pw.FontWeight.bold)),
              if (m.payoutDate != null) pw.Text('Received: ${DateFormat('dd/MM/yy').format(m.payoutDate!)}', style: const pw.TextStyle(fontSize: 8, color: PdfColors.grey600)),
            ]),
          )).toList()),
        ],
      ),
    );
    await Printing.layoutPdf(onLayout: (f) async => pdf.save(), name: '${committee.name}_report.pdf');
  }

  pw.Widget _buildHeader(Committee c) => pw.Container(margin: const pw.EdgeInsets.only(bottom: 20), child: pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
    pw.Column(crossAxisAlignment: pw.CrossAxisAlignment.start, children: [pw.Text(c.name, style: pw.TextStyle(fontSize: 24, fontWeight: pw.FontWeight.bold, color: _primaryColor)), pw.Text('Payment Report', style: const pw.TextStyle(fontSize: 12, color: PdfColors.grey600))]),
    pw.Container(padding: const pw.EdgeInsets.symmetric(horizontal: 16, vertical: 10), decoration: pw.BoxDecoration(color: _accentColor, borderRadius: pw.BorderRadius.circular(8)), child: pw.Column(children: [pw.Text('Code', style: const pw.TextStyle(fontSize: 9, color: PdfColors.white)), pw.Text(c.code, style: pw.TextStyle(fontSize: 18, fontWeight: pw.FontWeight.bold, color: PdfColors.white))])),
  ]));

  pw.Widget _buildFooter(pw.Context ctx) => pw.Container(margin: const pw.EdgeInsets.only(top: 16), padding: const pw.EdgeInsets.only(top: 8), decoration: const pw.BoxDecoration(border: pw.Border(top: pw.BorderSide(color: PdfColors.grey300))), child: pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [pw.Text('Generated: ${DateFormat('dd MMM yyyy, hh:mm a').format(DateTime.now())}', style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey600)), pw.Text('Page ${ctx.pageNumber}/${ctx.pagesCount}', style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey600))]));

  pw.Widget _statBox(String l, String v, {bool highlight = false}) => pw.Column(crossAxisAlignment: pw.CrossAxisAlignment.start, children: [pw.Text(l, style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey600)), pw.Text(v, style: pw.TextStyle(fontSize: 14, fontWeight: pw.FontWeight.bold, color: highlight ? _accentColor : _primaryColor))]);
  pw.Widget _th(String t) => pw.Container(padding: const pw.EdgeInsets.all(8), child: pw.Text(t, style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold, color: PdfColors.white), textAlign: pw.TextAlign.center));
  pw.Widget _td(String t, {bool bold = false, bool center = false, PdfColor? bg}) => pw.Container(padding: const pw.EdgeInsets.all(6), color: bg, child: pw.Text(t, style: pw.TextStyle(fontSize: 10, fontWeight: bold ? pw.FontWeight.bold : pw.FontWeight.normal), textAlign: center ? pw.TextAlign.center : pw.TextAlign.left));
  pw.Widget _tf(String t) => pw.Container(padding: const pw.EdgeInsets.all(8), child: pw.Text(t, style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold, color: PdfColors.white), textAlign: pw.TextAlign.center));

  // ============ CSV EXPORT ============

  Future<void> exportToCsv(Committee committee, {DateTime? startDate, DateTime? endDate}) async {
    final members = _dbService.getMembersByCommittee(committee.id);
    members.sort((a, b) => a.payoutOrder.compareTo(b.payoutOrder));
    final payments = _dbService.getPaymentsByCommittee(committee.id);
    final dates = _generateDates(committee, startDate: startDate, endDate: endDate);

    final totalPayments = payments.where((p) => p.isPaid).length;
    final totalExpected = members.length * dates.length;
    final totalCollected = totalPayments * committee.contributionAmount;
    final totalPending = (totalExpected - totalPayments) * committee.contributionAmount;
    final collectionRate = totalExpected > 0 ? (totalPayments / totalExpected * 100).toStringAsFixed(1) : '0';
    final paidCount = members.where((m) => m.hasReceivedPayout).length;

    List<List<dynamic>> rows = [];

    rows.add(['=========================================================']);
    rows.add(['${committee.name.toUpperCase()} - PAYMENT REPORT']);
    rows.add(['=========================================================']);
    rows.add(['Committee Code', committee.code]);
    rows.add(['Generated', DateFormat('dd MMMM yyyy, hh:mm a').format(DateTime.now())]);
    rows.add([]);
    rows.add(['------------------ SUMMARY ------------------']);
    rows.add(['Field', 'Value']);
    rows.add(['Contribution', 'Rs. ${committee.contributionAmount.toInt()}']);
    rows.add(['Frequency', committee.frequency.toUpperCase()]);
    rows.add(['Members', members.length]);
    rows.add(['Cycles', dates.length]);
    rows.add(['Collected', 'Rs. ${totalCollected.toInt()}']);
    rows.add(['Pending', 'Rs. ${totalPending.toInt()}']);
    rows.add(['Collection Rate', '$collectionRate%']);
    rows.add(['Payouts Done', '$paidCount / ${members.length}']);
    rows.add([]);
    rows.add(['------------------ PAYOUT SCHEDULE ------------------']);
    rows.add(['Order', 'Member', 'Phone', 'Code', 'Status', 'Date']);
    for (var m in members) rows.add([m.payoutOrder, m.name, m.phone, m.memberCode, m.hasReceivedPayout ? 'RECEIVED' : 'Pending', m.payoutDate != null ? DateFormat('dd/MM/yyyy').format(m.payoutDate!) : '-']);
    rows.add([]);
    rows.add(['------------------ MEMBER PAYMENTS ------------------']);
    rows.add(['Order', 'Member', 'Paid', 'Missed', '%', 'Amount Paid', 'Amount Due']);
    for (var m in members) {
      int paid = 0;
      for (var d in dates) if (_isPaymentMarked(payments, m.id, d)) paid++;
      final missed = dates.length - paid;
      final pct = dates.isNotEmpty ? (paid / dates.length * 100).toInt() : 0;
      rows.add([m.payoutOrder, m.name, paid, missed, '$pct%', 'Rs. ${(paid * committee.contributionAmount).toInt()}', 'Rs. ${(missed * committee.contributionAmount).toInt()}']);
    }
    rows.add([]);
    rows.add(['', 'TOTAL', totalPayments, totalExpected - totalPayments, '$collectionRate%', 'Rs. ${totalCollected.toInt()}', 'Rs. ${totalPending.toInt()}']);
    rows.add([]);
    rows.add(['=========================================================']);
    rows.add(['END OF REPORT']);

    String csv = const ListToCsvConverter().convert(rows);
    if (kIsWeb) {
      await Printing.sharePdf(bytes: Uint8List.fromList(csv.codeUnits), filename: '${committee.name}_report.csv');
    } else {
      final dir = await getApplicationDocumentsDirectory();
      final file = File('${dir.path}/${committee.name}_report.csv');
      await file.writeAsString(csv);
      await Share.shareXFiles([XFile(file.path)], subject: '${committee.name} Report');
    }
  }

  List<DateTime> _generateDates(Committee c, {DateTime? startDate, DateTime? endDate}) {
    List<DateTime> dates = [];
    DateTime cur = DateTime((startDate ?? c.startDate).year, (startDate ?? c.startDate).month, (startDate ?? c.startDate).day);
    final end = endDate ?? DateTime.now();
    while (!cur.isAfter(end)) {
      dates.add(cur);
      cur = c.frequency == 'monthly' ? DateTime(cur.year, cur.month + 1, cur.day) : c.frequency == 'weekly' ? cur.add(const Duration(days: 7)) : cur.add(const Duration(days: 1));
    }
    return dates;
  }

  bool _isPaymentMarked(List<Payment> payments, String memberId, DateTime date) {
    try { return payments.firstWhere((p) => p.memberId == memberId && p.date.year == date.year && p.date.month == date.month && p.date.day == date.day).isPaid; }
    catch (e) { return false; }
  }
}
