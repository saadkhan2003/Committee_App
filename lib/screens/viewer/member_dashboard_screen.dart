import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:intl/intl.dart';
import '../../services/database_service.dart';
import '../../models/committee.dart';
import '../../models/member.dart';
import '../../utils/app_theme.dart';

class MemberDashboardScreen extends StatefulWidget {
  final Committee committee;
  final Member member;

  const MemberDashboardScreen({
    super.key,
    required this.committee,
    required this.member,
  });

  @override
  State<MemberDashboardScreen> createState() => _MemberDashboardScreenState();
}

class _MemberDashboardScreenState extends State<MemberDashboardScreen> {
  final _dbService = DatabaseService();
  int _totalPayments = 0;
  int _paidPayments = 0;
  int _pendingPayments = 0;
  DateTime? _payoutDate;
  int _daysUntilPayout = 0;
  DateTime? _nextPaymentDate;

  @override
  void initState() {
    super.initState();
    _calculateStats();
  }

  void _calculateStats() {
    final payments = _dbService.getPaymentsByMember(
      widget.member.id,
    );

    final startDate = widget.committee.startDate;
    final today = DateTime.now();
    final members = _dbService.getMembersByCommittee(widget.committee.id);
    final totalCycles = members.length;
    final intervalDays = widget.committee.paymentIntervalDays;
    final endDate = startDate.add(Duration(days: intervalDays * totalCycles));
    
    int totalDays = 0;
    if (widget.committee.frequency == 'daily') {
      final actualEndDate = today.isBefore(endDate) ? today : endDate;
      totalDays = actualEndDate.difference(startDate).inDays + 1;
      if (totalDays < 0) totalDays = 0;
    } else if (widget.committee.frequency == 'weekly') {
      final actualEndDate = today.isBefore(endDate) ? today : endDate;
      totalDays = (actualEndDate.difference(startDate).inDays / 7).ceil();
    } else if (widget.committee.frequency == 'monthly') {
      final actualEndDate = today.isBefore(endDate) ? today : endDate;
      totalDays = (actualEndDate.difference(startDate).inDays / 30).ceil();
    }
    
    _totalPayments = totalDays > 0 ? totalDays : 0;
    _paidPayments = payments.where((p) => p.isPaid).length;
    _pendingPayments = _totalPayments - _paidPayments;
    if (_pendingPayments < 0) _pendingPayments = 0;

    if (widget.member.payoutOrder > 0) {
      // Payout happens at the END of the cycle, not beginning
      // So payoutOrder 1 gets payout after 1 interval (e.g., 1 month for monthly)
      _payoutDate = startDate.add(
        Duration(days: intervalDays * widget.member.payoutOrder),
      );
      _daysUntilPayout = _payoutDate!.difference(today).inDays;
      if (_daysUntilPayout < 0) _daysUntilPayout = 0;
    }

    if (widget.committee.frequency == 'daily') {
      _nextPaymentDate = DateTime(today.year, today.month, today.day + 1);
    } else if (widget.committee.frequency == 'weekly') {
      _nextPaymentDate = today.add(Duration(days: 7 - today.weekday % 7));
    } else {
      _nextPaymentDate = DateTime(today.year, today.month + 1, 1);
    }

    setState(() {});
  }

  @override
  Widget build(BuildContext context) {
    final progressPercent = _totalPayments > 0 
        ? (_paidPayments / _totalPayments * 100).clamp(0, 100) 
        : 0.0;

    return Scaffold(
      appBar: AppBar(
        title: const Text('My Dashboard'),
        centerTitle: true,
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(20),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            // Welcome Header
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(24),
              decoration: BoxDecoration(
                gradient: LinearGradient(
                  colors: [AppTheme.primaryColor, AppTheme.primaryDark],
                  begin: Alignment.topLeft,
                  end: Alignment.bottomRight,
                ),
                borderRadius: BorderRadius.circular(20),
                boxShadow: [
                  BoxShadow(
                    color: AppTheme.primaryColor.withOpacity(0.3),
                    blurRadius: 20,
                    offset: const Offset(0, 10),
                  ),
                ],
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'Welcome, ${widget.member.name}!',
                    style: GoogleFonts.inter(
                      fontSize: 24,
                      fontWeight: FontWeight.bold,
                      color: Colors.white,
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    widget.committee.name,
                    style: GoogleFonts.inter(
                      fontSize: 16,
                      color: Colors.white70,
                    ),
                  ),
                ],
              ),
            ),
            const SizedBox(height: 24),

            // Payout Card
            _buildInfoCard(
              icon: Icons.account_balance_wallet_rounded,
              title: 'Your Payout',
              content: widget.member.hasReceivedPayout
                  ? 'Already Received!'
                  : _payoutDate != null
                      ? DateFormat('MMMM d, yyyy').format(_payoutDate!)
                      : 'Not assigned yet',
              subtitle: widget.member.hasReceivedPayout
                  ? 'Congratulations!'
                  : _daysUntilPayout > 0
                      ? '$_daysUntilPayout days remaining'
                      : _payoutDate != null
                          ? 'Payout day!'
                          : 'Waiting for assignment',
              color: widget.member.hasReceivedPayout
                  ? AppTheme.secondaryColor
                  : AppTheme.warningColor,
            ),
            const SizedBox(height: 16),

            // Payment Status Card
            Container(
              padding: const EdgeInsets.all(20),
              decoration: BoxDecoration(
                color: AppTheme.darkCard,
                borderRadius: BorderRadius.circular(16),
                border: Border.all(color: Colors.grey[800]!),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Row(
                    children: [
                      Icon(Icons.bar_chart_rounded, color: Colors.white, size: 24),
                      const SizedBox(width: 12),
                      Text(
                        'Your Payment Status',
                        style: GoogleFonts.inter(
                          fontSize: 18,
                          fontWeight: FontWeight.bold,
                          color: Colors.white,
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 20),
                  Row(
                    children: [
                      Expanded(
                        child: _buildStatItem(
                          'Paid',
                          '$_paidPayments/${_totalPayments > 0 ? _totalPayments : "-"}',
                          AppTheme.secondaryColor,
                        ),
                      ),
                      const SizedBox(width: 16),
                      Expanded(
                        child: _buildStatItem(
                          'Pending',
                          '$_pendingPayments days',
                          AppTheme.errorColor,
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 20),
                  Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        mainAxisAlignment: MainAxisAlignment.spaceBetween,
                        children: [
                          Text(
                            'Progress',
                            style: TextStyle(color: Colors.grey[400], fontSize: 13),
                          ),
                          Text(
                            '${progressPercent.toStringAsFixed(0)}%',
                            style: GoogleFonts.inter(
                              fontWeight: FontWeight.bold,
                              color: AppTheme.secondaryColor,
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 8),
                      ClipRRect(
                        borderRadius: BorderRadius.circular(10),
                        child: LinearProgressIndicator(
                          value: progressPercent / 100,
                          minHeight: 12,
                          backgroundColor: Colors.grey[800],
                          valueColor: AlwaysStoppedAnimation<Color>(
                            progressPercent >= 80
                                ? AppTheme.secondaryColor
                                : progressPercent >= 50
                                    ? AppTheme.warningColor
                                    : AppTheme.primaryColor,
                          ),
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
            const SizedBox(height: 16),

            // Next Payment Card
            _buildInfoCard(
              icon: Icons.calendar_today_rounded,
              title: 'Next Payment',
              content: _nextPaymentDate != null
                  ? _isToday(_nextPaymentDate!)
                      ? 'Today'
                      : _isTomorrow(_nextPaymentDate!)
                          ? 'Tomorrow'
                          : DateFormat('EEEE, MMM d').format(_nextPaymentDate!)
                  : 'N/A',
              subtitle: 'Amount: PKR ${NumberFormat('#,###').format(widget.committee.contributionAmount.toInt())}',
              color: AppTheme.primaryColor,
            ),
            const SizedBox(height: 16),

            // Committee Info Card
            Container(
              padding: const EdgeInsets.all(20),
              decoration: BoxDecoration(
                color: AppTheme.darkCard,
                borderRadius: BorderRadius.circular(16),
                border: Border.all(color: Colors.grey[800]!),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'Committee Details',
                    style: GoogleFonts.inter(
                      fontSize: 16,
                      fontWeight: FontWeight.bold,
                      color: Colors.white,
                    ),
                  ),
                  const SizedBox(height: 16),
                  _buildDetailRow('Committee Code', widget.committee.code),
                  _buildDetailRow('Your Member Code', widget.member.memberCode),
                  _buildDetailRow('Your Payout Order', '#${widget.member.payoutOrder}'),
                  _buildDetailRow('Collection', _capitalize(widget.committee.frequency)),
                  _buildDetailRow(
                    'Contribution',
                    'PKR ${NumberFormat('#,###').format(widget.committee.contributionAmount.toInt())}',
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildInfoCard({
    required IconData icon,
    required String title,
    required String content,
    required String subtitle,
    required Color color,
  }) {
    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        color: AppTheme.darkCard,
        borderRadius: BorderRadius.circular(16),
        border: Border.all(color: color.withOpacity(0.3)),
      ),
      child: Row(
        children: [
          Container(
            padding: const EdgeInsets.all(12),
            decoration: BoxDecoration(
              color: color.withOpacity(0.15),
              borderRadius: BorderRadius.circular(12),
            ),
            child: Icon(icon, color: color, size: 28),
          ),
          const SizedBox(width: 16),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  title,
                  style: TextStyle(color: Colors.grey[400], fontSize: 13),
                ),
                const SizedBox(height: 4),
                Text(
                  content,
                  style: GoogleFonts.inter(
                    fontSize: 20,
                    fontWeight: FontWeight.bold,
                    color: Colors.white,
                  ),
                ),
                const SizedBox(height: 4),
                Text(
                  subtitle,
                  style: TextStyle(color: color, fontSize: 14),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildStatItem(String label, String value, Color color) {
    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: color.withOpacity(0.1),
        borderRadius: BorderRadius.circular(12),
      ),
      child: Column(
        children: [
          Text(
            value,
            style: GoogleFonts.inter(
              fontSize: 24,
              fontWeight: FontWeight.bold,
              color: color,
            ),
          ),
          const SizedBox(height: 4),
          Text(
            label,
            style: TextStyle(color: Colors.grey[400], fontSize: 13),
          ),
        ],
      ),
    );
  }

  Widget _buildDetailRow(String label, String value) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 8),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: TextStyle(color: Colors.grey[500])),
          Text(
            value,
            style: GoogleFonts.inter(
              fontWeight: FontWeight.w600,
              color: Colors.white,
            ),
          ),
        ],
      ),
    );
  }

  bool _isToday(DateTime date) {
    final now = DateTime.now();
    return date.year == now.year && date.month == now.month && date.day == now.day;
  }

  bool _isTomorrow(DateTime date) {
    final tomorrow = DateTime.now().add(const Duration(days: 1));
    return date.year == tomorrow.year && date.month == tomorrow.month && date.day == tomorrow.day;
  }

  String _capitalize(String s) {
    if (s.isEmpty) return s;
    return '${s[0].toUpperCase()}${s.substring(1)}';
  }
}
