import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:intl/intl.dart';
import 'package:flutter/foundation.dart' show kIsWeb;

// Project imports
import '../../services/auth_service.dart';
import '../../services/database_service.dart';
import '../../services/export_service.dart';
import '../../services/auto_sync_service.dart';
import '../../services/sync_service.dart';
import '../../services/update_service.dart';
import '../../services/analytics_service.dart';
import '../../services/toast_service.dart';
import '../../models/committee.dart';
import '../../models/member.dart';
import '../../utils/app_theme.dart';
import 'member_management_screen.dart';
import '../viewer/member_dashboard_screen.dart';

class PaymentSheetScreen extends StatefulWidget {
  final Committee committee;
  final Member? viewAsMember;

  const PaymentSheetScreen({
    super.key,
    required this.committee,
    this.viewAsMember,
  });

  @override
  State<PaymentSheetScreen> createState() => _PaymentSheetScreenState();
}

class _PaymentSheetScreenState extends State<PaymentSheetScreen> {
  final _dbService = DatabaseService();
  final _authService = AuthService();
  final _exportService = ExportService();
  final _autoSyncService = AutoSyncService();
  final SyncService _syncService = SyncService();

  List<Member> _members = [];
  List<DateTime> _dates = [];
  Map<String, Map<String, bool>> _paymentGrid = {};
  bool _isLoading = true;
  int _selectedCycle = 1;
  int _maxCycles = 1;

  // Date filter
  DateTime? _filterStartDate;
  DateTime? _filterEndDate;

  // Number of extra future periods to show (for advance payments)
  int _extraPeriods = 1;

  @override
  void initState() {
    super.initState();
    _syncAndLoad();
  }

  Future<void> _syncAndLoad() async {
    setState(() => _isLoading = true);
    
    // Sync from cloud first
    try {
      await _syncService.syncMembers(widget.committee.id);
      await _syncService.syncPayments(widget.committee.id);
    } catch (e) {
      debugPrint('Sync error: $e');
    }
    
    await _loadData();
  }

  Future<void> _loadData() async {
    if (!mounted) return;
    setState(() => _isLoading = true);

    _members = _dbService.getMembersByCommittee(widget.committee.id);
    _members.sort((a, b) => a.payoutOrder.compareTo(b.payoutOrder));

    final allPayments = _dbService.getPaymentsByCommittee(widget.committee.id);
    _maxCycles = _computeMaxCycles(allPayments);

    _selectedCycle = _dbService.getSelectedCycle(widget.committee.id);
    if (_selectedCycle < 1) _selectedCycle = 1;
    if (_selectedCycle > _maxCycles) {
      _selectedCycle = _maxCycles;
      _dbService.setSelectedCycle(widget.committee.id, _selectedCycle);
    }

    _generateDates();
    _loadPayments();

    if (mounted) setState(() => _isLoading = false);
  }

  // Helper to safely add months without skipping (e.g., Jan 31 -> Feb 28)
  DateTime _addMonths(DateTime date, int monthsToAdd) {
    var newYear = date.year;
    var newMonth = date.month + monthsToAdd;

    // Adjust year if month goes over 12
    while (newMonth > 12) {
      newYear++;
      newMonth -= 12;
    }

    // Determine the last day of the target month
    final firstDayOfNextMonth = DateTime(newYear, newMonth + 1, 1);
    final lastDayOfTargetMonth =
        firstDayOfNextMonth.subtract(const Duration(days: 1)).day;

    // Clamp the day
    final newDay =
        (date.day > lastDayOfTargetMonth) ? lastDayOfTargetMonth : date.day;

    return DateTime(newYear, newMonth, newDay);
  }

  void _generateDates() {
    _dates = [];
    final committeeStartDate = widget.committee.startDate;
    final payoutIntervalDays = widget.committee.paymentIntervalDays;

    // Use collection frequency to determine interval between collection dates
    int collectionInterval = 30;
    if (widget.committee.frequency == 'daily') collectionInterval = 1;
    if (widget.committee.frequency == 'weekly') collectionInterval = 7;
    if (widget.committee.frequency == 'monthly') collectionInterval = 30;

    // Calculate how many collection periods fit in one payout cycle
    // e.g., payoutIntervalDays=30 with weekly collection (7) = ~4 columns per payout
    // e.g., payoutIntervalDays=30 with monthly collection (30) = 1 column per payout
    // e.g., payoutIntervalDays=30 with daily collection (1) = 30 columns per payout
    int periodsPerPayout;
    if (widget.committee.frequency == 'monthly') {
      // For monthly collection, calculate roughly how many months per payout
      periodsPerPayout = (payoutIntervalDays / 30).ceil();
      if (periodsPerPayout < 1) periodsPerPayout = 1;
    } else {
      periodsPerPayout = (payoutIntervalDays / collectionInterval).ceil();
      if (periodsPerPayout < 1) periodsPerPayout = 1;
    }

    // If committee has members, show only the current selected payout cycle's dates
    final numMembers =
        _members.isNotEmpty
            ? _members.length
            : (widget.committee.totalMembers > 0
                ? widget.committee.totalMembers
                : 0);

    if (numMembers > 0) {
      // Calculate start date for the selected payout cycle
      // Payout 1 starts at committeeStartDate
      // Payout 2 starts after payoutIntervalDays from startDate
      // etc.
      DateTime cycleStartDate;
      if (widget.committee.frequency == 'monthly') {
        cycleStartDate = _addMonths(committeeStartDate, (_selectedCycle - 1) * periodsPerPayout);
      } else {
        final daysOffset = (_selectedCycle - 1) * payoutIntervalDays;
        cycleStartDate = committeeStartDate.add(Duration(days: daysOffset));
      }

      // Apply filter start date if set
      final effectiveStartDate = (_filterStartDate != null && _filterStartDate!.isAfter(cycleStartDate))
          ? _filterStartDate!
          : cycleStartDate;

      // Generate exactly periodsPerPayout dates for this payout cycle
      DateTime current = effectiveStartDate;
      for (int i = 0; i < periodsPerPayout; i++) {
        _dates.add(current);
        if (widget.committee.frequency == 'monthly') {
          current = _addMonths(current, 1);
        } else {
          current = current.add(Duration(days: collectionInterval));
        }
      }
    } else {
      // Fallback: no members, use extraPeriods logic
      final startDate = _filterStartDate ?? committeeStartDate;
      DateTime baseEndDate = _filterEndDate ?? DateTime.now();

      DateTime endDate;
      if (widget.committee.frequency == 'monthly') {
        endDate = _addMonths(baseEndDate, _extraPeriods);
      } else {
        endDate = baseEndDate.add(Duration(days: collectionInterval * _extraPeriods));
      }

      DateTime current = DateTime(
        startDate.year,
        startDate.month,
        startDate.day,
      );

      // Safety break to prevent infinite loops if dates are messed up
      int safetyCounter = 0;

      while (!current.isAfter(endDate) && safetyCounter < 500) {
        _dates.add(current);
        if (widget.committee.frequency == 'monthly') {
          current = _addMonths(current, 1);
        } else {
          current = current.add(Duration(days: collectionInterval));
        }
        safetyCounter++;
      }

      if (_dates.isEmpty) {
        _dates.add(current);
      }
    }
  }

  void _loadPayments() {
    _paymentGrid = {};
    final payments = _dbService.getPaymentsByCommittee(widget.committee.id);
    for (final payment in payments) {
      final dateKey = _getDateKey(payment.date);
      _paymentGrid[payment.memberId] ??= {};
      _paymentGrid[payment.memberId]![dateKey] = payment.isPaid;
    }
    // No setState here; callers manage state
  }

  String _getDateKey(DateTime date) {
    return '${date.year}-${date.month}-${date.day}';
  }

  bool _isPaymentMarked(String memberId, DateTime date) {
    final dateKey = _getDateKey(date);
    return _paymentGrid[memberId]?[dateKey] ?? false;
  }

  // _togglePayment is defined later with host checks and optimistic update.

  int _periodIndexForDate(DateTime date) {
    final start = widget.committee.startDate;
    if (widget.committee.frequency == 'monthly') {
      return (date.year - start.year) * 12 + (date.month - start.month) + 1;
    } else {
      int interval = 30;
      if (widget.committee.frequency == 'daily') interval = 1;
      if (widget.committee.frequency == 'weekly') interval = 7;
      final daysDiff = date.difference(start).inDays;
      return (daysDiff >= 0) ? (daysDiff ~/ interval) + 1 : 1;
    }
  }

  int _computeMaxCycles(List payments) {
    // If we have members, the number of cycles is equal to the number of members
    // (each member receives the payout once per full rotation).
    if (_members.isNotEmpty) {
      return _members.length;
    }

    // Otherwise estimate from existing payments/dates (fallback)
    DateTime farDate =
        _dates.isNotEmpty ? _dates.last : widget.committee.startDate;
    for (final p in payments) {
      if (p.date.isAfter(farDate)) farDate = p.date;
    }

    final periodIndex = _periodIndexForDate(farDate);
    final numMembers =
        widget.committee.totalMembers > 0 ? widget.committee.totalMembers : 1;
    final cycle = ((periodIndex - 1) ~/ numMembers) + 1;

    // Allow a small buffer
    return (cycle + 2) < 1 ? 1 : (cycle + 2);
  }

  // _loadPayments is defined earlier; avoid duplicate definition and restore
  // proper member debt calculator here.

  Map<String, dynamic> _calculateMemberDebt(String memberId) {
    final now = DateTime.now();
    int paidCount = 0;
    int duePeriods = 0;

    for (var date in _dates) {
      if (!date.isAfter(now)) {
        duePeriods++;
        if (_isPaymentMarked(memberId, date)) paidCount++;
      }
    }

    final unpaidCount = duePeriods - paidCount;
    final debtAmount = unpaidCount * widget.committee.contributionAmount;

    return {
      'paidCount': paidCount,
      'duePeriods': duePeriods,
      'unpaidCount': unpaidCount,
      'debtAmount': debtAmount,
      'isDefaulter': unpaidCount > 0,
      'severity':
          unpaidCount >= 3 ? 'high' : (unpaidCount >= 1 ? 'medium' : 'none'),
    };
  }

  double _calculateTotalDebt() {
    double total = 0;
    for (var member in _members) {
      final debt = _calculateMemberDebt(member.id);
      total += debt['debtAmount'] as double;
    }
    return total;
  }

  Map<String, dynamic> _calculateCurrentStats() {
    final now = DateTime.now();
    final amountPerCell = widget.committee.contributionAmount;
    final payoutInterval = widget.committee.paymentIntervalDays;
    final startDate = widget.committee.startDate;

    final daysElapsed = now.difference(startDate).inDays;
    final currentPayoutCycle =
        payoutInterval > 0 ? (daysElapsed ~/ payoutInterval) : 0;

    int currentCyclePaid = 0;
    int currentCycleDue = 0;
    int totalPaid = 0;
    int totalDue = 0;

    for (var member in _members) {
      for (var date in _dates) {
        if (!date.isAfter(now)) {
          totalDue++;
          if (_isPaymentMarked(member.id, date)) {
            totalPaid++;
          }

          final dateDaysElapsed = date.difference(startDate).inDays;
          final datePayoutCycle =
              payoutInterval > 0 ? (dateDaysElapsed ~/ payoutInterval) : 0;

          if (datePayoutCycle == currentPayoutCycle) {
            currentCycleDue++;
            if (_isPaymentMarked(member.id, date)) {
              currentCyclePaid++;
            }
          }
        }
      }
    }

    return {
      'totalPaid': totalPaid,
      'totalDue': totalDue,
      'totalUnpaid': totalDue - totalPaid,
      'totalCollected': totalPaid * amountPerCell,
      'totalPending': (totalDue - totalPaid) * amountPerCell,
      'currentCyclePaid': currentCyclePaid,
      'currentCycleDue': currentCycleDue,
      'currentCycleCollected': currentCyclePaid * amountPerCell,
      'currentPayoutCycle': currentPayoutCycle + 1,
    };
  }

  int _getMinimumPeriods() {
    final now = DateTime.now();
    int count = 0;
    for (var date in _dates) {
      if (!date.isAfter(now)) {
        count++;
      }
    }
    return count > 0 ? count : 1;
  }

  Map<String, dynamic> _calculateMemberAdvance(String memberId) {
    final now = DateTime.now();
    int advancePaymentCount = 0;
    for (var date in _dates) {
      if (date.isAfter(now) && _isPaymentMarked(memberId, date)) {
        advancePaymentCount++;
      }
    }
    final advanceAmount =
        advancePaymentCount * widget.committee.contributionAmount;
    return {
      'advanceCount': advancePaymentCount,
      'advanceAmount': advanceAmount,
      'hasAdvance': advancePaymentCount > 0,
    };
  }

  Future<void> _togglePayment(String memberId, DateTime date) async {
    final currentUser = _authService.currentUser;
    if (currentUser?.uid != widget.committee.hostId) {
      if (!mounted) return;
      ToastService.warning(context, 'Only the host can mark payments');
      return;
    }

    // Optimistic Update: Toggle immediately in UI
    setState(() {
      final dateKey = _getDateKey(date);
      _paymentGrid[memberId] ??= {};
      final currentStatus = _paymentGrid[memberId]![dateKey] ?? false;
      _paymentGrid[memberId]![dateKey] = !currentStatus;
    });

    try {
      final hostId = currentUser?.uid ?? '';
      await _autoSyncService.togglePayment(
        memberId,
        widget.committee.id,
        date,
        hostId,
      );
      
      // Log analytics event
      final newStatus = _paymentGrid[memberId]?[_getDateKey(date)] ?? false;
      AnalyticsService.logPaymentMarked(
        amount: widget.committee.contributionAmount,
        isPaid: newStatus,
      );

      // Reload to ensure sync
      if (mounted) {
        _loadPayments();
        setState(() {});
      }
    } catch (e) {
      // Revert on error
      if (mounted) {
        _loadPayments(); // Reload from DB to revert
        ToastService.error(context, 'Update failed: $e');
      }
    }
  }

  void _showDateFilterDialog() {
    showDialog(
      context: context,
      builder:
          (context) => AlertDialog(
            backgroundColor: AppTheme.darkCard,
            title: const Text('Filter by Date Range'),
            content: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                ListTile(
                  title: Text(
                    'Start Date',
                    style: TextStyle(color: Colors.grey[400]),
                  ),
                  subtitle: Text(
                    _filterStartDate != null
                        ? DateFormat('dd/MM/yyyy').format(_filterStartDate!)
                        : 'Committee Start',
                    style: const TextStyle(color: Colors.white),
                  ),
                  trailing: const Icon(Icons.calendar_today),
                  onTap: () async {
                    final date = await showDatePicker(
                      context: context,
                      initialDate:
                          _filterStartDate ?? widget.committee.startDate,
                      firstDate: widget.committee.startDate,
                      lastDate: DateTime.now(),
                    );
                    if (date != null && mounted) {
                      setState(() => _filterStartDate = date);
                      _generateDates();
                      _loadPayments();
                      Navigator.pop(context);
                      _showDateFilterDialog();
                    }
                  },
                ),
                ListTile(
                  title: Text(
                    'End Date',
                    style: TextStyle(color: Colors.grey[400]),
                  ),
                  subtitle: Text(
                    _filterEndDate != null
                        ? DateFormat('dd/MM/yyyy').format(_filterEndDate!)
                        : 'Today',
                    style: const TextStyle(color: Colors.white),
                  ),
                  trailing: const Icon(Icons.calendar_today),
                  onTap: () async {
                    final date = await showDatePicker(
                      context: context,
                      initialDate: _filterEndDate ?? DateTime.now(),
                      firstDate: widget.committee.startDate,
                      lastDate: DateTime.now(),
                    );
                    if (date != null && mounted) {
                      setState(() => _filterEndDate = date);
                      _generateDates();
                      _loadPayments();
                      Navigator.pop(context);
                      _showDateFilterDialog();
                    }
                  },
                ),
              ],
            ),
            actions: [
              TextButton(
                onPressed: () {
                  setState(() {
                    _filterStartDate = null;
                    _filterEndDate = null;
                  });
                  _generateDates();
                  _loadPayments();
                  Navigator.pop(context);
                },
                child: const Text('Clear Filter'),
              ),
              ElevatedButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Done'),
              ),
            ],
          ),
    );
  }

  void _showExportOptions() {
    showModalBottomSheet(
      context: context,
      backgroundColor: AppTheme.darkCard,
      shape: const RoundedRectangleBorder(
        borderRadius: BorderRadius.vertical(top: Radius.circular(20)),
      ),
      builder:
          (context) => Container(
            padding: const EdgeInsets.all(20),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Export Payment Sheet',
                  style: GoogleFonts.inter(
                    fontSize: 18,
                    fontWeight: FontWeight.bold,
                    color: Colors.white,
                  ),
                ),
                const SizedBox(height: 8),
                Text(
                  _filterStartDate != null || _filterEndDate != null
                      ? 'Exporting filtered range'
                      : 'Exporting all records',
                  style: TextStyle(color: Colors.grey[400], fontSize: 12),
                ),
                const SizedBox(height: 20),
                ListTile(
                  leading: Container(
                    padding: const EdgeInsets.all(10),
                    decoration: BoxDecoration(
                      color: Colors.red.withOpacity(0.1),
                      borderRadius: BorderRadius.circular(10),
                    ),
                    child: const Icon(Icons.picture_as_pdf, color: Colors.red),
                  ),
                  title: const Text(
                    'Export as PDF',
                    style: TextStyle(color: Colors.white),
                  ),
                  subtitle: Text(
                    'Print or share as document',
                    style: TextStyle(color: Colors.grey[500]),
                  ),
                  onTap: () async {
                    Navigator.pop(context);
                    ToastService.info(context, 'Generating PDF...');
                    await _exportService.exportToPdf(
                      widget.committee,
                      startDate: _filterStartDate,
                      endDate: _filterEndDate,
                    );
                  },
                ),
                ListTile(
                  leading: Container(
                    padding: const EdgeInsets.all(10),
                    decoration: BoxDecoration(
                      color: Colors.green.withOpacity(0.1),
                      borderRadius: BorderRadius.circular(10),
                    ),
                    child: const Icon(Icons.table_chart, color: Colors.green),
                  ),
                  title: const Text(
                    'Export as CSV',
                    style: TextStyle(color: Colors.white),
                  ),
                  subtitle: Text(
                    'Open in Excel or Google Sheets',
                    style: TextStyle(color: Colors.grey[500]),
                  ),
                  onTap: () async {
                    Navigator.pop(context);
                    ToastService.info(context, 'Generating CSV...');
                    await _exportService.exportToCsv(
                      widget.committee,
                      startDate: _filterStartDate,
                      endDate: _filterEndDate,
                    );
                  },
                ),
              ],
            ),
          ),
    );
  }

  @override
  Widget build(BuildContext context) {
    if (widget.viewAsMember != null) {
      return _buildMemberPersonalView();
    }

    final amountPerCell = widget.committee.contributionAmount;

    return Scaffold(
      appBar: AppBar(
        title: const Text('Payment Sheet'),
        actions: [
          IconButton(
            icon: Icon(
              Icons.date_range,
              color:
                  (_filterStartDate != null || _filterEndDate != null)
                      ? AppTheme.primaryColor
                      : null,
            ),
            tooltip: 'Filter by Date',
            onPressed: _showDateFilterDialog,
          ),
          IconButton(
            icon: const Icon(Icons.file_download_outlined),
            tooltip: 'Export',
            onPressed: _showExportOptions,
          ),
          IconButton(icon: const Icon(Icons.refresh), onPressed: _loadData),
        ],
      ),
      body:
          _isLoading
              ? const Center(child: CircularProgressIndicator())
              : _members.isEmpty
              ? _buildEmptyState()
              : Column(
                children: [
                  // Stats Card
                  Builder(
                    builder: (context) {
                      final stats = _calculateCurrentStats();
                      return Container(
                        width: double.infinity,
                        padding: const EdgeInsets.all(16),
                        margin: const EdgeInsets.fromLTRB(16, 16, 16, 8),
                        decoration: BoxDecoration(
                          gradient: LinearGradient(
                            colors: [
                              AppTheme.primaryColor.withAlpha(40),
                              AppTheme.secondaryColor.withAlpha(20),
                            ],
                          ),
                          borderRadius: BorderRadius.circular(12),
                          border: Border.all(
                            color: AppTheme.primaryColor.withAlpha(50),
                          ),
                        ),
                        child: Column(
                          children: [
                            Row(
                              children: [
                                Expanded(
                                  child: _buildStatItem(
                                    icon: Icons.check_circle,
                                    color: AppTheme.secondaryColor,
                                    value: '${stats['totalPaid']}',
                                    label: 'Paid',
                                  ),
                                ),
                                const SizedBox(width: 8),
                                Expanded(
                                  child: _buildStatItem(
                                    icon: Icons.cancel,
                                    color: Colors.grey[500]!,
                                    value: '${stats['totalUnpaid']}',
                                    label: 'Unpaid',
                                  ),
                                ),
                                const SizedBox(width: 8),
                                Expanded(
                                  child: _buildStatItem(
                                    icon: Icons.today,
                                    color: Colors.blue[400]!,
                                    value:
                                        'PKR ${(stats['currentCycleCollected'] as double).toInt()}',
                                    label: 'Payout Amount',
                                  ),
                                ),
                              ],
                            ),
                            const SizedBox(height: 12),
                            Container(
                              padding: const EdgeInsets.symmetric(
                                horizontal: 12,
                                vertical: 8,
                              ),
                              decoration: BoxDecoration(
                                color:
                                    (stats['totalPending'] as double) == 0
                                        ? Colors.green.withOpacity(0.1)
                                        : Colors.red.withOpacity(0.1),
                                borderRadius: BorderRadius.circular(8),
                              ),
                              child: Row(
                                mainAxisAlignment:
                                    MainAxisAlignment.spaceBetween,
                                children: [
                                  Text(
                                    (stats['totalPending'] as double) == 0
                                        ? 'All Paid:'
                                        : 'Pending Dues:',
                                    style: TextStyle(
                                      color: Colors.grey[400],
                                      fontSize: 12,
                                    ),
                                  ),
                                  Text(
                                    (stats['totalPending'] as double) == 0
                                        ? 'No Pending ✓'
                                        : 'PKR ${(stats['totalPending'] as double).toInt()}',
                                    style: TextStyle(
                                      color:
                                          (stats['totalPending'] as double) == 0
                                              ? Colors.green[300]
                                              : Colors.red[300],
                                      fontWeight: FontWeight.bold,
                                      fontSize: 14,
                                    ),
                                  ),
                                ],
                              ),
                            ),
                            const SizedBox(height: 8),
                            Row(
                              mainAxisAlignment: MainAxisAlignment.spaceBetween,
                              children: [
                                Text(
                                  'Collection: ${widget.committee.frequency.toUpperCase()} • Payout: Every ${widget.committee.paymentIntervalDays} Days',
                                  style: GoogleFonts.inter(
                                    fontSize: 10,
                                    color: Colors.grey[500],
                                  ),
                                ),
                              ],
                            ),
                          ],
                        ),
                      );
                    },
                  ),


                  // Payout Selector - styled with radius
                  if (_members.isNotEmpty)
                    Center(
                      child: Container(
                        margin: const EdgeInsets.symmetric(vertical: 8),
                        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                        decoration: BoxDecoration(
                          gradient: LinearGradient(
                            colors: [
                              AppTheme.primaryColor.withAlpha(30),
                              AppTheme.secondaryColor.withAlpha(20),
                            ],
                          ),
                          borderRadius: BorderRadius.circular(20),
                          border: Border.all(
                            color: AppTheme.primaryColor.withAlpha(60),
                            width: 1,
                          ),
                        ),
                        child: DropdownButtonHideUnderline(
                          child: DropdownButton<int>(
                            value: _selectedCycle,
                            isDense: true,
                            icon: Icon(
                              Icons.keyboard_arrow_down_rounded,
                              color: AppTheme.primaryColor,
                              size: 20,
                            ),
                            style: TextStyle(
                              color: Colors.white,
                              fontSize: 14,
                              fontWeight: FontWeight.w600,
                            ),
                            dropdownColor: AppTheme.darkCard,
                            borderRadius: BorderRadius.circular(16),
                            elevation: 8,
                            menuMaxHeight: 300,
                            items: List.generate(
                              _maxCycles,
                              (i) => DropdownMenuItem<int>(
                                value: i + 1,
                                child: Padding(
                                  padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                                  child: Text(
                                    'Payout ${i + 1}',
                                    style: TextStyle(
                                      color: (i + 1) == _selectedCycle
                                          ? AppTheme.primaryColor
                                          : Colors.white,
                                      fontWeight: (i + 1) == _selectedCycle
                                          ? FontWeight.bold
                                          : FontWeight.normal,
                                    ),
                                  ),
                                ),
                              ),
                            ),
                            onChanged: (val) {
                              if (val == null) return;
                              setState(() {
                                _selectedCycle = val;
                                _dbService.setSelectedCycle(
                                  widget.committee.id,
                                  val,
                                );
                                _generateDates();
                                _loadPayments();
                              });
                            },
                          ),
                        ),
                      ),
                    ),






                  // Grid with Proper Scrolling
                  Expanded(
                    child: SingleChildScrollView(
                      scrollDirection: Axis.vertical,
                      child: SingleChildScrollView(
                        scrollDirection: Axis.horizontal,
                        child: _buildGrid(amountPerCell),
                      ),
                    ),
                  ),
                ],
              ),
    );
  }

  Widget _buildGrid(double amountPerCell) {
    final totals = List<double>.generate(_dates.length, (index) {
      double total = 0;
      final date = _dates[index];
      for (var member in _members) {
        if (_isPaymentMarked(member.id, date)) {
          total += amountPerCell;
        }
      }
      return total;
    });

    final targetAmount = amountPerCell * _members.length;
    final payoutInterval = widget.committee.paymentIntervalDays;
    final startDate = widget.committee.startDate;

    return DataTable(
      columnSpacing: 8,
      horizontalMargin: 16,
      headingRowColor: WidgetStateProperty.all(AppTheme.darkCard),
      columns: [
        const DataColumn(
          label: Text('Member', style: TextStyle(fontWeight: FontWeight.bold)),
        ),
        ..._dates.asMap().entries.map((entry) {
          final index = entry.key;
          final date = entry.value;
          final now = DateTime.now();
          final isFutureDate = date.isAfter(now);
          final format =
              widget.committee.frequency == 'monthly'
                  ? DateFormat('MMM')
                  : DateFormat('dd/MM');
          final isTargetMet = totals[index] >= targetAmount;

          final daysElapsed = date.difference(startDate).inDays + 1;
          final isPayoutDay =
              payoutInterval > 0 &&
              daysElapsed > 0 &&
              (daysElapsed % payoutInterval == 0);

          return DataColumn(
            label: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                RotatedBox(
                  quarterTurns: -1,
                  child: Row(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      if (isFutureDate)
                        Icon(Icons.schedule, size: 10, color: Colors.blue[300]),
                      if (isFutureDate) const SizedBox(width: 2),
                      Text(
                        format.format(date),
                        style: TextStyle(
                          fontSize: 11,
                          fontWeight: FontWeight.w600,
                          color:
                              isFutureDate
                                  ? Colors.blue[300]
                                  : isPayoutDay
                                  ? Colors.amber
                                  : null,
                        ),
                      ),
                    ],
                  ),
                ),
                if (totals[index] > 0) ...[
                  const SizedBox(height: 4),
                  Container(
                    width: 4,
                    height: 4,
                    decoration: BoxDecoration(
                      color:
                          isTargetMet ? AppTheme.secondaryColor : Colors.orange,
                      shape: BoxShape.circle,
                    ),
                  ),
                ],
              ],
            ),
          );
        }),
        const DataColumn(
          label: Text(
            'Dues',
            style: TextStyle(fontWeight: FontWeight.bold, color: Colors.red),
          ),
        ),
      ],
      rows: [
        ..._members.asMap().entries.map((entry) {
          final memberIndex = entry.key;
          final member = entry.value;
          final memberDebt = _calculateMemberDebt(member.id);
          final isDefaulter = memberDebt['isDefaulter'] as bool;
          final unpaidCount = memberDebt['unpaidCount'] as int;
          final memberAdvance = _calculateMemberAdvance(member.id);
          final hasAdvance = memberAdvance['hasAdvance'] as bool;
          final advanceCount = memberAdvance['advanceCount'] as int;

          return DataRow(
            cells: [
              DataCell(
                Row(
                  children: [
                    // Payout order number on LEFT
                    Container(
                      padding: const EdgeInsets.symmetric(
                        horizontal: 4,
                        vertical: 2,
                      ),
                      decoration: BoxDecoration(
                        color: Colors.grey[800],
                        borderRadius: BorderRadius.circular(4),
                      ),
                      child: Text(
                        '#${member.payoutOrder}',
                        style: TextStyle(fontSize: 9, color: Colors.grey[400]),
                      ),
                    ),
                    const SizedBox(width: 4),
                    // Warning icon if defaulter
                    if (isDefaulter) ...[
                      const Icon(
                        Icons.warning_amber_rounded,
                        size: 16,
                        color: Colors.amber,
                      ),
                      const SizedBox(width: 4),
                    ],
                    // Member name in MIDDLE
                    SizedBox(
                      width: (isDefaulter || hasAdvance) ? 60 : 70,
                      child: Text(
                        member.name,
                        overflow: TextOverflow.ellipsis,
                        style: TextStyle(
                          fontWeight: FontWeight.w500,
                          color:
                              isDefaulter
                                  ? Colors.amber[300]
                                  : hasAdvance
                                  ? Colors.blue[300]
                                  : null,
                        ),
                      ),
                    ),
                    // Advance count on RIGHT
                    if (hasAdvance) ...[
                      const SizedBox(width: 4),
                      Container(
                        padding: const EdgeInsets.symmetric(
                          horizontal: 4,
                          vertical: 2,
                        ),
                        decoration: BoxDecoration(
                          color: Colors.blue.withOpacity(0.2),
                          borderRadius: BorderRadius.circular(4),
                        ),
                        child: Text(
                          '+$advanceCount',
                          style: TextStyle(
                            fontSize: 8,
                            color: Colors.blue[300],
                            fontWeight: FontWeight.bold,
                          ),
                        ),
                      ),
                    ],
                  ],
                ),
              ),
              ..._dates.asMap().entries.map((dateEntry) {
                final date = dateEntry.value;
                final isPaid = _isPaymentMarked(member.id, date);

                final daysElapsed = date.difference(startDate).inDays;
                final currentRound =
                    payoutInterval > 0 ? (daysElapsed ~/ payoutInterval) : 0;
                final receiverIndex = currentRound % _members.length;
                final isPayoutReceiver = (receiverIndex == memberIndex);

                // Highlight Payout Day
                final isPayoutDay =
                    payoutInterval > 0 &&
                    ((daysElapsed + 1) % payoutInterval == 0);
                final isPayoutCell = isPayoutReceiver && isPayoutDay;

                return DataCell(
                  GestureDetector(
                    onTap: () => _togglePayment(member.id, date),
                    child: Container(
                      width: 32,
                      height: 32,
                      decoration: BoxDecoration(
                        color:
                            isPaid ? AppTheme.secondaryColor : Colors.grey[800],
                        borderRadius: BorderRadius.circular(6),
                        border:
                            isPayoutCell
                                ? Border.all(color: Colors.amber, width: 2)
                                : Border.all(
                                  color:
                                      isPaid
                                          ? AppTheme.secondaryColor
                                          : Colors.grey[700]!,
                                  width: 1,
                                ),
                        boxShadow:
                            isPayoutCell
                                ? [
                                  BoxShadow(
                                    color: Colors.amber.withOpacity(0.3),
                                    blurRadius: 4,
                                  ),
                                ]
                                : null,
                      ),
                      child: Stack(
                        alignment: Alignment.center,
                        children: [
                          if (isPaid)
                            const Icon(
                              Icons.check,
                              color: Colors.white,
                              size: 18,
                            ),
                          if (isPayoutCell && !isPaid)
                            const Icon(
                              Icons.star_outline,
                              color: Colors.amber,
                              size: 18,
                            ),
                          if (isPayoutCell && isPaid)
                            const Positioned(
                              right: 0,
                              top: 0,
                              child: Icon(
                                Icons.star,
                                color: Colors.amber,
                                size: 8,
                              ),
                            ),
                        ],
                      ),
                    ),
                  ),
                );
              }),
              DataCell(
                Container(
                  padding: const EdgeInsets.symmetric(
                    horizontal: 6,
                    vertical: 2,
                  ),
                  decoration: BoxDecoration(
                    color:
                        isDefaulter
                            ? Colors.amber.withOpacity(0.2)
                            : hasAdvance
                            ? Colors.blue.withOpacity(0.2)
                            : Colors.green.withOpacity(0.1),
                    borderRadius: BorderRadius.circular(6),
                  ),
                  child: Column(
                    mainAxisAlignment: MainAxisAlignment.center,
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      if (isDefaulter) ...[
                        Text(
                          'PKR ${(unpaidCount * widget.committee.contributionAmount).toInt()}',
                          style: TextStyle(
                            fontSize: 11,
                            fontWeight: FontWeight.bold,
                            color: Colors.amber[300],
                          ),
                        ),
                        Text(
                          '$unpaidCount unpaid',
                          style: TextStyle(
                            fontSize: 10,
                            color: Colors.grey[400],
                          ),
                        ),
                      ] else if (hasAdvance) ...[
                        Text(
                          '+$advanceCount',
                          style: TextStyle(
                            fontSize: 11,
                            fontWeight: FontWeight.bold,
                            color: Colors.blue[300],
                          ),
                        ),
                      ] else ...[
                        Text(
                          '✓',
                          style: TextStyle(
                            fontSize: 12,
                            color: AppTheme.secondaryColor,
                            fontWeight: FontWeight.bold,
                          ),
                        ),
                      ],
                    ],
                  ),
                ),
              ),
            ],
          );
        }),
        DataRow(
          cells: [
            const DataCell(
              Text(
                'Collected',
                style: TextStyle(
                  fontWeight: FontWeight.bold,
                  fontSize: 12,
                  color: Colors.grey,
                ),
              ),
            ),
            ..._dates.asMap().entries.map((entry) {
              final index = entry.key;
              final total = totals[index];
              final isMet = total >= targetAmount;
              return DataCell(
                Text(
                  '${(total).toInt()}/${(targetAmount).toInt()}',
                  style: TextStyle(
                    fontSize: 9,
                    fontWeight: FontWeight.bold,
                    color: isMet ? AppTheme.secondaryColor : Colors.grey[400],
                  ),
                ),
              );
            }),
            DataCell(
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                decoration: BoxDecoration(
                  color:
                      _calculateTotalDebt() > 0
                          ? Colors.red.withOpacity(0.2)
                          : Colors.green.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(6),
                ),
                child: Text(
                  _calculateTotalDebt() > 0
                      ? 'PKR ${_calculateTotalDebt().toInt()}'
                      : 'All Paid ✓',
                  style: TextStyle(
                    fontSize: 11,
                    fontWeight: FontWeight.bold,
                    color:
                        _calculateTotalDebt() > 0
                            ? Colors.red[300]
                            : AppTheme.secondaryColor,
                  ),
                ),
              ),
            ),
          ],
        ),
      ],
    );
  }

  Widget _buildStatItem({
    required IconData icon,
    required Color color,
    required String value,
    required String label,
  }) {
    return Container(
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(
        color: color.withOpacity(0.1),
        borderRadius: BorderRadius.circular(8),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.center,
        children: [
          Icon(icon, color: color, size: 24),
          const SizedBox(height: 4),
          Text(
            value,
            textAlign: TextAlign.center,
            style: GoogleFonts.inter(
              fontSize: 14,
              fontWeight: FontWeight.bold,
              color: color,
            ),
          ),
          Text(
            label,
            textAlign: TextAlign.center,
            style: GoogleFonts.inter(fontSize: 12, color: Colors.grey[400]),
          ),
        ],
      ),
    );
  }

  Widget _buildEmptyState() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Icon(Icons.grid_off_rounded, size: 80, color: Colors.grey[700]),
          const SizedBox(height: 16),
          Text(
            'No Members Yet',
            style: GoogleFonts.inter(
              fontSize: 20,
              fontWeight: FontWeight.w600,
              color: Colors.grey[400],
            ),
          ),
          const SizedBox(height: 12),
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 40),
            child: Text(
              'Add members to start collecting payments. The Period and Cycle options will appear once members are added.',
              textAlign: TextAlign.center,
              style: TextStyle(color: Colors.grey[500]),
            ),
          ),
          const SizedBox(height: 20),
          ElevatedButton.icon(
            onPressed: () async {
              // Navigate to member management to add members
              await Navigator.push(
                context,
                MaterialPageRoute(
                  builder:
                      (_) =>
                          MemberManagementScreen(committee: widget.committee),
                ),
              );
              _loadData();
            },
            icon: const Icon(Icons.person_add_alt_1),
            label: const Text('Add Members'),
          ),
        ],
      ),
    );
  }

  Widget _buildLegendItem({required Color color, required String label}) {
    return Row(
      children: [
        Container(
          width: 20,
          height: 20,
          decoration: BoxDecoration(
            color: color,
            borderRadius: BorderRadius.circular(4),
          ),
        ),
        const SizedBox(width: 8),
        Text(label, style: TextStyle(color: Colors.grey[400])),
      ],
    );
  }

  Widget _buildMemberPersonalView() {
    final member = widget.viewAsMember!;

    int paidCount = 0;
    int advanceCount = 0;
    final now = DateTime.now();

    for (var date in _dates) {
      if (_isPaymentMarked(member.id, date)) {
        paidCount++;
        if (date.isAfter(now)) advanceCount++;
      }
    }

    final totalDue = _dates.length;
    final totalContribution = paidCount * widget.committee.contributionAmount;
    final advanceAmount = advanceCount * widget.committee.contributionAmount;

    return _MemberCalendarView(
      member: member,
      committee: widget.committee,
      dates: _dates,
      paidCount: paidCount,
      totalDue: totalDue,
      totalContribution: totalContribution,
      advanceCount: advanceCount,
      advanceAmount: advanceAmount,
      isPaymentMarked: _isPaymentMarked,
      onRefresh: _loadData,
    );
  }
}

// -----------------------------------------------------------------------------
// MEMBER CALENDAR VIEW
// -----------------------------------------------------------------------------

class _MemberCalendarView extends StatefulWidget {
  final Committee committee;
  final Member member;
  final List<DateTime> dates;
  final int paidCount;
  final int totalDue;
  final double totalContribution;
  final int advanceCount;
  final double advanceAmount;
  final bool Function(String, DateTime) isPaymentMarked;
  final VoidCallback onRefresh;

  const _MemberCalendarView({
    required this.committee,
    required this.member,
    required this.dates,
    required this.paidCount,
    required this.totalDue,
    required this.totalContribution,
    required this.advanceCount,
    required this.advanceAmount,
    required this.isPaymentMarked,
    required this.onRefresh,
  });

  @override
  State<_MemberCalendarView> createState() => _MemberCalendarViewState();
}

class _MemberCalendarViewState extends State<_MemberCalendarView> {
  late DateTime _selectedMonth;
  DateTime? _selectedDate;

  @override
  void initState() {
    super.initState();
    _selectedMonth = DateTime.now();
    // Check for updates (viewers only, Android only)
    if (!kIsWeb) {
      WidgetsBinding.instance.addPostFrameCallback((_) {
        UpdateService.checkForUpdate(context);
      });
    }
  }

  List<DateTime> _getDaysInMonth(DateTime month) {
    final firstDay = DateTime(month.year, month.month, 1);
    final lastDay = DateTime(month.year, month.month + 1, 0);

    List<DateTime> days = [];
    int startingWeekday = firstDay.weekday % 7; // Sunday = 0

    // Days from prev month
    for (int i = 0; i < startingWeekday; i++) {
      days.add(DateTime(month.year, month.month, 1 - (startingWeekday - i)));
    }

    // Days of current month
    for (int i = 1; i <= lastDay.day; i++) {
      days.add(DateTime(month.year, month.month, i));
    }

    // Days of next month
    int remainingDays = 42 - days.length;
    for (int i = 1; i <= remainingDays; i++) {
      days.add(DateTime(month.year, month.month + 1, i));
    }

    return days;
  }

  bool _isPaymentDateForMember(DateTime date) {
    final frequency = widget.committee.frequency;
    for (var paymentDate in widget.dates) {
      if (frequency == 'monthly') {
        if (paymentDate.year == date.year &&
            paymentDate.month == date.month &&
            paymentDate.day == date.day) {
          return true;
        }
      } else if (frequency == 'weekly') {
        if (_isSameDay(paymentDate, date)) {
          return true;
        }
      } else {
        if (_isSameDay(paymentDate, date)) {
          return true;
        }
      }
    }
    return false;
  }

  bool _isInCurrentMonth(DateTime date) {
    return date.month == _selectedMonth.month &&
        date.year == _selectedMonth.year;
  }

  @override
  Widget build(BuildContext context) {
    final days = _getDaysInMonth(_selectedMonth);
    final weekDays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    return Scaffold(
      backgroundColor: AppTheme.darkBg,
      appBar: AppBar(
        backgroundColor: AppTheme.darkSurface,
        leading: IconButton(
          icon: const Icon(Icons.arrow_back),
          onPressed: () => Navigator.pop(context),
        ),
        title: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(widget.committee.name, style: const TextStyle(fontSize: 18)),
            Text(
              'Welcome, ${widget.member.name}',
              style: TextStyle(fontSize: 14, color: Colors.grey[400]),
            ),
          ],
        ),
        actions: [
          IconButton(
            icon: const Icon(Icons.refresh),
            onPressed: widget.onRefresh,
          ),
        ],
      ),
      body: Column(
        children: [
          // Stats Card
          Container(
            margin: const EdgeInsets.all(16),
            padding: const EdgeInsets.all(16),
            decoration: BoxDecoration(
              gradient: LinearGradient(
                colors: [
                  AppTheme.primaryColor.withOpacity(0.8),
                  AppTheme.secondaryColor.withOpacity(0.8),
                ],
                begin: Alignment.topLeft,
                end: Alignment.bottomRight,
              ),
              borderRadius: BorderRadius.circular(16),
            ),
            child: Column(
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceAround,
                  children: [
                    _buildStatItem(
                      'Paid',
                      '${widget.paidCount} / ${widget.dates.length}',
                      Icons.check_circle,
                    ),
                    Container(width: 1, height: 40, color: Colors.white24),
                    _buildStatItem(
                      'Contributed',
                      'PKR ${widget.totalContribution.toStringAsFixed(0)}',
                      Icons.account_balance_wallet,
                    ),
                    Container(width: 1, height: 40, color: Colors.white24),
                    _buildStatItem(
                      _getFrequencyLabel(),
                      'PKR ${widget.committee.contributionAmount.toStringAsFixed(0)}',
                      Icons.calendar_today,
                    ),
                  ],
                ),
                if (widget.advanceCount > 0) ...[
                  const SizedBox(height: 12),
                  Container(
                    padding: const EdgeInsets.symmetric(
                      horizontal: 12,
                      vertical: 8,
                    ),
                    decoration: BoxDecoration(
                      color: Colors.blue.withOpacity(0.3),
                      borderRadius: BorderRadius.circular(8),
                      border: Border.all(color: Colors.blue.withOpacity(0.5)),
                    ),
                    child: Row(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                        const Icon(Icons.star, color: Colors.white, size: 18),
                        const SizedBox(width: 8),
                        Text(
                          '${widget.advanceCount} advance payment${widget.advanceCount > 1 ? 's' : ''} • +PKR ${widget.advanceAmount.toStringAsFixed(0)}',
                          style: const TextStyle(
                            color: Colors.white,
                            fontWeight: FontWeight.w600,
                            fontSize: 13,
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ],
            ),
          ),

          // Month Navigator
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
            child: Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                IconButton(
                  icon: const Icon(Icons.chevron_left, color: Colors.white),
                  onPressed: () {
                    setState(() {
                      _selectedMonth = DateTime(
                        _selectedMonth.year,
                        _selectedMonth.month - 1,
                      );
                    });
                  },
                ),
                Text(
                  '${_getMonthName(_selectedMonth.month)} ${_selectedMonth.year}',
                  style: const TextStyle(
                    color: Colors.white,
                    fontSize: 18,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                IconButton(
                  icon: const Icon(Icons.chevron_right, color: Colors.white),
                  onPressed: () {
                    setState(() {
                      _selectedMonth = DateTime(
                        _selectedMonth.year,
                        _selectedMonth.month + 1,
                      );
                    });
                  },
                ),
              ],
            ),
          ),

          // Week days header
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 8),
            child: Row(
              children:
                  weekDays
                      .map(
                        (day) => Expanded(
                          child: Center(
                            child: Text(
                              day,
                              style: TextStyle(
                                color: Colors.grey[500],
                                fontSize: 12,
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                          ),
                        ),
                      )
                      .toList(),
            ),
          ),
          const SizedBox(height: 8),

          // Calendar Grid
          Expanded(
            child: Padding(
              padding: const EdgeInsets.symmetric(horizontal: 8),
              child: GridView.builder(
                gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(
                  crossAxisCount: 7,
                  childAspectRatio: 1,
                ),
                itemCount: days.length,
                itemBuilder: (context, index) {
                  final date = days[index];
                  final now = DateTime.now();
                  final isFutureDate = date.isAfter(now);
                  final isPaymentDate = _isPaymentDateForMember(date);
                  final isPaid =
                      isPaymentDate &&
                      widget.isPaymentMarked(widget.member.id, date);
                  final isAdvancePaid = isPaid && isFutureDate;
                  final isCurrentMonth = _isInCurrentMonth(date);
                  final isToday = _isSameDay(date, DateTime.now());
                  final isSelected =
                      _selectedDate != null && _isSameDay(date, _selectedDate!);

                  return GestureDetector(
                    onTap: () {
                      if (isCurrentMonth && isPaymentDate) {
                        setState(() {
                          _selectedDate = date;
                        });
                        _showPaymentDetails(date, isPaid, isAdvancePaid);
                      }
                    },
                    child: Container(
                      margin: const EdgeInsets.all(2),
                      decoration: BoxDecoration(
                        color:
                            isSelected
                                ? AppTheme.primaryColor.withOpacity(0.3)
                                : (isAdvancePaid && isCurrentMonth)
                                ? Colors.blue.withOpacity(0.3)
                                : (isPaid && isCurrentMonth)
                                ? Colors.green.withOpacity(0.2)
                                : (isPaymentDate && isCurrentMonth && !isPaid)
                                ? Colors.orange.withOpacity(0.1)
                                : Colors.transparent,
                        borderRadius: BorderRadius.circular(8),
                        border:
                            isToday
                                ? Border.all(
                                  color: AppTheme.primaryColor,
                                  width: 2,
                                )
                                : isAdvancePaid
                                ? Border.all(color: Colors.blue, width: 1)
                                : null,
                      ),
                      child: Column(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          Text(
                            '${date.day}',
                            style: TextStyle(
                              color:
                                  isCurrentMonth
                                      ? Colors.white
                                      : Colors.grey[700],
                              fontWeight:
                                  isToday ? FontWeight.bold : FontWeight.normal,
                            ),
                          ),
                          if (isPaymentDate && isCurrentMonth) ...[
                            const SizedBox(height: 2),
                            Icon(
                              isPaid
                                  ? Icons.check_circle
                                  : Icons.circle_outlined,
                              size: 14,
                              color: isPaid ? Colors.green : Colors.orange,
                            ),
                          ],
                        ],
                      ),
                    ),
                  );
                },
              ),
            ),
          ),

          // Legend
          Container(
            padding: const EdgeInsets.all(16),
            child: Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                _buildLegendItem(Colors.green, 'Paid'),
                const SizedBox(width: 16),
                _buildLegendItem(Colors.blue, 'Advance'),
                const SizedBox(width: 16),
                _buildLegendItem(Colors.orange, 'Pending'),
              ],
            ),
          ),

          // More Details Button
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
            child: SizedBox(
              width: double.infinity,
              child: ElevatedButton.icon(
                onPressed: () {
                  Navigator.push(
                    context,
                    MaterialPageRoute(
                      builder: (context) => MemberDashboardScreen(
                        committee: widget.committee,
                        member: widget.member,
                      ),
                    ),
                  );
                },
                icon: const Icon(Icons.dashboard_rounded),
                label: const Text('More Details'),
                style: ElevatedButton.styleFrom(
                  backgroundColor: AppTheme.primaryColor,
                  foregroundColor: Colors.white,
                  padding: const EdgeInsets.symmetric(vertical: 14),
                  shape: RoundedRectangleBorder(
                    borderRadius: BorderRadius.circular(12),
                  ),
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildStatItem(String label, String value, IconData icon) {
    return Column(
      children: [
        Icon(icon, color: Colors.white, size: 20),
        const SizedBox(height: 4),
        Text(
          value,
          style: const TextStyle(
            color: Colors.white,
            fontWeight: FontWeight.bold,
            fontSize: 16,
          ),
        ),
        Text(
          label,
          style: TextStyle(color: Colors.white.withOpacity(0.8), fontSize: 12),
        ),
      ],
    );
  }

  Widget _buildLegendItem(Color color, String label) {
    return Row(
      children: [
        Container(
          width: 16,
          height: 16,
          decoration: BoxDecoration(
            color: color.withOpacity(0.3),
            borderRadius: BorderRadius.circular(4),
            border: Border.all(color: color),
          ),
          child: Icon(Icons.check, size: 12, color: color),
        ),
        const SizedBox(width: 8),
        Text(label, style: TextStyle(color: Colors.grey[400])),
      ],
    );
  }

  String _getMonthName(int month) {
    const months = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December',
    ];
    return months[month - 1];
  }

  String _getFrequencyLabel() {
    switch (widget.committee.frequency) {
      case 'daily':
        return 'Per Day';
      case 'weekly':
        return 'Per Week';
      case 'monthly':
        return 'Per Month';
      default:
        return 'Per Period';
    }
  }

  bool _isSameDay(DateTime a, DateTime b) {
    return a.year == b.year && a.month == b.month && a.day == b.day;
  }

  void _showPaymentDetails(DateTime date, bool isPaid, bool isAdvancePaid) {
    int dayNumber = 0;
    for (int i = 0; i < widget.dates.length; i++) {
      if (_isSameDay(widget.dates[i], date)) {
        dayNumber = i + 1;
        break;
      }
    }

    showModalBottomSheet(
      context: context,
      backgroundColor: AppTheme.darkSurface,
      shape: const RoundedRectangleBorder(
        borderRadius: BorderRadius.vertical(top: Radius.circular(20)),
      ),
      builder:
          (context) => Padding(
            padding: const EdgeInsets.all(24),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Container(
                      padding: const EdgeInsets.all(12),
                      decoration: BoxDecoration(
                        color:
                            isAdvancePaid
                                ? Colors.blue.withOpacity(0.2)
                                : isPaid
                                ? Colors.green.withOpacity(0.2)
                                : Colors.orange.withOpacity(0.2),
                        borderRadius: BorderRadius.circular(12),
                      ),
                      child: Icon(
                        isAdvancePaid
                            ? Icons.star
                            : isPaid
                            ? Icons.check_circle
                            : Icons.schedule,
                        color:
                            isAdvancePaid
                                ? Colors.blue
                                : isPaid
                                ? Colors.green
                                : Colors.orange,
                        size: 28,
                      ),
                    ),
                    const SizedBox(width: 16),
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          '${date.day} ${_getMonthName(date.month)} ${date.year}',
                          style: const TextStyle(
                            color: Colors.white,
                            fontSize: 18,
                            fontWeight: FontWeight.bold,
                          ),
                        ),
                        Text(
                          isAdvancePaid
                              ? 'Paid in Advance ⭐'
                              : isPaid
                              ? 'Payment Completed'
                              : 'Payment Pending',
                          style: TextStyle(
                            color:
                                isAdvancePaid
                                    ? Colors.blue
                                    : isPaid
                                    ? Colors.green
                                    : Colors.orange,
                            fontSize: 14,
                          ),
                        ),
                      ],
                    ),
                  ],
                ),
                const SizedBox(height: 24),
                _buildDetailRow(
                  'Amount',
                  'PKR ${widget.committee.contributionAmount.toStringAsFixed(0)}',
                ),
                _buildDetailRow(
                  _getPeriodLabel(),
                  '$_getPeriodPrefix $dayNumber',
                ),
                _buildDetailRow(
                  'Status',
                  isAdvancePaid
                      ? 'Advance ⭐'
                      : isPaid
                      ? 'Paid ✓'
                      : 'Pending',
                ),
                const SizedBox(height: 16),
              ],
            ),
          ),
    );
  }

  String get _getPeriodPrefix {
    switch (widget.committee.frequency) {
      case 'daily':
        return 'Day';
      case 'weekly':
        return 'Week';
      case 'monthly':
        return 'Month';
      default:
        return 'Period';
    }
  }

  String _getPeriodLabel() {
    switch (widget.committee.frequency) {
      case 'daily':
        return 'Day Number';
      case 'weekly':
        return 'Week Number';
      case 'monthly':
        return 'Month Number';
      default:
        return 'Period Number';
    }
  }

  Widget _buildDetailRow(String label, String value) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 8),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: TextStyle(color: Colors.grey[400])),
          Text(
            value,
            style: const TextStyle(
              color: Colors.white,
              fontWeight: FontWeight.w600,
            ),
          ),
        ],
      ),
    );
  }
}
