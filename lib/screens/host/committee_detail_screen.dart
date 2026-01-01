import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:share_plus/share_plus.dart';
import '../../services/database_service.dart';
import '../../services/sync_service.dart';
import '../../services/auto_sync_service.dart';
import '../../models/committee.dart';
import '../../models/member.dart';
import '../../utils/app_theme.dart';
import 'member_management_screen.dart';
import 'payment_sheet_screen.dart';
import 'shuffle_members_screen.dart';

class CommitteeDetailScreen extends StatefulWidget {
  final Committee committee;

  const CommitteeDetailScreen({super.key, required this.committee});

  @override
  State<CommitteeDetailScreen> createState() => _CommitteeDetailScreenState();
}

class _CommitteeDetailScreenState extends State<CommitteeDetailScreen> {
  final _dbService = DatabaseService();
  final _syncService = SyncService();
  final _autoSyncService = AutoSyncService();
  late Committee _committee;
  List<Member> _members = [];
  bool _showCollectionDetails = false;
  bool _isRefreshing = false;

  @override
  void initState() {
    super.initState();
    _committee = widget.committee;
    _loadMembers();
  }

  void _loadMembers() {
    if (!mounted) return;
    setState(() {
      _members = _dbService.getMembersByCommittee(_committee.id);
    });
  }

  Future<void> _refreshCommittee() async {
    if (_isRefreshing) return;
    setState(() => _isRefreshing = true);
    
    try {
      // Sync only this committee's members and payments
      await _syncService.syncMembers(_committee.id);
      await _syncService.syncPayments(_committee.id);
      
      // Reload from local DB
      _loadMembers();
      
      // Refresh committee data from cloud
      final updatedCommittee = _dbService.getCommitteeById(_committee.id);
      if (updatedCommittee != null && mounted) {
        setState(() => _committee = updatedCommittee);
      }
    } catch (e) {
      debugPrint('Refresh error: $e');
    } finally {
      if (mounted) {
        setState(() => _isRefreshing = false);
      }
    }
  }

  void _showShareOptions() {
    // Share committee info only (no member codes)
    String message = '📋 *${_committee.name}*\n\n'
        '*Committee Code:* ${_committee.code}\n'
        '*Contribution:* Rs. ${_committee.contributionAmount.toInt()}\n'
        '*Duration:* ${_members.length} months\n\n'
        '_Download Committee App to view payments!_';
    Share.share(message, subject: '${_committee.name} - Committee Details');
  }

  void _showEditDialog() {
    final nameController = TextEditingController(text: _committee.name);
    final amountController = TextEditingController(
      text: _committee.contributionAmount.toInt().toString(),
    );

    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: AppTheme.darkCard,
      shape: const RoundedRectangleBorder(
        borderRadius: BorderRadius.vertical(top: Radius.circular(20)),
      ),
      builder: (context) {
        return Padding(
          padding: EdgeInsets.fromLTRB(
            24,
            24,
            24,
            MediaQuery.of(context).viewInsets.bottom + 24,
          ),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Text(
                'Edit Committee',
                style: GoogleFonts.inter(
                  fontSize: 20,
                  fontWeight: FontWeight.bold,
                  color: Colors.white,
                ),
              ),
              const SizedBox(height: 20),
              TextField(
                controller: nameController,
                decoration: const InputDecoration(
                  labelText: 'Kameti Name',
                  prefixIcon: Icon(Icons.group_outlined),
                ),
              ),
              const SizedBox(height: 16),
              TextField(
                controller: amountController,
                keyboardType: TextInputType.number,
                decoration: InputDecoration(
                  labelText: 'Contribution Amount',
                  prefixIcon: Container(
                    padding: const EdgeInsets.symmetric(horizontal: 12),
                    alignment: Alignment.center,
                    width: 60,
                    child: Text(
                      'Rs.',
                      style: GoogleFonts.inter(
                        fontWeight: FontWeight.bold,
                        color: Colors.grey[400],
                      ),
                    ),
                  ),
                ),
              ),
              const SizedBox(height: 24),
              ElevatedButton(
                onPressed: () async {
                  final name = nameController.text.trim();
                  final amount =
                      double.tryParse(amountController.text.trim()) ?? 0;

                  if (name.isEmpty || amount <= 0) return;

                  final updated = _committee.copyWith(
                    name: name,
                    contributionAmount: amount,
                  );
                  await _autoSyncService.saveCommittee(updated);
                  setState(() {
                    _committee = updated;
                  });
                  Navigator.pop(context);
                },
                style: ElevatedButton.styleFrom(
                  padding: const EdgeInsets.symmetric(vertical: 16),
                ),
                child: const Text('Save Changes'),
              ),
            ],
          ),
        );
      },
    );
  }

  void _showDeleteConfirmation() {
    // Store reference to the scaffold messenger before showing dialog
    final scaffoldMessenger = ScaffoldMessenger.of(context);
    final navigator = Navigator.of(context);

    showDialog(
      context: context,
      builder:
          (dialogContext) => AlertDialog(
            backgroundColor: AppTheme.darkCard,
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(16),
            ),
            title: const Text(
              'Delete Committee?',
              style: TextStyle(color: Colors.white),
            ),
            content: Text(
              'This will permanently delete "${_committee.name}" and all its members and payment records. This action cannot be undone.',
              style: TextStyle(color: Colors.grey[400]),
            ),
            actions: [
              TextButton(
                onPressed: () => Navigator.pop(dialogContext),
                child: const Text('Cancel'),
              ),
              ElevatedButton(
                onPressed: () async {
                  Navigator.pop(dialogContext); // Close dialog first

                  // Delete locally first (instant), cloud in background
                  await _autoSyncService.deleteCommittee(
                    _committee.id,
                    _committee.hostId,
                  );

                  scaffoldMessenger.showSnackBar(
                    SnackBar(
                      content: const Text('Kameti deleted'),
                      backgroundColor: AppTheme.secondaryColor,
                      behavior: SnackBarBehavior.floating,
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(10),
                      ),
                    ),
                  );

                  // Go back to dashboard immediately
                  navigator.pop();
                },
                style: ElevatedButton.styleFrom(
                  backgroundColor: AppTheme.errorColor,
                ),
                child: const Text('Delete'),
              ),
            ],
          ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final paidMembers = _members.where((m) => m.hasReceivedPayout).length;
    final totalAmount = _committee.contributionAmount * _members.length;

    // Calculate total money collected from payments
    final payments = _dbService.getPaymentsByCommittee(_committee.id);
    final paidPayments = payments.where((p) => p.isPaid).length;
    final totalCollected = paidPayments * _committee.contributionAmount;

    // Calculate per payout amount
    final collectionInterval =
        _committee.frequency == 'daily'
            ? 1
            : _committee.frequency == 'weekly'
            ? 7
            : 30;
    final collectionsPerPayout =
        _committee.paymentIntervalDays ~/ collectionInterval;

    // Calculate actual collected for current payout cycle (match payment sheet logic)
    final now = DateTime.now();
    final payoutInterval = _committee.paymentIntervalDays;
    final startDate = _committee.startDate;
    final daysElapsed = now.difference(startDate).inDays;
    final currentPayoutCycle = daysElapsed ~/ payoutInterval;
    double currentCycleCollected = 0;
    final isMonthlyMonthly =
        _committee.frequency == 'monthly' && payoutInterval == 30;
    for (var member in _members) {
      for (var payment in payments.where(
        (p) => p.memberId == member.id && p.isPaid,
      )) {
        final dateDaysElapsed = payment.date.difference(startDate).inDays;
        final datePayoutCycle = dateDaysElapsed ~/ payoutInterval;
        if (datePayoutCycle == currentPayoutCycle) {
          if (isMonthlyMonthly) {
            // Only count payments for current month, no advance
            if (payment.date.month == now.month &&
                payment.date.year == now.year) {
              currentCycleCollected += _committee.contributionAmount;
            }
          } else {
            // For daily/monthly or weekly/monthly, include advance as they become due
            currentCycleCollected += _committee.contributionAmount;
          }
        }
      }
    }

    return Scaffold(
      appBar: AppBar(
        title: Text(_committee.name),
        actions: [
          IconButton(
            icon: const Icon(Icons.share_rounded),
            onPressed: _showShareOptions,
            tooltip: 'Share Code',
          ),
          PopupMenuButton<String>(
            icon: const Icon(Icons.more_vert),
            onSelected: (value) {
              if (value == 'edit') {
                _showEditDialog();
              } else if (value == 'delete') {
                _showDeleteConfirmation();
              }
            },
            itemBuilder:
                (context) => [
                  const PopupMenuItem(
                    value: 'edit',
                    child: Row(
                      children: [
                        Icon(Icons.edit, size: 18),
                        SizedBox(width: 8),
                        Text('Edit Committee'),
                      ],
                    ),
                  ),
                  const PopupMenuItem(
                    value: 'delete',
                    child: Row(
                      children: [
                        Icon(
                          Icons.delete,
                          size: 18,
                          color: AppTheme.errorColor,
                        ),
                        SizedBox(width: 8),
                        Text(
                          'Delete Committee',
                          style: TextStyle(color: AppTheme.errorColor),
                        ),
                      ],
                    ),
                  ),
                ],
          ),
        ],
      ),
      body: RefreshIndicator(
        onRefresh: _refreshCommittee,
        child: SingleChildScrollView(
          physics: const AlwaysScrollableScrollPhysics(),
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            // Committee Info Card
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(20),
              decoration: BoxDecoration(
                gradient: const LinearGradient(
                  colors: [AppTheme.primaryColor, AppTheme.primaryDark],
                ),
                borderRadius: BorderRadius.circular(16),
              ),
              child: Column(
                children: [
                  Row(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Expanded(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Text(
                              'Kameti Code',
                              style: GoogleFonts.inter(
                                fontSize: 12,
                                color: Colors.white70,
                              ),
                            ),
                            const SizedBox(height: 4),
                            GestureDetector(
                              onTap: _showShareOptions,
                              child: Row(
                                children: [
                                  Text(
                                    _committee.code,
                                    style: GoogleFonts.inter(
                                      fontSize: 28,
                                      fontWeight: FontWeight.bold,
                                      color: Colors.white,
                                      letterSpacing: 3,
                                    ),
                                  ),
                                  const SizedBox(width: 8),
                                  const Icon(
                                    Icons.copy_rounded,
                                    color: Colors.white70,
                                    size: 18,
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(width: 16),
                      Container(
                        padding: const EdgeInsets.symmetric(
                          horizontal: 12,
                          vertical: 6,
                        ),
                        decoration: BoxDecoration(
                          color: Colors.white.withOpacity(0.2),
                          borderRadius: BorderRadius.circular(20),
                        ),
                        child: Text(
                          _committee.frequency.toUpperCase(),
                          style: GoogleFonts.inter(
                            fontSize: 12,
                            fontWeight: FontWeight.w600,
                            color: Colors.white,
                          ),
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 20),
                  Row(
                    children: [
                      _buildStatItem(
                        'Members',
                        '${_members.length}',
                        Icons.people_outline,
                      ),
                      const SizedBox(width: 16),
                      _buildStatItem(
                        'Amount',
                        'Rs. ${_committee.contributionAmount.toInt()}',
                        Icons.payments_outlined,
                      ),
                      const SizedBox(width: 16),
                      _buildStatItem(
                        'Per Cycle',
                        'Rs. ${totalAmount.toInt()}',
                        Icons.account_balance_wallet_outlined,
                      ),
                    ],
                  ),
                ],
              ),
            ),
            const SizedBox(height: 24),

            // Progress Card
            Container(
              padding: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                color: AppTheme.darkCard,
                borderRadius: BorderRadius.circular(12),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'Payout Progress',
                    style: GoogleFonts.inter(
                      fontSize: 14,
                      color: Colors.grey[400],
                    ),
                  ),
                  const SizedBox(height: 12),
                  LinearProgressIndicator(
                    value: _members.isEmpty ? 0 : paidMembers / _members.length,
                    backgroundColor: Colors.grey[800],
                    valueColor: const AlwaysStoppedAnimation<Color>(
                      AppTheme.secondaryColor,
                    ),
                    minHeight: 8,
                    borderRadius: BorderRadius.circular(4),
                  ),
                  const SizedBox(height: 8),
                  Text(
                    '$paidMembers of ${_members.length} members received payout',
                    style: GoogleFonts.inter(
                      fontSize: 12,
                      color: Colors.grey[500],
                    ),
                  ),
                ],
              ),
            ),
            const SizedBox(height: 24),

            // Money Collected Card
            Container(
              padding: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                color: AppTheme.darkCard,
                borderRadius: BorderRadius.circular(12),
                border: Border.all(
                  color: AppTheme.secondaryColor.withAlpha(50),
                ),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  GestureDetector(
                    onTap: () {
                      setState(() {
                        _showCollectionDetails = !_showCollectionDetails;
                      });
                    },
                    child: Row(
                      children: [
                        Container(
                          padding: const EdgeInsets.all(8),
                          decoration: BoxDecoration(
                            color: AppTheme.secondaryColor.withAlpha(25),
                            borderRadius: BorderRadius.circular(8),
                          ),
                          child: const Icon(
                            Icons.account_balance_wallet_rounded,
                            color: AppTheme.secondaryColor,
                            size: 24,
                          ),
                        ),
                        const SizedBox(width: 12),
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Row(
                                children: [
                                  Text(
                                    'Per Payout Collection',
                                    style: GoogleFonts.inter(
                                      fontSize: 12,
                                      color: Colors.grey[400],
                                    ),
                                  ),
                                  const SizedBox(width: 4),
                                  Icon(
                                    _showCollectionDetails
                                        ? Icons.keyboard_arrow_up
                                        : Icons.keyboard_arrow_down,
                                    size: 16,
                                    color: Colors.grey[400],
                                  ),
                                ],
                              ),
                              Text(
                                'Rs. ${currentCycleCollected.toInt().toString().replaceAllMapped(RegExp(r'(\d{1,3})(?=(\d{3})+(?!\d))'), (m) => '${m[1]},')}',
                                style: GoogleFonts.inter(
                                  fontSize: 24,
                                  fontWeight: FontWeight.bold,
                                  color: AppTheme.secondaryColor,
                                ),
                              ),
                            ],
                          ),
                        ),
                      ],
                    ),
                  ),
                  // Expandable details
                  if (_showCollectionDetails) ...[
                    const SizedBox(height: 12),
                    Container(
                      padding: const EdgeInsets.all(12),
                      decoration: BoxDecoration(
                        color: AppTheme.primaryColor.withOpacity(0.1),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: Column(
                        children: [
                          Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              Text(
                                'Total Collected (All Time):',
                                style: GoogleFonts.inter(
                                  fontSize: 12,
                                  color: Colors.grey[400],
                                ),
                              ),
                              Text(
                                'Rs. ${totalCollected.toInt().toString().replaceAllMapped(RegExp(r'(\d{1,3})(?=(\d{3})+(?!\d))'), (m) => '${m[1]},')}',
                                style: GoogleFonts.inter(
                                  fontSize: 14,
                                  fontWeight: FontWeight.bold,
                                  color: AppTheme.primaryColor,
                                ),
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),
                  ],
                  const SizedBox(height: 12),
                  const Divider(color: Colors.grey),
                  const SizedBox(height: 8),
                  Row(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Text(
                        '${_members.length * collectionsPerPayout} payments per payout',
                        style: GoogleFonts.inter(
                          fontSize: 12,
                          color: Colors.grey[500],
                        ),
                      ),
                      Text(
                        '${_members.length} × $collectionsPerPayout collections',
                        style: GoogleFonts.inter(
                          fontSize: 12,
                          color: Colors.grey[500],
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
            const SizedBox(height: 24),

            _buildActionCard(
              icon: Icons.people_rounded,
              title: 'Manage Members',
              subtitle: 'Add, edit or remove members',
              color: AppTheme.primaryColor,
              onTap: () async {
                await Navigator.push(
                  context,
                  MaterialPageRoute(
                    builder:
                        (context) =>
                            MemberManagementScreen(committee: _committee),
                  ),
                );
                _loadMembers();
              },
            ),
            const SizedBox(height: 12),

            _buildActionCard(
              icon: Icons.grid_on_rounded,
              title: 'Payment Sheet',
              subtitle: 'Mark daily payments',
              color: AppTheme.secondaryColor,
              onTap: () {
                Navigator.push(
                  context,
                  MaterialPageRoute(
                    builder:
                        (context) => PaymentSheetScreen(committee: _committee),
                  ),
                );
              },
            ),
            const SizedBox(height: 12),

            _buildActionCard(
              icon: Icons.shuffle_rounded,
              title: 'Shuffle & Assign',
              subtitle: 'Randomly assign payout order',
              color: AppTheme.warningColor,
              onTap: () async {
                await Navigator.push(
                  context,
                  MaterialPageRoute(
                    builder:
                        (context) =>
                            ShuffleMembersScreen(committee: _committee),
                  ),
                );
                _loadMembers();
              },
            ),
          ],
          ),
        ),
      ),
    );
  }

  Widget _buildStatItem(String label, String value, IconData icon) {
    return Expanded(
      child: Container(
        padding: const EdgeInsets.all(12),
        decoration: BoxDecoration(
          color: Colors.white.withOpacity(0.1),
          borderRadius: BorderRadius.circular(10),
        ),
        child: Column(
          children: [
            Icon(icon, color: Colors.white70, size: 20),
            const SizedBox(height: 4),
            Text(
              value,
              style: GoogleFonts.inter(
                fontSize: 14,
                fontWeight: FontWeight.bold,
                color: Colors.white,
              ),
            ),
            Text(
              label,
              style: GoogleFonts.inter(fontSize: 10, color: Colors.white70),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildActionCard({
    required IconData icon,
    required String title,
    required String subtitle,
    required Color color,
    required VoidCallback onTap,
  }) {
    return Card(
      margin: EdgeInsets.zero,
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(16),
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Row(
            children: [
              Container(
                width: 48,
                height: 48,
                decoration: BoxDecoration(
                  color: color.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(12),
                ),
                child: Icon(icon, color: color, size: 24),
              ),
              const SizedBox(width: 16),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      title,
                      style: GoogleFonts.inter(
                        fontSize: 16,
                        fontWeight: FontWeight.w600,
                        color: Colors.white,
                      ),
                    ),
                    Text(
                      subtitle,
                      style: GoogleFonts.inter(
                        fontSize: 13,
                        color: Colors.grey[500],
                      ),
                    ),
                  ],
                ),
              ),
              Icon(Icons.chevron_right_rounded, color: Colors.grey[600]),
            ],
          ),
        ),
      ),
    );
  }
}
