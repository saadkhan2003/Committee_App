import 'dart:async';
import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:flutter/foundation.dart' show kIsWeb;
import '../../services/auth_service.dart';
import '../../services/database_service.dart';
import '../../services/sync_service.dart';
import '../../services/auto_sync_service.dart';
import '../../services/realtime_sync_service.dart';
import '../../services/update_service.dart';
import '../../services/toast_service.dart';
import '../../models/committee.dart';
import '../../utils/app_theme.dart';
import '../home_screen.dart';
import 'create_committee_screen.dart';
import 'committee_detail_screen.dart';
import '../viewer/join_committee_screen.dart';
import 'profile_screen.dart';

class HostDashboardScreen extends StatefulWidget {
  const HostDashboardScreen({super.key});

  @override
  State<HostDashboardScreen> createState() => _HostDashboardScreenState();
}

class _HostDashboardScreenState extends State<HostDashboardScreen>
    with SingleTickerProviderStateMixin {
  final _authService = AuthService();
  final _dbService = DatabaseService();
  final _syncService = SyncService();
  final _autoSyncService = AutoSyncService();
  final _realtimeSyncService = RealtimeSyncService();
  List<Committee> _activeCommittees = [];
  List<Committee> _archivedCommittees = [];
  bool _isSyncing = false;
  late TabController _tabController;
  Timer? _emailVerificationTimer;

  @override
  void initState() {
    super.initState();
    _tabController = TabController(length: 2, vsync: this);
    _loadCommittees();
    
    // Start real-time sync listener
    final userId = _authService.currentUser?.uid ?? '';
    if (userId.isNotEmpty) {
      _realtimeSyncService.onDataChanged = _loadCommittees;
      _realtimeSyncService.startListening(userId);
    }
    
    // Trigger silent sync on load to fetch fresh data
    _syncDataSilent();
    
    // Check for app updates (Android only)
    if (!kIsWeb) {
      WidgetsBinding.instance.addPostFrameCallback((_) {
        UpdateService.checkForUpdate(context);
      });
    }
    
    // Start email verification check timer
    _startEmailVerificationCheck();
  }

  void _startEmailVerificationCheck() {
    final user = _authService.currentUser;
    if (user != null && !user.emailVerified) {
      _emailVerificationTimer = Timer.periodic(const Duration(seconds: 3), (timer) async {
        await _authService.reloadUser();
        if (_authService.isEmailVerified) {
          timer.cancel();
          if (mounted) {
            setState(() {}); // Refresh UI to hide banner
            ToastService.success(context, 'Email verified successfully! ✓');
          }
        }
      });
    }
  }

  @override
  void dispose() {
    _emailVerificationTimer?.cancel();
    _realtimeSyncService.stopListening();
    _tabController.dispose();
    super.dispose();
  }

  void _loadCommittees() {
    if (!mounted) return;
    final userId = _authService.currentUser?.uid ?? '';
    final all = _dbService.getHostedCommittees(userId);
    setState(() {
      _activeCommittees = all.where((c) => !c.isArchived).toList();
      _archivedCommittees = all.where((c) => c.isArchived).toList();
    });
  }

  // Silent sync for initState (no SnackBar)
  Future<void> _syncDataSilent() async {
    if (_isSyncing) return;

    setState(() {
      _isSyncing = true;
    });

    final hostId = _authService.currentUser?.uid ?? '';
    final result = await _syncService.syncAll(hostId);

    if (mounted) {
      setState(() {
        _isSyncing = false;
      });

      if (result.success) {
        _loadCommittees();
      }
    }
  }

  // Sync with SnackBar feedback (for manual sync button)
  Future<void> _syncData() async {
    if (_isSyncing) return;

    setState(() {
      _isSyncing = true;
    });

    // Save reference before async operation
    final scaffoldMessenger = ScaffoldMessenger.of(context);

    final hostId = _authService.currentUser?.uid ?? '';
    final result = await _syncService.syncAll(hostId);

    if (mounted) {
      setState(() {
        _isSyncing = false;
      });

      // Only show error toast, not success
      if (!result.success) {
        scaffoldMessenger.showSnackBar(
          SnackBar(
            content: Row(
              children: [
                const Icon(Icons.cloud_off, color: Colors.white, size: 20),
                const SizedBox(width: 12),
                Text(result.message),
              ],
            ),
            backgroundColor: AppTheme.errorColor,
            behavior: SnackBarBehavior.floating,
            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
          ),
        );
      }

      if (result.success) {
        _loadCommittees();
      }
    }
  }

  Future<void> _logout() async {
    await _authService.signOut();
    if (mounted) {
      Navigator.pushAndRemoveUntil(
        context,
        MaterialPageRoute(builder: (context) => const HomeScreen()),
        (route) => false,
      );
    }
  }

  void _showArchivedSheet() {
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
                Row(
                  children: [
                    const Icon(Icons.archive, color: AppTheme.primaryColor),
                    const SizedBox(width: 12),
                    Text(
                      'Archived Committees',
                      style: GoogleFonts.inter(
                        fontSize: 18,
                        fontWeight: FontWeight.bold,
                        color: Colors.white,
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 16),
                if (_archivedCommittees.isEmpty)
                  Padding(
                    padding: const EdgeInsets.all(20),
                    child: Center(
                      child: Text(
                        'No archived committees',
                        style: TextStyle(color: Colors.grey[500]),
                      ),
                    ),
                  )
                else
                  ...(_archivedCommittees
                      .map(
                        (committee) => ListTile(
                          leading: Container(
                            padding: const EdgeInsets.all(8),
                            decoration: BoxDecoration(
                              color: Colors.grey[800],
                              borderRadius: BorderRadius.circular(8),
                            ),
                            child: const Icon(Icons.group, color: Colors.grey),
                          ),
                          title: Text(
                            committee.name,
                            style: const TextStyle(color: Colors.white),
                          ),
                          subtitle: Text(
                            'Archived on ${committee.archivedAt != null ? "${committee.archivedAt!.day}/${committee.archivedAt!.month}/${committee.archivedAt!.year}" : "Unknown"}',
                            style: TextStyle(
                              color: Colors.grey[500],
                              fontSize: 12,
                            ),
                          ),
                          trailing: TextButton(
                            onPressed: () async {
                              await _autoSyncService.unarchiveCommittee(
                                committee,
                              );
                              _loadCommittees();
                              if (mounted) Navigator.pop(context);
                            },
                            child: const Text('Restore'),
                          ),
                          onTap: () {
                            Navigator.pop(context);
                            Navigator.push(
                              context,
                              MaterialPageRoute(
                                builder:
                                    (context) => CommitteeDetailScreen(
                                      committee: committee,
                                    ),
                              ),
                            );
                          },
                        ),
                      )
                      .toList()),
              ],
            ),
          ),
    );
  }

  Future<void> _archiveCommittee(Committee committee) async {
    final confirm = await showDialog<bool>(
      context: context,
      builder:
          (context) => AlertDialog(
            backgroundColor: AppTheme.darkCard,
            title: const Text('Archive Committee?'),
            content: Text(
              'This will move "${committee.name}" to the archived section. You can restore it later.',
            ),
            actions: [
              TextButton(
                onPressed: () => Navigator.pop(context, false),
                child: const Text('Cancel'),
              ),
              ElevatedButton(
                onPressed: () => Navigator.pop(context, true),
                style: ElevatedButton.styleFrom(
                  backgroundColor: AppTheme.primaryColor,
                ),
                child: const Text('Archive'),
              ),
            ],
          ),
    );

    if (confirm == true) {
      await _autoSyncService.archiveCommittee(committee);
      _loadCommittees();
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('${committee.name} archived'),
            action: SnackBarAction(
              label: 'Undo',
              onPressed: () async {
                await _autoSyncService.unarchiveCommittee(committee);
                _loadCommittees();
              },
            ),
          ),
        );
      }
    }
  }

  Future<void> _deleteCommittee(Committee committee) async {
    final confirm = await showDialog<bool>(
      context: context,
      builder:
          (context) => AlertDialog(
            backgroundColor: AppTheme.darkCard,
            title: const Text('Delete Committee?'),
            content: Text(
              'This will permanently delete "${committee.name}" and all its data. This cannot be undone.',
            ),
            actions: [
              TextButton(
                onPressed: () => Navigator.pop(context, false),
                child: const Text('Cancel'),
              ),
              ElevatedButton(
                onPressed: () => Navigator.pop(context, true),
                style: ElevatedButton.styleFrom(
                  backgroundColor: AppTheme.errorColor,
                ),
                child: const Text('Delete'),
              ),
            ],
          ),
    );

    if (confirm == true) {
      final hostId = _authService.currentUser?.uid ?? '';
      await _autoSyncService.deleteCommittee(committee.id, hostId);
      _loadCommittees();
      if (mounted) {
        ToastService.error(context, '${committee.name} deleted');
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final user = _authService.currentUser;
    final displayName =
        user?.displayName ?? user?.email?.split('@')[0] ?? 'Host';

    return Scaffold(
      appBar: AppBar(
        title: const Text('My Committees'),
        automaticallyImplyLeading: false,
        actions: [
          // Profile Icon
          IconButton(
            icon: const Icon(Icons.person_outline),
            tooltip: 'Profile',
            onPressed: () {
              Navigator.push(
                context,
                MaterialPageRoute(
                  builder: (context) => const ProfileScreen(),
                ),
              );
            },
          ),
        ],
      ),
      body: RefreshIndicator(
        onRefresh: _syncData,
        color: AppTheme.primaryColor,
        backgroundColor: AppTheme.darkCard,
        child: ListView(
          padding: const EdgeInsets.only(bottom: 80), // Fab space
          physics: const AlwaysScrollableScrollPhysics(), // Required for Web
          children: [
            // Email Verification Banner
            if (user != null && !user.emailVerified)
              Container(
                margin: const EdgeInsets.fromLTRB(16, 16, 16, 0),
                padding: const EdgeInsets.all(16),
                decoration: BoxDecoration(
                  color: AppTheme.warningColor.withOpacity(0.15),
                  borderRadius: BorderRadius.circular(12),
                  border: Border.all(color: AppTheme.warningColor.withOpacity(0.3)),
                ),
                child: Row(
                  children: [
                    const Icon(Icons.email_outlined, color: AppTheme.warningColor),
                    const SizedBox(width: 12),
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const Text(
                            'Verify your email',
                            style: TextStyle(color: AppTheme.warningColor, fontWeight: FontWeight.bold),
                          ),
                          Text(
                            'Check your inbox for verification link',
                            style: TextStyle(color: Colors.grey[400], fontSize: 12),
                          ),
                        ],
                      ),
                    ),
                    TextButton(
                      onPressed: () async {
                        await _authService.sendEmailVerification();
                        if (mounted) {
                          ToastService.success(context, 'Verification email sent!');
                        }
                      },
                      child: const Text('Resend', style: TextStyle(color: AppTheme.warningColor)),
                    ),
                  ],
                ),
              ),
            // Welcome Header
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(20),
              margin: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                gradient: const LinearGradient(
                  colors: [AppTheme.primaryColor, AppTheme.primaryDark],
                ),
                borderRadius: BorderRadius.circular(16),
                boxShadow: [
                  BoxShadow(
                    color: AppTheme.primaryColor.withOpacity(0.3),
                    blurRadius: 20,
                    offset: const Offset(0, 8),
                  ),
                ],
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    '${_authService.currentUser?.displayName}!',
                    style: GoogleFonts.inter(
                      fontSize: 22,
                      fontWeight: FontWeight.bold,
                      color: Colors.white,
                    ),
                  ),
                  const SizedBox(height: 4),
                  Text(
                    '${_activeCommittees.length} active committees',
                    style: GoogleFonts.inter(
                      fontSize: 14,
                      color: Colors.white70,
                    ),
                  ),
                ],
              ),
            ),

            // View Payments Action Card
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
              child: Material(
                color: Colors.transparent,
                child: InkWell(
                  onTap: () {
                    Navigator.push(
                      context,
                      MaterialPageRoute(
                        builder: (context) => const JoinCommitteeScreen(),
                      ),
                    );
                  },
                  borderRadius: BorderRadius.circular(16),
                  child: Container(
                    padding: const EdgeInsets.all(16),
                    decoration: BoxDecoration(
                      color: AppTheme.darkCard,
                      borderRadius: BorderRadius.circular(16),
                      border: Border.all(
                        color: AppTheme.secondaryColor.withOpacity(0.3),
                      ),
                    ),
                    child: Row(
                      children: [
                        Container(
                          padding: const EdgeInsets.all(12),
                          decoration: BoxDecoration(
                            color: AppTheme.secondaryColor.withOpacity(0.1),
                            borderRadius: BorderRadius.circular(12),
                          ),
                          child: const Icon(
                            Icons.receipt_long_rounded,
                            color: AppTheme.secondaryColor,
                            size: 24,
                          ),
                        ),
                        const SizedBox(width: 16),
                        Expanded(
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text(
                                'Joined a Committee?',
                                style: GoogleFonts.inter(
                                  fontSize: 14,
                                  color: Colors.grey[400],
                                ),
                              ),
                              const SizedBox(height: 2),
                              Text(
                                'View Committee Payments',
                                style: GoogleFonts.inter(
                                  fontSize: 16,
                                  fontWeight: FontWeight.w600,
                                  color: Colors.white,
                                ),
                              ),
                            ],
                          ),
                        ),
                        const Icon(
                          Icons.chevron_right_rounded,
                          color: Colors.grey,
                        ),
                      ],
                    ),
                  ),
                ),
              ),
            ),

            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Text(
                    'Your Hosted Committees',
                    style: GoogleFonts.inter(
                      fontSize: 16,
                      fontWeight: FontWeight.w600,
                      color: Colors.grey[400],
                    ),
                  ),
                  if (_archivedCommittees.isNotEmpty)
                    TextButton.icon(
                      onPressed: () => _showArchivedSheet(),
                      icon: const Icon(Icons.archive_outlined, size: 18),
                      label: Text('Archived (${_archivedCommittees.length})'),
                      style: TextButton.styleFrom(
                        foregroundColor: Colors.grey[400],
                      ),
                    ),
                ],
              ),
            ),

            // Committees List
            if (_activeCommittees.isEmpty)
              Padding(
                padding: const EdgeInsets.only(top: 40),
                child: _buildEmptyState(),
              )
            else
              ..._activeCommittees
                  .map(
                    (committee) => Padding(
                      padding: const EdgeInsets.symmetric(horizontal: 16),
                      child: _buildCommitteeCard(committee),
                    ),
                  )
                  .toList(),
          ],
        ),
      ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () async {
          final result = await Navigator.push(
            context,
            MaterialPageRoute(
              builder: (context) => const CreateCommitteeScreen(),
            ),
          );
          if (result == true) {
            _loadCommittees();
          }
        },
        icon: const Icon(Icons.add),
        label: const Text('New Committee'),
      ),
    );
  }

  Widget _buildEmptyState() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Icon(Icons.group_off_outlined, size: 80, color: Colors.grey[700]),
          const SizedBox(height: 16),
          Text(
            'No Committees Yet',
            style: GoogleFonts.inter(
              fontSize: 20,
              fontWeight: FontWeight.w600,
              color: Colors.grey[400],
            ),
          ),
          const SizedBox(height: 8),
          Text(
            'Create your first committee to get started',
            style: GoogleFonts.inter(fontSize: 14, color: Colors.grey[600]),
          ),
        ],
      ),
    );
  }

  Widget _buildCommitteeCard(Committee committee) {
    final members = _dbService.getMembersByCommittee(committee.id);

    return Card(
      margin: const EdgeInsets.only(bottom: 12),
      child: InkWell(
        onTap: () async {
          await Navigator.push(
            context,
            MaterialPageRoute(
              builder: (context) => CommitteeDetailScreen(committee: committee),
            ),
          );
          _loadCommittees();
        },
        borderRadius: BorderRadius.circular(16),
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Row(
            children: [
              Container(
                width: 56,
                height: 56,
                decoration: BoxDecoration(
                  color: AppTheme.primaryColor.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(14),
                ),
                child: const Icon(
                  Icons.group_rounded,
                  color: AppTheme.primaryColor,
                  size: 28,
                ),
              ),
              const SizedBox(width: 16),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      committee.name,
                      style: GoogleFonts.inter(
                        fontSize: 16,
                        fontWeight: FontWeight.w600,
                        color: Colors.white,
                      ),
                    ),
                    const SizedBox(height: 4),
                    Text(
                      '${members.length} members • ${committee.frequency}',
                      style: GoogleFonts.inter(
                        fontSize: 13,
                        color: Colors.grey[400],
                      ),
                    ),
                    const SizedBox(height: 4),
                    Container(
                      padding: const EdgeInsets.symmetric(
                        horizontal: 8,
                        vertical: 4,
                      ),
                      decoration: BoxDecoration(
                        color: AppTheme.secondaryColor.withOpacity(0.1),
                        borderRadius: BorderRadius.circular(6),
                      ),
                      child: Text(
                        'Code: ${committee.code}',
                        style: GoogleFonts.inter(
                          fontSize: 11,
                          fontWeight: FontWeight.w600,
                          color: AppTheme.secondaryColor,
                        ),
                      ),
                    ),
                  ],
                ),
              ),
              PopupMenuButton<String>(
                icon: const Icon(Icons.more_vert, color: Colors.grey),
                onSelected: (value) {
                  if (value == 'archive') {
                    _archiveCommittee(committee);
                  } else if (value == 'delete') {
                    _deleteCommittee(committee);
                  }
                },
                itemBuilder:
                    (context) => [
                      const PopupMenuItem(
                        value: 'archive',
                        child: Row(
                          children: [
                            Icon(Icons.archive_outlined, size: 20),
                            SizedBox(width: 8),
                            Text('Archive'),
                          ],
                        ),
                      ),
                      const PopupMenuItem(
                        value: 'delete',
                        child: Row(
                          children: [
                            Icon(Icons.delete_outline, size: 20, color: Colors.red),
                            SizedBox(width: 8),
                            Text('Delete', style: TextStyle(color: Colors.red)),
                          ],
                        ),
                      ),
                    ],
              ),
            ],
          ),
        ),
      ),
    );
  }
}
