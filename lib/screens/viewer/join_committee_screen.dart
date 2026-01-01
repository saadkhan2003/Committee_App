import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:hive_flutter/hive_flutter.dart';
import '../../services/auth_service.dart';
import '../../services/database_service.dart';
import '../../services/sync_service.dart';
import '../../utils/app_theme.dart';
import '../host/payment_sheet_screen.dart';

class JoinCommitteeScreen extends StatefulWidget {
  const JoinCommitteeScreen({super.key});

  @override
  State<JoinCommitteeScreen> createState() => _JoinCommitteeScreenState();
}

class _JoinCommitteeScreenState extends State<JoinCommitteeScreen> {
  final _committeeCodeController = TextEditingController();
  final _memberCodeController = TextEditingController();
  final _dbService = DatabaseService();
  final _syncService = SyncService();
  final _authService = AuthService();
  String? _errorMessage;
  bool _isLoading = false;

  @override
  void dispose() {
    _committeeCodeController.dispose();
    _memberCodeController.dispose();
    super.dispose();
  }

  List<Map<String, dynamic>> _recentCommittees = [];

  @override
  void initState() {
    super.initState();
    _loadRecentCommittees();
  }

  Future<void> _loadRecentCommittees() async {
    final box = await Hive.openBox('viewer_prefs');
    // Storing as List of Maps: [{'name': '...', 'committeeCode': '...', 'memberCode': '...'}]
    final List<dynamic>? saved = box.get('recent_committees');
    if (saved != null) {
      setState(() {
        _recentCommittees = List<Map<String, dynamic>>.from(
          saved.map((e) => Map<String, dynamic>.from(e)),
        );
      });
    }
  }

  Future<void> _saveRecentCommittee(
    String name,
    String committeeCode,
    String memberCode,
  ) async {
    final box = await Hive.openBox('viewer_prefs');

    // Remove if already exists to move to top
    _recentCommittees.removeWhere(
      (item) =>
          item['committeeCode'] == committeeCode &&
          item['memberCode'] == memberCode,
    );

    // Add to top
    _recentCommittees.insert(0, {
      'name': name,
      'committeeCode': committeeCode,
      'memberCode': memberCode,
      'timestamp': DateTime.now().toIso8601String(),
    });

    // Keep only last 5
    if (_recentCommittees.length > 5) {
      _recentCommittees = _recentCommittees.sublist(0, 5);
    }

    await box.put('recent_committees', _recentCommittees);
    setState(() {});
  }

  Future<void> _deleteRecentCommittee(int index) async {
    final box = await Hive.openBox('viewer_prefs');
    _recentCommittees.removeAt(index);
    await box.put('recent_committees', _recentCommittees);
    setState(() {});
  }

  void _showDeleteConfirmation(int index, String name) {
    showDialog(
      context: context,
      builder:
          (context) => AlertDialog(
            backgroundColor: AppTheme.darkCard,
            title: const Text('Remove from Recent'),
            content: Text('Remove "$name" from your recent kametis?'),
            actions: [
              TextButton(
                onPressed: () => Navigator.pop(context),
                child: const Text('Cancel'),
              ),
              TextButton(
                onPressed: () {
                  Navigator.pop(context);
                  _deleteRecentCommittee(index);
                },
                style: TextButton.styleFrom(foregroundColor: Colors.red),
                child: const Text('Remove'),
              ),
            ],
          ),
    );
  }

  Future<void> _viewPayments({String? code, String? member}) async {
    final committeeCode = code ?? _committeeCodeController.text.trim();
    final memberCode =
        member?.toUpperCase() ??
        _memberCodeController.text.trim().toUpperCase();

    if (committeeCode.isEmpty || memberCode.isEmpty) {
      setState(() {
        _errorMessage = 'Please enter both codes';
      });
      return;
    }

    setState(() {
      _isLoading = true;
      _errorMessage = null;
    });

    try {
      // 0. Ensure Authenticated (Anonymous) for Firestore Security Rules
      // If we are not logged in (as host), sign in anonymously
      if (!_authService.isLoggedIn) {
        try {
          await _authService.signInAnonymously();
        } catch (e) {
          print('Anonymous auth failed: $e');
          // Proceed anyway? Or stop?
          // If auth fails, sync will fail. But maybe we have local data.
        }
      }

      // 1. Try to fetch/sync from Firebase first (Data Freshness)
      try {
        await _syncService.syncCommitteeByCode(committeeCode);
      } catch (e) {
        print('Sync failed, falling back to local: $e');
      }

      // 2. Fetch from local (now ostensibly updated)
      final committee = _dbService.getCommitteeByCode(committeeCode);

      // 3. If missing, it doesn't exist
      if (committee == null) {
        setState(() {
          _errorMessage =
              'Kameti not found. Please check connectivity or the code.';
        });
        return;
      }

      // Find member by code
      final member = _dbService.getMemberByCode(committee.id, memberCode);
      if (member == null) {
        setState(() {
          _errorMessage = 'Member not found. Please check your member code.';
        });
        return;
      }

      // Save to recent list
      await _saveRecentCommittee(committee.name, committeeCode, memberCode);

      // Navigate to member view
      if (mounted) {
        Navigator.push(
          context,
          MaterialPageRoute(
            builder:
                (context) => PaymentSheetScreen(
                  committee: committee,
                  viewAsMember: member,
                ),
          ),
        );
      }
    } finally {
      if (mounted) {
        setState(() {
          _isLoading = false;
        });
      }
    }
  }

  void _fillFromRecent(Map<String, dynamic> item) {
    _committeeCodeController.text = item['committeeCode'];
    _memberCodeController.text = item['memberCode'];
    _viewPayments();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('View Payments'),
        backgroundColor: Colors.transparent,
      ),
      body: Container(
        decoration: BoxDecoration(
          gradient: LinearGradient(
            begin: Alignment.topCenter,
            end: Alignment.bottomCenter,
            colors: [AppTheme.darkBg, AppTheme.darkSurface],
          ),
        ),
        child: SafeArea(
          child: SingleChildScrollView(
            padding: const EdgeInsets.all(24.0),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                const SizedBox(height: 20),

                // Recent Committees Section
                if (_recentCommittees.isNotEmpty) ...[
                  Text(
                    'Recent Kametis',
                    style: GoogleFonts.inter(
                      fontSize: 16,
                      fontWeight: FontWeight.w600,
                      color: Colors.grey[400],
                    ),
                  ),
                  const SizedBox(height: 12),
                  ..._recentCommittees.asMap().entries.map((entry) {
                    final index = entry.key;
                    final item = entry.value;
                    return Card(
                      margin: const EdgeInsets.only(bottom: 12),
                      color: AppTheme.darkCard,
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(12),
                        side: BorderSide(
                          color: AppTheme.primaryColor.withOpacity(0.3),
                        ),
                      ),
                      child: ListTile(
                        onTap: () => _fillFromRecent(item),
                        leading: Container(
                          padding: const EdgeInsets.all(8),
                          decoration: BoxDecoration(
                            color: AppTheme.primaryColor.withOpacity(0.1),
                            borderRadius: BorderRadius.circular(8),
                          ),
                          child: const Icon(
                            Icons.history,
                            color: AppTheme.primaryColor,
                          ),
                        ),
                        title: Text(
                          item['name'] ?? 'Unknown Kameti',
                          style: GoogleFonts.inter(
                            fontWeight: FontWeight.bold,
                            color: Colors.white,
                          ),
                        ),
                        subtitle: Text(
                          'Code: ${item['committeeCode']}',
                          style: GoogleFonts.inter(
                            fontSize: 12,
                            color: Colors.grey[400],
                          ),
                        ),
                        trailing: Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            IconButton(
                              icon: const Icon(
                                Icons.delete_outline,
                                color: Colors.red,
                                size: 20,
                              ),
                              onPressed:
                                  () => _showDeleteConfirmation(
                                    index,
                                    item['name'] ?? 'this kameti',
                                  ),
                              tooltip: 'Remove from recent',
                            ),
                            const Icon(Icons.chevron_right, color: Colors.grey),
                          ],
                        ),
                      ),
                    );
                  }).toList(),
                  const SizedBox(height: 24),
                  const Divider(color: Colors.grey),
                  const SizedBox(height: 24),
                ],

                // Header Icon
                Container(
                  width: 80,
                  height: 80,
                  margin: const EdgeInsets.only(bottom: 24),
                  decoration: BoxDecoration(
                    color: AppTheme.secondaryColor.withOpacity(0.1),
                    borderRadius: BorderRadius.circular(20),
                  ),
                  child: const Icon(
                    Icons.visibility_rounded,
                    size: 40,
                    color: AppTheme.secondaryColor,
                  ),
                ),
                Text(
                  'View Your Payments',
                  style: GoogleFonts.inter(
                    fontSize: 24,
                    fontWeight: FontWeight.bold,
                    color: Colors.white,
                  ),
                ),
                const SizedBox(height: 8),
                Text(
                  'Enter the codes provided by your kameti host',
                  style: GoogleFonts.inter(
                    fontSize: 14,
                    color: Colors.grey[400],
                  ),
                ),
                const SizedBox(height: 32),

                // Error Message
                if (_errorMessage != null)
                  Container(
                    padding: const EdgeInsets.all(12),
                    margin: const EdgeInsets.only(bottom: 16),
                    decoration: BoxDecoration(
                      color: AppTheme.errorColor.withOpacity(0.1),
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(
                        color: AppTheme.errorColor.withOpacity(0.3),
                      ),
                    ),
                    child: Row(
                      children: [
                        const Icon(
                          Icons.error_outline,
                          color: AppTheme.errorColor,
                          size: 20,
                        ),
                        const SizedBox(width: 12),
                        Expanded(
                          child: Text(
                            _errorMessage!,
                            style: const TextStyle(
                              color: AppTheme.errorColor,
                              fontSize: 13,
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),

                // Committee Code Field
                TextFormField(
                  controller: _committeeCodeController,
                  keyboardType: TextInputType.number,
                  maxLength: 6,
                  decoration: const InputDecoration(
                    labelText: 'Kameti Code',
                    prefixIcon: Icon(Icons.group_outlined),
                    hintText: 'e.g., 847293',
                    counterText: '',
                  ),
                  style: GoogleFonts.inter(fontSize: 18, letterSpacing: 2),
                ),
                const SizedBox(height: 16),

                // Member Code Field
                TextFormField(
                  controller: _memberCodeController,
                  textCapitalization: TextCapitalization.characters,
                  onChanged: (value) {
                    // Auto-add dash after 3 letters
                    final upper = value.toUpperCase();
                    if (upper.length == 3 && !upper.contains('-')) {
                      // If exactly 3 letters and no dash, add dash
                      _memberCodeController.text = '$upper-';
                      _memberCodeController.selection = TextSelection.fromPosition(
                        TextPosition(offset: _memberCodeController.text.length),
                      );
                    } else if (upper.length > 4 && !upper.contains('-')) {
                      // If typed more than 3 without dash, insert it
                      final letters = upper.substring(0, 3);
                      final numbers = upper.substring(3);
                      _memberCodeController.text = '$letters-$numbers';
                      _memberCodeController.selection = TextSelection.fromPosition(
                        TextPosition(offset: _memberCodeController.text.length),
                      );
                    }
                  },
                  decoration: const InputDecoration(
                    labelText: 'Your Member Code',
                    prefixIcon: Icon(Icons.person_outline),
                    hintText: 'e.g., ABD-1577',
                  ),
                  style: GoogleFonts.inter(fontSize: 18, letterSpacing: 1),
                ),
                const SizedBox(height: 32),

                // Submit Button
                ElevatedButton.icon(
                  onPressed: _isLoading ? null : () => _viewPayments(),
                  icon:
                      _isLoading
                          ? const SizedBox(
                            height: 20,
                            width: 20,
                            child: CircularProgressIndicator(
                              strokeWidth: 2,
                              color: Colors.white,
                            ),
                          )
                          : const Icon(Icons.visibility),
                  label: Text(_isLoading ? 'Loading...' : 'View My Payments'),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: AppTheme.secondaryColor,
                    padding: const EdgeInsets.symmetric(vertical: 18),
                  ),
                ),
                const SizedBox(height: 24),

                // Help Text
                Container(
                  padding: const EdgeInsets.all(16),
                  decoration: BoxDecoration(
                    color: AppTheme.darkCard.withOpacity(0.5),
                    borderRadius: BorderRadius.circular(12),
                    border: Border.all(color: Colors.grey[800]!, width: 1),
                  ),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Icon(
                            Icons.help_outline,
                            color: Colors.grey[400],
                            size: 18,
                          ),
                          const SizedBox(width: 8),
                          Text(
                            'How to get codes?',
                            style: GoogleFonts.inter(
                              fontSize: 14,
                              fontWeight: FontWeight.w600,
                              color: Colors.grey[300],
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 8),
                      Text(
                        '• Kameti Code: Ask your kameti host\n• Member Code: Your unique code given by the host',
                        style: GoogleFonts.inter(
                          fontSize: 12,
                          color: Colors.grey[500],
                          height: 1.5,
                        ),
                      ),
                    ],
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
