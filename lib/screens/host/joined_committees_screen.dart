import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import '../../services/database_service.dart';
import '../../services/toast_service.dart';
import '../../utils/app_theme.dart';
import '../viewer/join_committee_screen.dart';
import '../viewer/member_view_screen.dart';
import 'payment_sheet_screen.dart';


class JoinedCommitteesScreen extends StatefulWidget {
  const JoinedCommitteesScreen({super.key});

  @override
  State<JoinedCommitteesScreen> createState() => _JoinedCommitteesScreenState();
}

class _JoinedCommitteesScreenState extends State<JoinedCommitteesScreen> {
  final _dbService = DatabaseService();
  List<Map> _joinedCommittees = [];

  @override
  void initState() {
    super.initState();
    _loadJoinedCommittees();
  }

  void _loadJoinedCommittees() {
    setState(() {
      _joinedCommittees = _dbService.getJoinedCommittees();
    });
  }

  Future<void> _viewCommittee(Map joined) async {
    final committeeCode = joined['committeeCode'] as String;
    final memberCode = joined['memberCode'] as String;

    final committee = _dbService.getCommitteeByCode(committeeCode);
    if (committee == null) {
      ToastService.error(context, 'Kameti not found');
      return;
    }

    final member = _dbService.getMemberByCode(committee.id, memberCode);
    if (member == null) {
      ToastService.error(context, 'Member not found');
      return;
    }

    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => PaymentSheetScreen(
          committee: committee,
        ),
      ),
    );
  }

  Future<void> _removeCommittee(String committeeCode) async {
    await _dbService.removeJoinedCommittee(committeeCode);
    _loadJoinedCommittees();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Joined Committees'),
      ),
      body: _joinedCommittees.isEmpty
          ? _buildEmptyState()
          : ListView.builder(
              padding: const EdgeInsets.all(16),
              itemCount: _joinedCommittees.length,
              itemBuilder: (context, index) {
                final joined = _joinedCommittees[index];
                return _buildCommitteeCard(joined);
              },
            ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () async {
          await Navigator.push(
            context,
            MaterialPageRoute(
              builder: (context) => const JoinCommitteeScreen(),
            ),
          );
          _loadJoinedCommittees();
        },
        icon: const Icon(Icons.add),
        label: const Text('Join New'),
      ),
    );
  }

  Widget _buildEmptyState() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Icon(
            Icons.group_off_outlined,
            size: 80,
            color: Colors.grey[700],
          ),
          const SizedBox(height: 16),
          Text(
            'No Joined Committees',
            style: GoogleFonts.inter(
              fontSize: 20,
              fontWeight: FontWeight.w600,
              color: Colors.grey[400],
            ),
          ),
          const SizedBox(height: 8),
          Text(
            'Join a committee to view your payments',
            style: GoogleFonts.inter(
              fontSize: 14,
              color: Colors.grey[600],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildCommitteeCard(Map joined) {
    final committeeCode = joined['committeeCode'] as String;
    final memberCode = joined['memberCode'] as String;
    final committee = _dbService.getCommitteeByCode(committeeCode);

    return Card(
      margin: const EdgeInsets.only(bottom: 12),
      child: InkWell(
        onTap: () => _viewCommittee(joined),
        borderRadius: BorderRadius.circular(16),
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Row(
            children: [
              Container(
                width: 48,
                height: 48,
                decoration: BoxDecoration(
                  color: AppTheme.secondaryColor.withAlpha(25),
                  borderRadius: BorderRadius.circular(12),
                ),
                child: const Icon(
                  Icons.group_rounded,
                  color: AppTheme.secondaryColor,
                ),
              ),
              const SizedBox(width: 16),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      committee?.name ?? 'Unknown Committee',
                      style: GoogleFonts.inter(
                        fontSize: 16,
                        fontWeight: FontWeight.w600,
                        color: Colors.white,
                      ),
                    ),
                    const SizedBox(height: 4),
                    Text(
                      'Code: $committeeCode • $memberCode',
                      style: GoogleFonts.inter(
                        fontSize: 12,
                        color: Colors.grey[500],
                      ),
                    ),
                  ],
                ),
              ),
              IconButton(
                icon: const Icon(Icons.delete_outline, color: Colors.grey),
                onPressed: () {
                  showDialog(
                    context: context,
                    builder: (context) => AlertDialog(
                      backgroundColor: AppTheme.darkCard,
                      title: const Text('Remove?'),
                      content: const Text(
                        'Remove this committee from your list?',
                      ),
                      actions: [
                        TextButton(
                          onPressed: () => Navigator.pop(context),
                          child: const Text('Cancel'),
                        ),
                        ElevatedButton(
                          onPressed: () {
                            Navigator.pop(context);
                            _removeCommittee(committeeCode);
                          },
                          style: ElevatedButton.styleFrom(
                            backgroundColor: AppTheme.errorColor,
                          ),
                          child: const Text('Remove'),
                        ),
                      ],
                    ),
                  );
                },
              ),
            ],
          ),
        ),
      ),
    );
  }
}
