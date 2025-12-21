import 'dart:math';
import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import '../../services/database_service.dart';
import '../../services/sync_service.dart';
import '../../services/auto_sync_service.dart';
import '../../services/toast_service.dart';
import '../../models/committee.dart';
import '../../models/member.dart';
import '../../utils/app_theme.dart';

class ShuffleMembersScreen extends StatefulWidget {
  final Committee committee;

  const ShuffleMembersScreen({super.key, required this.committee});

  @override
  State<ShuffleMembersScreen> createState() => _ShuffleMembersScreenState();
}

class _ShuffleMembersScreenState extends State<ShuffleMembersScreen>
    with SingleTickerProviderStateMixin {
  final _dbService = DatabaseService();
  final _syncService = SyncService();
  final _autoSyncService = AutoSyncService();
  List<Member> _members = [];
  bool _isShuffling = false;
  bool _isReordering = false;
  bool _isSyncing = false;
  late AnimationController _animController;

  @override
  void initState() {
    super.initState();
    _animController = AnimationController(
      vsync: this,
      duration: const Duration(milliseconds: 500),
    );
    // Sync from cloud first, then load
    _syncAndLoad();
  }

  @override
  void dispose() {
    _animController.dispose();
    super.dispose();
  }

  Future<void> _syncAndLoad() async {
    if (_isSyncing) return;
    setState(() => _isSyncing = true);
    
    try {
      await _syncService.syncMembers(widget.committee.id);
    } catch (e) {
      debugPrint('Sync error: $e');
    }
    
    _loadMembers();
    if (mounted) {
      setState(() => _isSyncing = false);
    }
  }

  void _loadMembers() {
    if (!mounted) return;
    setState(() {
      _members = _dbService.getMembersByCommittee(widget.committee.id);
      _members.sort((a, b) => a.payoutOrder.compareTo(b.payoutOrder));
    });
  }

  Future<void> _shuffleMembers() async {
    if (_members.isEmpty) return;

    setState(() {
      _isShuffling = true;
    });

    // Animate shuffle effect
    for (int i = 0; i < 10; i++) {
      await Future.delayed(const Duration(milliseconds: 100));
      setState(() {
        _members.shuffle(Random());
      });
    }

    // Assign final order
    for (int i = 0; i < _members.length; i++) {
      final member = _members[i];
      await _autoSyncService.updateMemberPayoutOrder(
        member.id,
        i + 1,
        widget.committee.id,
      );
    }

    _loadMembers();
    setState(() {
      _isShuffling = false;
    });

    if (mounted) {
      ToastService.success(context, 'Payout order shuffled successfully!');
    }
  }

  Future<void> _markPayout(Member member) async {
    final updated = member.copyWith(
      hasReceivedPayout: true,
      payoutDate: DateTime.now(),
    );
    await _autoSyncService.saveMember(updated);
    _loadMembers();
  }

  Future<void> _revertPayout(Member member) async {
    final updated = member.copyWith(
      hasReceivedPayout: false,
      clearPayoutDate: true,
    );
    await _autoSyncService.saveMember(updated);
    _loadMembers();

    if (mounted) {
      ToastService.info(context, '${member.name} payout reverted');
    }
  }

  Future<void> _editPayoutOrder(Member member) async {
    final controller = TextEditingController(
      text: member.payoutOrder > 0 ? member.payoutOrder.toString() : '',
    );

    final result = await showDialog<int>(
      context: context,
      builder: (context) => AlertDialog(
        backgroundColor: AppTheme.darkCard,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        title: Text(
          'Set Order for ${member.name}',
          style: const TextStyle(color: Colors.white),
        ),
        content: TextField(
          controller: controller,
          keyboardType: TextInputType.number,
          autofocus: true,
          decoration: InputDecoration(
            labelText: 'Payout Order',
            hintText: 'e.g., 1',
            helperText: 'Enter 1 to ${_members.length}',
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: const Text('Cancel'),
          ),
          ElevatedButton(
            onPressed: () {
              final value = int.tryParse(controller.text);
              if (value != null && value >= 1 && value <= _members.length) {
                Navigator.pop(context, value);
              }
            },
            child: const Text('Save'),
          ),
        ],
      ),
    );

    if (result != null) {
      await _autoSyncService.updateMemberPayoutOrder(
        member.id,
        result,
        widget.committee.id,
      );
      _loadMembers();
    }
  }

  Future<void> _onReorder(int oldIndex, int newIndex) async {
    if (newIndex > oldIndex) {
      newIndex -= 1;
    }

    setState(() {
      final item = _members.removeAt(oldIndex);
      _members.insert(newIndex, item);
    });

    // Update all payout orders
    for (int i = 0; i < _members.length; i++) {
      await _autoSyncService.updateMemberPayoutOrder(
        _members[i].id,
        i + 1,
        widget.committee.id,
      );
    }
    _loadMembers();
  }

  @override
  Widget build(BuildContext context) {
    final hasShuffle = _members.any((m) => m.payoutOrder > 0);
    final receivedCount = _members.where((m) => m.hasReceivedPayout).length;

    return Scaffold(
      appBar: AppBar(
        title: const Text('Payout Order'),
        centerTitle: true,
        leadingWidth: _isReordering ? 80 : null,
        leading: _isReordering
            ? GestureDetector(
                onTap: () {
                  setState(() {
                    _isReordering = false;
                    _loadMembers();
                  });
                },
                child: Container(
                  alignment: Alignment.center,
                  child: Text(
                    'Cancel',
                    style: TextStyle(
                      color: Colors.grey[400],
                      fontSize: 14,
                    ),
                  ),
                ),
              )
            : null,
        actions: [
          if (hasShuffle)
            Container(
              margin: const EdgeInsets.only(right: 8),
              child: TextButton.icon(
                onPressed: () {
                  setState(() {
                    _isReordering = !_isReordering;
                  });
                },
                icon: Icon(
                  _isReordering ? Icons.check : Icons.edit,
                  size: 18,
                  color: _isReordering ? AppTheme.secondaryColor : Colors.grey[400],
                ),
                label: Text(
                  _isReordering ? 'Done' : 'Edit',
                  style: TextStyle(
                    color: _isReordering ? AppTheme.secondaryColor : Colors.grey[400],
                  ),
                ),
              ),
            ),
        ],
      ),
      body: _members.isEmpty
          ? _buildEmptyState()
          : CustomScrollView(
              slivers: [
                // Header Section
                SliverToBoxAdapter(
                  child: Container(
                    margin: const EdgeInsets.all(16),
                    padding: const EdgeInsets.all(20),
                    decoration: BoxDecoration(
                      gradient: LinearGradient(
                        begin: Alignment.topLeft,
                        end: Alignment.bottomRight,
                        colors: [
                          AppTheme.primaryColor.withAlpha(40),
                          AppTheme.secondaryColor.withAlpha(20),
                        ],
                      ),
                      borderRadius: BorderRadius.circular(20),
                      border: Border.all(
                        color: AppTheme.primaryColor.withAlpha(50),
                        width: 1,
                      ),
                    ),
                    child: Column(
                      children: [
                        // Stats Row
                        Row(
                          mainAxisAlignment: MainAxisAlignment.spaceAround,
                          children: [
                            _buildStatColumn(
                              '${_members.length}',
                              'Members',
                              Icons.people_outline,
                            ),
                            Container(
                              height: 40,
                              width: 1,
                              color: Colors.grey[700],
                            ),
                            _buildStatColumn(
                              '$receivedCount',
                              'Received',
                              Icons.check_circle_outline,
                            ),
                            Container(
                              height: 40,
                              width: 1,
                              color: Colors.grey[700],
                            ),
                            _buildStatColumn(
                              '${_members.length - receivedCount}',
                              'Pending',
                              Icons.pending_outlined,
                            ),
                          ],
                        ),
                        const SizedBox(height: 20),
                        // Shuffle Button
                        SizedBox(
                          width: double.infinity,
                          child: ElevatedButton.icon(
                            onPressed: _isShuffling ? null : _shuffleMembers,
                            icon: _isShuffling
                                ? const SizedBox(
                                    width: 18,
                                    height: 18,
                                    child: CircularProgressIndicator(
                                      strokeWidth: 2,
                                      color: Colors.white,
                                    ),
                                  )
                                : const Icon(Icons.shuffle_rounded, size: 20),
                            label: Text(
                              _isShuffling
                                  ? 'Shuffling...'
                                  : hasShuffle
                                      ? 'Reshuffle Order'
                                      : 'Shuffle Members',
                              style: GoogleFonts.inter(
                                fontSize: 15,
                                fontWeight: FontWeight.w600,
                              ),
                            ),
                            style: ElevatedButton.styleFrom(
                              backgroundColor: AppTheme.primaryColor,
                              foregroundColor: Colors.white,
                              padding: const EdgeInsets.symmetric(vertical: 14),
                              shape: RoundedRectangleBorder(
                                borderRadius: BorderRadius.circular(12),
                              ),
                              elevation: 0,
                            ),
                          ),
                        ),
                        if (_isReordering) ...[
                          const SizedBox(height: 12),
                          Container(
                            padding: const EdgeInsets.symmetric(
                              horizontal: 12,
                              vertical: 8,
                            ),
                            decoration: BoxDecoration(
                              color: Colors.blue.withAlpha(30),
                              borderRadius: BorderRadius.circular(8),
                            ),
                            child: Row(
                              mainAxisSize: MainAxisSize.min,
                              children: [
                                Icon(
                                  Icons.info_outline,
                                  size: 16,
                                  color: Colors.blue[300],
                                ),
                                const SizedBox(width: 8),
                                Text(
                                  'Drag to reorder or tap # to edit',
                                  style: TextStyle(
                                    fontSize: 12,
                                    color: Colors.blue[300],
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ],
                    ),
                  ),
                ),

                // Members List
                SliverPadding(
                  padding: const EdgeInsets.symmetric(horizontal: 16),
                  sliver: _isReordering
                      ? SliverReorderableList(
                          itemCount: _members.length,
                          onReorder: _onReorder,
                          itemBuilder: (context, index) {
                            final member = _members[index];
                            return _buildMemberTile(
                              member,
                              index,
                              key: ValueKey(member.id),
                            );
                          },
                        )
                      : SliverList(
                          delegate: SliverChildBuilderDelegate(
                            (context, index) {
                              final member = _members[index];
                              return _buildMemberTile(member, index);
                            },
                            childCount: _members.length,
                          ),
                        ),
                ),
                const SliverToBoxAdapter(
                  child: SizedBox(height: 20),
                ),
              ],
            ),
    );
  }

  Widget _buildStatColumn(String value, String label, IconData icon) {
    return Column(
      children: [
        Icon(icon, color: AppTheme.primaryColor, size: 22),
        const SizedBox(height: 6),
        Text(
          value,
          style: GoogleFonts.inter(
            fontSize: 20,
            fontWeight: FontWeight.bold,
            color: Colors.white,
          ),
        ),
        Text(
          label,
          style: TextStyle(
            fontSize: 11,
            color: Colors.grey[500],
          ),
        ),
      ],
    );
  }

  Widget _buildEmptyState() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Container(
            padding: const EdgeInsets.all(24),
            decoration: BoxDecoration(
              color: Colors.grey[800]?.withAlpha(50),
              shape: BoxShape.circle,
            ),
            child: Icon(
              Icons.people_outline,
              size: 60,
              color: Colors.grey[600],
            ),
          ),
          const SizedBox(height: 20),
          Text(
            'No Members Yet',
            style: GoogleFonts.inter(
              fontSize: 18,
              fontWeight: FontWeight.w600,
              color: Colors.grey[400],
            ),
          ),
          const SizedBox(height: 8),
          Text(
            'Add members to assign payout order',
            style: TextStyle(
              fontSize: 13,
              color: Colors.grey[600],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildMemberTile(Member member, int index, {Key? key}) {
    final order = member.payoutOrder;
    final hasOrder = order > 0;
    final hasReceived = member.hasReceivedPayout;

    return Container(
      key: key,
      margin: const EdgeInsets.only(bottom: 8),
      decoration: BoxDecoration(
        color: hasReceived
            ? AppTheme.secondaryColor.withAlpha(15)
            : AppTheme.darkCard,
        borderRadius: BorderRadius.circular(14),
        border: Border.all(
          color: hasReceived
              ? AppTheme.secondaryColor.withAlpha(40)
              : Colors.transparent,
          width: 1,
        ),
      ),
      child: Padding(
        padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
        child: Row(
          children: [
            // Order Badge
            GestureDetector(
              onTap: _isReordering ? () => _editPayoutOrder(member) : null,
              child: Container(
                width: 40,
                height: 40,
                decoration: BoxDecoration(
                  gradient: hasOrder
                      ? LinearGradient(
                          begin: Alignment.topLeft,
                          end: Alignment.bottomRight,
                          colors: hasReceived
                              ? [
                                  AppTheme.secondaryColor,
                                  AppTheme.secondaryColor.withAlpha(200),
                                ]
                              : [
                                  AppTheme.primaryColor,
                                  AppTheme.primaryColor.withAlpha(200),
                                ],
                        )
                      : null,
                  color: hasOrder ? null : Colors.grey[700],
                  borderRadius: BorderRadius.circular(10),
                  boxShadow: hasOrder
                      ? [
                          BoxShadow(
                            color: (hasReceived
                                    ? AppTheme.secondaryColor
                                    : AppTheme.primaryColor)
                                .withAlpha(60),
                            blurRadius: 8,
                            offset: const Offset(0, 2),
                          ),
                        ]
                      : null,
                ),
                child: Center(
                  child: hasOrder
                      ? Text(
                          '$order',
                          style: GoogleFonts.inter(
                            fontSize: 15,
                            fontWeight: FontWeight.bold,
                            color: Colors.white,
                          ),
                        )
                      : const Icon(Icons.help_outline,
                          color: Colors.white54, size: 20),
                ),
              ),
            ),
            const SizedBox(width: 12),
            // Name and Status
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    member.name,
                    style: GoogleFonts.inter(
                      fontSize: 14,
                      fontWeight: FontWeight.w600,
                      color: hasReceived ? Colors.grey[400] : Colors.white,
                      decoration: hasReceived ? TextDecoration.lineThrough : null,
                    ),
                    overflow: TextOverflow.ellipsis,
                  ),
                  const SizedBox(height: 2),
                  Row(
                    children: [
                      if (hasReceived)
                        Row(
                          children: [
                            Icon(
                              Icons.check_circle,
                              size: 12,
                              color: AppTheme.secondaryColor,
                            ),
                            const SizedBox(width: 4),
                            Text(
                              'Received',
                              style: TextStyle(
                                fontSize: 11,
                                color: AppTheme.secondaryColor,
                              ),
                            ),
                          ],
                        )
                      else if (hasOrder)
                        Text(
                          'Queue #$order',
                          style: TextStyle(
                            fontSize: 11,
                            color: Colors.grey[600],
                          ),
                        ),
                    ],
                  ),
                ],
              ),
            ),
            // Action Button
            if (hasOrder)
              hasReceived
                  ? IconButton(
                      icon: Icon(
                        Icons.undo_rounded,
                        color: Colors.grey[500],
                        size: 20,
                      ),
                      tooltip: 'Revert Payout',
                      onPressed: () => _revertPayout(member),
                    )
                  : TextButton(
                      onPressed: () => _markPayout(member),
                      style: TextButton.styleFrom(
                        backgroundColor: AppTheme.secondaryColor.withAlpha(30),
                        foregroundColor: AppTheme.secondaryColor,
                        padding: const EdgeInsets.symmetric(
                          horizontal: 12,
                          vertical: 6,
                        ),
                        shape: RoundedRectangleBorder(
                          borderRadius: BorderRadius.circular(8),
                        ),
                      ),
                      child: const Text(
                        'Mark Paid',
                        style: TextStyle(
                          fontSize: 12,
                          fontWeight: FontWeight.w600,
                        ),
                      ),
                    ),
            // Drag Handle
            if (_isReordering) ...[
              const SizedBox(width: 4),
              ReorderableDragStartListener(
                index: index,
                child: Container(
                  padding: const EdgeInsets.all(4),
                  child: Icon(
                    Icons.drag_indicator,
                    color: Colors.grey[600],
                    size: 20,
                  ),
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }
}
