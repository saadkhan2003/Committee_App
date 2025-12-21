import 'dart:async';
import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:flutter/foundation.dart';
import '../models/committee.dart';
import '../models/member.dart';
import '../models/payment.dart';
import 'database_service.dart';

/// Helper to parse dates that can be Timestamp or String
DateTime? _parseDate(dynamic value) {
  if (value == null) return null;
  if (value is Timestamp) return value.toDate();
  if (value is String) return DateTime.tryParse(value);
  return null;
}

/// Real-time sync service that listens to Firestore changes
/// and updates local database automatically
class RealtimeSyncService {
  static final RealtimeSyncService _instance = RealtimeSyncService._internal();
  factory RealtimeSyncService() => _instance;
  RealtimeSyncService._internal();

  final FirebaseFirestore _firestore = FirebaseFirestore.instance;
  final DatabaseService _dbService = DatabaseService();

  // Stream subscriptions
  StreamSubscription<QuerySnapshot>? _committeesSubscription;
  final Map<String, StreamSubscription<QuerySnapshot>> _membersSubscriptions = {};
  final Map<String, StreamSubscription<QuerySnapshot>> _paymentsSubscriptions = {};

  // Track pending deletes to prevent re-syncing deleted items
  final Set<String> _pendingCommitteeDeletes = {};
  final Set<String> _pendingMemberDeletes = {};
  final Set<String> _pendingPaymentUpdates = {};  // For payment toggles and reverts

  // Callback for UI updates
  VoidCallback? onDataChanged;

  bool _isListening = false;
  String? _currentHostId;

  /// Mark a committee as pending delete (call before deleting)
  void markCommitteeForDelete(String committeeId) {
    _pendingCommitteeDeletes.add(committeeId);
    // Auto-clear after 10 seconds (in case cloud delete fails)
    Future.delayed(const Duration(seconds: 10), () {
      _pendingCommitteeDeletes.remove(committeeId);
    });
  }

  /// Mark a member as pending delete
  void markMemberForDelete(String memberId) {
    _pendingMemberDeletes.add(memberId);
    Future.delayed(const Duration(seconds: 10), () {
      _pendingMemberDeletes.remove(memberId);
    });
  }

  /// Mark a member as pending update (for payout revert and edits)
  void markMemberForUpdate(String memberId) {
    _pendingMemberDeletes.add(memberId); // Reuse the same set to skip sync
    Future.delayed(const Duration(seconds: 10), () {
      _pendingMemberDeletes.remove(memberId);
    });
  }

  /// Mark a payment as pending update (for toggles and reverts)
  void markPaymentForUpdate(String paymentId) {
    _pendingPaymentUpdates.add(paymentId);
    Future.delayed(const Duration(seconds: 10), () {
      _pendingPaymentUpdates.remove(paymentId);
    });
  }

  /// Check if payment is pending update
  bool isPaymentPendingUpdate(String paymentId) {
    return _pendingPaymentUpdates.contains(paymentId);
  }

  /// Start listening to real-time updates for a host
  void startListening(String hostId) {
    if (_isListening && _currentHostId == hostId) return;
    
    // Stop any existing listeners
    stopListening();

    _currentHostId = hostId;
    _isListening = true;

    debugPrint('🔄 Starting real-time sync for host: $hostId');

    // Listen to committees
    _committeesSubscription = _firestore
        .collection('committees')
        .where('hostId', isEqualTo: hostId)
        .snapshots()
        .listen(
      (snapshot) {
        _handleCommitteesChange(snapshot, hostId);
      },
      onError: (e) => debugPrint('Committees listener error: $e'),
    );
  }

  /// Handle committee changes from Firestore
  void _handleCommitteesChange(QuerySnapshot snapshot, String hostId) async {
    debugPrint('📥 Received ${snapshot.docChanges.length} committee changes');

    for (final change in snapshot.docChanges) {
      final data = change.doc.data() as Map<String, dynamic>?;
      if (data == null) continue;

      final committeeId = change.doc.id;

      // Skip if this committee is pending delete locally
      if (_pendingCommitteeDeletes.contains(committeeId)) {
        debugPrint('⏳ Skipping committee $committeeId (pending delete)');
        continue;
      }

      final committee = Committee(
        id: committeeId,
        code: data['code'] ?? '',
        name: data['name'] ?? '',
        hostId: data['hostId'] ?? '',
        contributionAmount: (data['contributionAmount'] ?? 0).toDouble(),
        frequency: data['frequency'] ?? 'monthly',
        startDate: _parseDate(data['startDate']) ?? DateTime.now(),
        totalMembers: data['totalMembers'] ?? 0,
        createdAt: _parseDate(data['createdAt']) ?? DateTime.now(),
        isActive: data['isActive'] ?? true,
        paymentIntervalDays: data['paymentIntervalDays'] ?? 30,
        isArchived: data['isArchived'] ?? false,
        archivedAt: _parseDate(data['archivedAt']),
      );

      if (change.type == DocumentChangeType.removed) {
        // Committee was deleted on another device
        await _dbService.deleteCommittee(committee.id);
        _stopListeningToCommitteeData(committee.id);
        _pendingCommitteeDeletes.remove(committeeId); // Clear pending if cloud deleted
        debugPrint('🗑️ Committee deleted: ${committee.name}');
      } else {
        // Committee was added or modified
        await _dbService.saveCommittee(committee);
        // Start listening to this committee's members and payments
        _startListeningToCommitteeData(committee.id);
        debugPrint('💾 Committee updated: ${committee.name}');
      }
    }

    // Notify UI to refresh
    onDataChanged?.call();
  }

  /// Start listening to members and payments for a committee
  void _startListeningToCommitteeData(String committeeId) {
    // Skip if already listening
    if (_membersSubscriptions.containsKey(committeeId)) return;

    // Listen to members
    _membersSubscriptions[committeeId] = _firestore
        .collection('members')
        .where('committeeId', isEqualTo: committeeId)
        .snapshots()
        .listen(
      (snapshot) => _handleMembersChange(snapshot, committeeId),
      onError: (e) => debugPrint('Members listener error: $e'),
    );

    // Listen to payments
    _paymentsSubscriptions[committeeId] = _firestore
        .collection('payments')
        .where('committeeId', isEqualTo: committeeId)
        .snapshots()
        .listen(
      (snapshot) => _handlePaymentsChange(snapshot, committeeId),
      onError: (e) => debugPrint('Payments listener error: $e'),
    );
  }

  /// Handle member changes
  void _handleMembersChange(QuerySnapshot snapshot, String committeeId) async {
    for (final change in snapshot.docChanges) {
      final data = change.doc.data() as Map<String, dynamic>?;
      if (data == null) continue;

      final memberId = change.doc.id;

      // Skip if pending delete
      if (_pendingMemberDeletes.contains(memberId)) {
        debugPrint('⏳ Skipping member $memberId (pending delete)');
        continue;
      }

      final member = Member(
        id: memberId,
        committeeId: data['committeeId'] ?? '',
        memberCode: data['memberCode'] ?? '',
        name: data['name'] ?? '',
        phone: data['phone'] ?? '',
        payoutOrder: data['payoutOrder'] ?? 0,
        hasReceivedPayout: data['hasReceivedPayout'] ?? false,
        payoutDate: _parseDate(data['payoutDate']),
        createdAt: _parseDate(data['createdAt']) ?? DateTime.now(),
      );

      if (change.type == DocumentChangeType.removed) {
        await _dbService.deleteMember(member.id);
        _pendingMemberDeletes.remove(memberId);
        debugPrint('🗑️ Member deleted: ${member.name}');
      } else {
        await _dbService.saveMember(member);
        debugPrint('💾 Member updated: ${member.name}');
      }
    }

    onDataChanged?.call();
  }

  /// Handle payment changes
  void _handlePaymentsChange(QuerySnapshot snapshot, String committeeId) async {
    for (final change in snapshot.docChanges) {
      final data = change.doc.data() as Map<String, dynamic>?;
      if (data == null) continue;

      final paymentId = change.doc.id;

      // Skip if pending update locally
      if (_pendingPaymentUpdates.contains(paymentId)) {
        debugPrint('⏳ Skipping payment $paymentId (pending update)');
        continue;
      }

      final payment = Payment(
        id: paymentId,
        memberId: data['memberId'] ?? '',
        committeeId: data['committeeId'] ?? '',
        date: _parseDate(data['date']) ?? DateTime.now(),
        isPaid: data['isPaid'] ?? false,
        markedBy: data['markedBy'] ?? '',
        markedAt: _parseDate(data['markedAt']),
      );

      if (change.type == DocumentChangeType.removed) {
        await _dbService.deletePayment(payment.id);
        _pendingPaymentUpdates.remove(paymentId);
      } else {
        await _dbService.savePayment(payment);
      }
    }

    onDataChanged?.call();
  }

  /// Stop listening to a specific committee's data
  void _stopListeningToCommitteeData(String committeeId) {
    _membersSubscriptions[committeeId]?.cancel();
    _membersSubscriptions.remove(committeeId);
    _paymentsSubscriptions[committeeId]?.cancel();
    _paymentsSubscriptions.remove(committeeId);
  }

  /// Stop all listeners
  void stopListening() {
    debugPrint('⏹️ Stopping real-time sync');
    
    _committeesSubscription?.cancel();
    _committeesSubscription = null;

    for (final sub in _membersSubscriptions.values) {
      sub.cancel();
    }
    _membersSubscriptions.clear();

    for (final sub in _paymentsSubscriptions.values) {
      sub.cancel();
    }
    _paymentsSubscriptions.clear();

    _isListening = false;
    _currentHostId = null;
  }
}
