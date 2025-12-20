import 'package:connectivity_plus/connectivity_plus.dart';
import 'sync_service.dart';
import 'database_service.dart';
import 'realtime_sync_service.dart';
import '../models/committee.dart';
import '../models/member.dart';
import '../models/payment.dart';

/// A wrapper service that automatically syncs data to Firebase
/// whenever changes are made locally.
class AutoSyncService {
  final SyncService _syncService = SyncService();
  final DatabaseService _dbService = DatabaseService();
  final RealtimeSyncService _realtimeSyncService = RealtimeSyncService();

  // ============ COMMITTEE OPERATIONS WITH AUTO-SYNC ============

  Future<void> saveCommittee(Committee committee) async {
    // Save locally first
    await _dbService.saveCommittee(committee);

    // Sync to cloud in background
    _syncInBackground(() async {
      await _syncService.syncCommittees(committee.hostId);
    });
  }

  Future<bool> deleteCommittee(String committeeId, String hostId) async {
    // Mark as pending delete to prevent real-time sync from re-adding
    _realtimeSyncService.markCommitteeForDelete(committeeId);
    
    // Delete locally FIRST for instant UI feedback
    await _dbService.deleteCommittee(committeeId);

    // Delete from cloud in background (non-blocking)
    _syncInBackground(() async {
      try {
        final success = await _syncService.deleteCommitteeFromCloud(committeeId);
        if (success) {
          print('Cloud delete succeeded for committee $committeeId');
        } else {
          print('Cloud delete failed for committee $committeeId');
        }
      } catch (e) {
        print('Cloud delete error: $e');
      }
    });

    return true; // Local delete always succeeds
  }

  Future<void> archiveCommittee(Committee committee) async {
    final archived = committee.copyWith(
      isArchived: true,
      archivedAt: DateTime.now(),
    );
    await saveCommittee(archived);
  }

  Future<void> unarchiveCommittee(Committee committee) async {
    final unarchived = Committee(
      id: committee.id,
      code: committee.code,
      name: committee.name,
      hostId: committee.hostId,
      contributionAmount: committee.contributionAmount,
      frequency: committee.frequency,
      startDate: committee.startDate,
      totalMembers: committee.totalMembers,
      createdAt: committee.createdAt,
      isActive: committee.isActive,
      paymentIntervalDays: committee.paymentIntervalDays,
      isArchived: false,
      archivedAt: null,
    );
    await _dbService.saveCommittee(unarchived);

    _syncInBackground(() async {
      await _syncService.syncCommittees(committee.hostId);
    });
  }

  // ============ MEMBER OPERATIONS WITH AUTO-SYNC ============

  Future<void> saveMember(Member member) async {
    // Mark as pending update to prevent real-time sync from reverting
    _realtimeSyncService.markMemberForUpdate(member.id);
    
    await _dbService.saveMember(member);

    _syncInBackground(() async {
      await _syncService.syncMembers(member.committeeId);
    });
  }

  Future<void> deleteMember(String memberId, String committeeId) async {
    // Mark as pending delete to prevent real-time sync from re-adding
    _realtimeSyncService.markMemberForDelete(memberId);
    
    // Delete locally first for instant UI
    await _dbService.deleteMember(memberId);

    // Delete from cloud
    _syncInBackground(() async {
      try {
        final success = await _syncService.deleteMemberFromCloud(memberId);
        if (success) {
          print('Cloud delete succeeded for member $memberId');
        } else {
          print('Cloud delete failed for member $memberId');
        }
      } catch (e) {
        print('Cloud member delete error: $e');
      }
    });
  }

  Future<void> updateMemberPayoutOrder(
    String memberId,
    int order,
    String committeeId,
  ) async {
    await _dbService.updateMemberPayoutOrder(memberId, order);

    _syncInBackground(() async {
      await _syncService.syncMembers(committeeId);
    });
  }

  // ============ PAYMENT OPERATIONS WITH AUTO-SYNC ============

  Future<void> savePayment(Payment payment) async {
    // Mark as pending update
    _realtimeSyncService.markPaymentForUpdate(payment.id);
    
    await _dbService.savePayment(payment);

    _syncInBackground(() async {
      await _syncService.syncPayments(payment.committeeId);
    });
  }

  Future<void> togglePayment(
    String memberId,
    String committeeId,
    DateTime date,
    String hostId,
  ) async {
    // Mark payment as pending to prevent real-time sync from reverting
    final paymentId = '${memberId}_${date.toIso8601String()}';
    _realtimeSyncService.markPaymentForUpdate(paymentId);
    
    await _dbService.togglePayment(memberId, committeeId, date, hostId);

    _syncInBackground(() async {
      await _syncService.syncPayments(committeeId);
    });
  }

  // ============ HELPER ============

  Future<void> _syncInBackground(Future<void> Function() syncFn) async {
    try {
      final connectivity = await Connectivity().checkConnectivity();
      if (connectivity != ConnectivityResult.none) {
        await syncFn();
      }
    } catch (e) {
      // Silently fail - data is saved locally and will sync on next refresh
      print('Auto-sync failed (will retry on refresh): $e');
    }
  }
}
