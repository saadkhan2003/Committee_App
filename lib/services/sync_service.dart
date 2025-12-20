import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:connectivity_plus/connectivity_plus.dart';
import '../models/committee.dart';
import '../models/member.dart';
import '../models/payment.dart';
import 'database_service.dart';

class SyncService {
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;
  final DatabaseService _dbService = DatabaseService();

  // Collection names
  static const String committeesCollection = 'committees';
  static const String membersCollection = 'members';
  static const String paymentsCollection = 'payments';

  // Check if online
  Future<bool> isOnline() async {
    final result = await Connectivity().checkConnectivity();
    return result != ConnectivityResult.none;
  }

  // ============ SYNC ALL DATA ============

  Future<SyncResult> syncAll(String hostId) async {
    if (!await isOnline()) {
      return SyncResult(success: false, message: 'No internet connection');
    }

    try {
      int uploaded = 0;
      int downloaded = 0;

      // Sync committees first (must complete before members/payments)
      final committeesResult = await syncCommittees(hostId);
      uploaded += committeesResult.uploaded;
      downloaded += committeesResult.downloaded;

      // Get all committees to sync
      final committees = _dbService.getHostedCommittees(hostId);
      
      // Sync all members and payments in PARALLEL for speed
      final futures = <Future<SyncCounts>>[];
      for (final committee in committees) {
        futures.add(syncMembers(committee.id));
        futures.add(syncPayments(committee.id));
      }
      
      // Wait for all parallel syncs to complete
      final results = await Future.wait(futures);
      for (final result in results) {
        uploaded += result.uploaded;
        downloaded += result.downloaded;
      }

      return SyncResult(
        success: true,
        message: 'Sync complete!',
        uploaded: uploaded,
        downloaded: downloaded,
      );
    } catch (e) {
      return SyncResult(success: false, message: 'Sync failed: $e');
    }
  }

  // ============ COMMITTEE SYNC ============

  Future<SyncCounts> syncCommittees(String hostId) async {
    int uploaded = 0;
    int downloaded = 0;

    // Upload local committees to Firestore
    final localCommittees = _dbService.getHostedCommittees(hostId);
    for (final committee in localCommittees) {
      await _firestore
          .collection(committeesCollection)
          .doc(committee.id)
          .set(committee.toJson(), SetOptions(merge: true));
      uploaded++;
    }

    // Download committees from Firestore
    final snapshot =
        await _firestore
            .collection(committeesCollection)
            .where('hostId', isEqualTo: hostId)
            .get();

    for (final doc in snapshot.docs) {
      final cloudCommittee = Committee.fromJson(doc.data());
      final localCommittee = _dbService.getCommitteeById(cloudCommittee.id);

      // If cloud is newer or doesn't exist locally, save it
      if (localCommittee == null ||
          cloudCommittee.createdAt.isAfter(localCommittee.createdAt)) {
        await _dbService.saveCommittee(cloudCommittee);
        downloaded++;
      }
    }

    return SyncCounts(uploaded: uploaded, downloaded: downloaded);
  }

  // ============ MEMBER SYNC ============

  Future<SyncCounts> syncMembers(String committeeId) async {
    int uploaded = 0;
    int downloaded = 0;

    // Upload local members to Firestore
    final localMembers = _dbService.getMembersByCommittee(committeeId);
    for (final member in localMembers) {
      await _firestore
          .collection(membersCollection)
          .doc(member.id)
          .set(member.toJson(), SetOptions(merge: true));
      uploaded++;
    }

    // Download members from Firestore
    final snapshot =
        await _firestore
            .collection(membersCollection)
            .where('committeeId', isEqualTo: committeeId)
            .get();

    for (final doc in snapshot.docs) {
      final cloudMember = Member.fromJson(doc.data());
      final localMember = _dbService.getMemberById(cloudMember.id);

      // Check if cloud is newer - compare payoutDate for payout changes
      bool shouldDownload = false;
      
      if (localMember == null) {
        shouldDownload = true;
      } else {
        // Compare payoutDate for payout status changes
        if (cloudMember.payoutDate != null && localMember.payoutDate != null) {
          shouldDownload = cloudMember.payoutDate!.isAfter(localMember.payoutDate!);
        } else if (cloudMember.payoutDate != null && localMember.payoutDate == null) {
          // Cloud has payout, local doesn't - cloud is newer
          shouldDownload = true;
        } else if (cloudMember.payoutDate == null && localMember.payoutDate != null) {
          // Cloud was reverted, local has payout - check createdAt as fallback
          // Don't download - local change is probably being uploaded
          shouldDownload = false;
        } else if (cloudMember.hasReceivedPayout != localMember.hasReceivedPayout) {
          // Payout status differs - prefer cloud if no date to compare
          shouldDownload = cloudMember.hasReceivedPayout;
        }
      }

      if (shouldDownload) {
        await _dbService.saveMember(cloudMember);
        downloaded++;
      }
    }

    return SyncCounts(uploaded: uploaded, downloaded: downloaded);
  }

  // ============ PAYMENT SYNC ============

  Future<SyncCounts> syncPayments(String committeeId) async {
    int uploaded = 0;
    int downloaded = 0;

    // Upload local payments to Firestore
    final localPayments = _dbService.getPaymentsByCommittee(committeeId);
    for (final payment in localPayments) {
      await _firestore
          .collection(paymentsCollection)
          .doc(payment.id)
          .set(payment.toJson(), SetOptions(merge: true));
      uploaded++;
    }

    // Download payments from Firestore
    final snapshot =
        await _firestore
            .collection(paymentsCollection)
            .where('committeeId', isEqualTo: committeeId)
            .get();

    for (final doc in snapshot.docs) {
      final cloudPayment = Payment.fromJson(doc.data());
      // Always take the latest markedAt timestamp
      final existingPayment = _dbService.getPayment(
        cloudPayment.memberId,
        cloudPayment.date,
      );

      if (existingPayment == null ||
          (cloudPayment.markedAt != null &&
              existingPayment.markedAt != null &&
              cloudPayment.markedAt!.isAfter(existingPayment.markedAt!)) ||
          (cloudPayment.markedAt != null && existingPayment.markedAt == null)) {
        await _dbService.savePayment(cloudPayment);
        downloaded++;
      }
    }

    return SyncCounts(uploaded: uploaded, downloaded: downloaded);
  }

  // ============ DELETE FROM CLOUD ============

  Future<bool> deleteCommitteeFromCloud(String committeeId) async {
    if (!await isOnline()) return false;

    try {
      // Delete committee document
      await _firestore
          .collection(committeesCollection)
          .doc(committeeId)
          .delete();

      // Delete all members of this committee
      final membersSnapshot =
          await _firestore
              .collection(membersCollection)
              .where('committeeId', isEqualTo: committeeId)
              .get();
      for (final doc in membersSnapshot.docs) {
        await doc.reference.delete();
      }

      // Delete all payments of this committee
      final paymentsSnapshot =
          await _firestore
              .collection(paymentsCollection)
              .where('committeeId', isEqualTo: committeeId)
              .get();
      for (final doc in paymentsSnapshot.docs) {
        await doc.reference.delete();
      }

      print('Successfully deleted committee $committeeId from cloud');
      return true;
    } catch (e) {
      print('Error deleting committee from cloud: $e');
      return false;
    }
  }

  /// Delete a single member from Firestore
  Future<bool> deleteMemberFromCloud(String memberId) async {
    if (!await isOnline()) return false;

    try {
      await _firestore
          .collection(membersCollection)
          .doc(memberId)
          .delete();

      // Also delete related payments for this member
      final paymentsSnapshot =
          await _firestore
              .collection(paymentsCollection)
              .where('memberId', isEqualTo: memberId)
              .get();
      for (final doc in paymentsSnapshot.docs) {
        await doc.reference.delete();
      }

      print('Successfully deleted member $memberId from cloud');
      return true;
    } catch (e) {
      print('Error deleting member from cloud: $e');
      return false;
    }
  }

  // ============ VIEWER SYNC ============

  Future<Committee?> syncCommitteeByCode(String code) async {
    if (!await isOnline()) return null;

    try {
      // 1. Fetch Committee
      final snapshot =
          await _firestore
              .collection(committeesCollection)
              .where('code', isEqualTo: code)
              .limit(1)
              .get();

      if (snapshot.docs.isEmpty) return null;

      final committee = Committee.fromJson(snapshot.docs.first.data());
      await _dbService.saveCommittee(committee);

      // 2. Sync Members
      await syncMembers(committee.id);

      // 3. Sync Payments
      await syncPayments(committee.id);

      return committee;
    } catch (e) {
      print('Viewer sync error: $e');
      return null;
    }
  }
}

class SyncResult {
  final bool success;
  final String message;
  final int uploaded;
  final int downloaded;

  SyncResult({
    required this.success,
    required this.message,
    this.uploaded = 0,
    this.downloaded = 0,
  });
}

class SyncCounts {
  final int uploaded;
  final int downloaded;

  SyncCounts({this.uploaded = 0, this.downloaded = 0});
}
