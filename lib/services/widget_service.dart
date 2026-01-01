import 'package:flutter/foundation.dart';
import 'package:home_widget/home_widget.dart';
import 'package:intl/intl.dart';
import '../models/committee.dart';
import '../models/member.dart';
import 'database_service.dart';

/// Service to update the Android home screen widget with payout info
class WidgetService {
  static const String _appGroupId = 'com.kameti.app';
  
  /// Update widget with next payout information
  static Future<void> updateWidget(String hostId) async {
    if (kIsWeb) return; // Skip on web
    
    try {
      final dbService = DatabaseService();
      final committees = dbService.getHostedCommittees(hostId);
      
      if (committees.isEmpty) {
        await _setWidgetData(
          committeeName: 'No Kametis',
          memberName: 'Create a committee to see payout info',
          payoutDate: '',
          amount: '',
        );
        return;
      }
      
      // Get first committee's next payout
      final committee = committees.first;
      final members = dbService.getMembersByCommittee(committee.id);
      
      if (members.isEmpty) {
        await _setWidgetData(
          committeeName: committee.name,
          memberName: 'No members added',
          payoutDate: '',
          amount: '',
        );
        return;
      }
      
      // Find next member to receive payout (first member who hasn't received payout)
      final sortedMembers = members..sort((a, b) => 
          (a.payoutOrder ?? 999).compareTo(b.payoutOrder ?? 999));
      
      Member? nextPayoutMember;
      for (final member in sortedMembers) {
        if (!member.hasReceivedPayout) {
          nextPayoutMember = member;
          break;
        }
      }
      
      if (nextPayoutMember == null) {
        await _setWidgetData(
          committeeName: committee.name,
          memberName: 'All payouts complete!',
          payoutDate: '',
          amount: '',
        );
        return;
      }
      
      // Calculate next payout date
      final payoutDate = _calculateNextPayoutDate(committee, nextPayoutMember.payoutOrder ?? 0);
      final amount = 'PKR ${(committee.contributionAmount * members.length).toInt()}';
      
      await _setWidgetData(
        committeeName: committee.name,
        memberName: nextPayoutMember.name,
        payoutDate: payoutDate != null ? DateFormat('MMM d, yyyy').format(payoutDate) : '',
        amount: amount,
      );
      
    } catch (e) {
      debugPrint('Widget update error: $e');
    }
  }
  
  static DateTime? _calculateNextPayoutDate(Committee committee, int payoutOrder) {
    if (payoutOrder <= 0) return null;
    
    final startDate = committee.startDate;
    final payoutInterval = committee.paymentIntervalDays;
    
    // Payout date is startDate + (payoutOrder - 1) * payoutInterval
    return startDate.add(Duration(days: (payoutOrder - 1) * payoutInterval));
  }
  
  static Future<void> _setWidgetData({
    required String committeeName,
    required String memberName,
    required String payoutDate,
    required String amount,
  }) async {
    await HomeWidget.saveWidgetData<String>('committee_name', committeeName);
    await HomeWidget.saveWidgetData<String>('next_payout_member', memberName);
    await HomeWidget.saveWidgetData<String>('next_payout_date', payoutDate);
    await HomeWidget.saveWidgetData<String>('payout_amount', amount);
    
    // Trigger widget update
    await HomeWidget.updateWidget(
      name: 'CommitteeWidget',
      androidName: 'CommitteeWidget',
      qualifiedAndroidName: 'com.kameti.app.CommitteeWidget',
    );
  }
  
  /// Initialize widget when app starts
  static Future<void> initialize() async {
    if (kIsWeb) return;
    
    // Set app group ID for iOS (not needed for Android but doesn't hurt)
    await HomeWidget.setAppGroupId(_appGroupId);
  }
}
