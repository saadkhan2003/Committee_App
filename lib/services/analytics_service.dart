import 'package:firebase_analytics/firebase_analytics.dart';

class AnalyticsService {
  static final FirebaseAnalytics _analytics = FirebaseAnalytics.instance;
  
  // Get analytics observer for navigation
  static FirebaseAnalyticsObserver get observer => 
      FirebaseAnalyticsObserver(analytics: _analytics);

  // Screen Views
  static Future<void> logScreenView(String screenName) async {
    await _analytics.logScreenView(screenName: screenName);
  }

  // User Events
  static Future<void> logLogin() async {
    await _analytics.logLogin(loginMethod: 'email');
  }

  static Future<void> logSignUp() async {
    await _analytics.logSignUp(signUpMethod: 'email');
  }

  static Future<void> logLogout() async {
    await _analytics.logEvent(name: 'logout');
  }

  // Committee Events
  static Future<void> logCommitteeCreated({
    required String committeeName,
    required int memberCount,
    required double contributionAmount,
  }) async {
    await _analytics.logEvent(
      name: 'committee_created',
      parameters: {
        'name': committeeName,
        'member_count': memberCount,
        'contribution_amount': contributionAmount,
      },
    );
  }

  static Future<void> logCommitteeDeleted() async {
    await _analytics.logEvent(name: 'committee_deleted');
  }

  // Member Events
  static Future<void> logMemberAdded() async {
    await _analytics.logEvent(name: 'member_added');
  }

  static Future<void> logMemberDeleted() async {
    await _analytics.logEvent(name: 'member_deleted');
  }

  // Payment Events
  static Future<void> logPaymentMarked({
    required double amount,
    required bool isPaid,
  }) async {
    await _analytics.logEvent(
      name: 'payment_marked',
      parameters: {
        'amount': amount,
        'status': isPaid ? 'paid' : 'unpaid',
      },
    );
  }

  static Future<void> logPayoutReceived({required double amount}) async {
    await _analytics.logEvent(
      name: 'payout_received',
      parameters: {'amount': amount},
    );
  }

  // Share Events
  static Future<void> logShare({required String contentType}) async {
    await _analytics.logShare(
      contentType: contentType,
      itemId: 'committee_info',
      method: 'native_share',
    );
  }

  // Viewer Events
  static Future<void> logViewerJoined() async {
    await _analytics.logEvent(name: 'viewer_joined_committee');
  }

  // Password Events
  static Future<void> logPasswordReset() async {
    await _analytics.logEvent(name: 'password_reset_requested');
  }
}
