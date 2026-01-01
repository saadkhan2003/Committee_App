import 'package:flutter/widgets.dart';
import 'database_service.dart';

/// Simple localization service for English and Urdu
class LocalizationService extends ChangeNotifier {
  static final LocalizationService _instance = LocalizationService._internal();
  factory LocalizationService() => _instance;
  LocalizationService._internal();

  String _currentLanguage = 'en';
  String get currentLanguage => _currentLanguage;

  /// Initialize with stored language preference
  Future<void> initialize() async {
    final dbService = DatabaseService();
    _currentLanguage = await dbService.getLanguage();
    notifyListeners();
  }

  /// Set language and persist
  Future<void> setLanguage(String languageCode) async {
    if (_currentLanguage == languageCode) return;
    
    _currentLanguage = languageCode;
    final dbService = DatabaseService();
    await dbService.setLanguage(languageCode);
    notifyListeners();
  }

  /// Get translated string
  String translate(String key) {
    return _translations[_currentLanguage]?[key] ?? _translations['en']?[key] ?? key;
  }

  /// Shorthand for translate
  String t(String key) => translate(key);

  /// Available languages
  static final List<Map<String, String>> availableLanguages = [
    {'code': 'en', 'name': 'English', 'nativeName': 'English'},
    {'code': 'ur', 'name': 'Urdu', 'nativeName': 'اردو'},
  ];

  /// Translation strings
  static final Map<String, Map<String, String>> _translations = {
    'en': {
      // Common
      'app_name': 'Kameti',
      'cancel': 'Cancel',
      'save': 'Save',
      'done': 'Done',
      'delete': 'Delete',
      'edit': 'Edit',
      'close': 'Close',
      'ok': 'OK',
      'yes': 'Yes',
      'no': 'No',
      'loading': 'Loading...',
      'error': 'Error',
      'success': 'Success',
      'warning': 'Warning',
      'about': 'About',
      'terms_conditions': 'Terms & Conditions',
      'privacy_policy': 'Privacy Policy',
      'contact_us': 'Contact Us',
      'my_committees': 'My Kametis',
      'your_hosted_committees': 'Your Hosted Kametis',
      'archived': 'Archived',
      'no_committees_yet': 'No Kametis Yet',
      'create_first_committee_msg': 'Create your first committee to get started',
      'joined_committee_q': 'Joined a Kameti?',
      'view_committee_payments': 'View Kameti Payments',
      
      // Navigation
      'home': 'Home',
      'profile': 'Profile',
      'settings': 'Settings',
      'dashboard': 'Dashboard',
      
      // Auth
      'login': 'Login',
      'logout': 'Logout',
      'sign_up': 'Sign Up',
      'email': 'Email',
      'password': 'Password',
      'forgot_password': 'Forgot Password?',
      'google_sign_in': 'Sign in with Google',
      
      // Committee
      'committee': 'Kameti',
      'committees': 'Kametis',
      'create_committee': 'Create Kameti',
      'join_committee': 'Join Kameti',
      'committee_code': 'Kameti Code',
      'contribution_amount': 'Contribution Amount',
      'members': 'Members',
      'payout': 'Payout',
      'payment': 'Payment',
      'payments': 'Payments',
      'paid': 'Paid',
      'unpaid': 'Unpaid',
      'pending': 'Pending',
      'collected': 'Collected',
      
      // Payment Sheet
      'payment_sheet': 'Payment Sheet',
      'mark_payment': 'Mark Payment',
      'total_paid': 'Total Paid',
      'total_pending': 'Total Pending',
      
      // Profile
      'member_since': 'Member Since',
      'change_password': 'Change Password',
      'delete_account': 'Delete Account',
      'security_settings': 'Security Settings',
      'biometric_lock': 'Biometric Lock',
      'language': 'Language',
      
      // Settings
      'app_settings': 'App Settings',
      'select_language': 'Select Language',
    },
    
    'ur': {
      // Common
      'app_name': 'کمیٹی',
      'cancel': 'منسوخ',
      'save': 'محفوظ',
      'done': 'ہو گیا',
      'delete': 'حذف',
      'edit': 'ترمیم',
      'close': 'بند',
      'ok': 'ٹھیک ہے',
      'yes': 'ہاں',
      'no': 'نہیں',
      'loading': 'لوڈ ہو رہا ہے...',
      'error': 'غلطی',
      'success': 'کامیاب',
      'warning': 'انتباہ',
      
      // Navigation
      'home': 'ہوم',
      'profile': 'پروفائل',
      'settings': 'ترتیبات',
      'dashboard': 'ڈیش بورڈ',
      
      // Auth
      'login': 'لاگ ان',
      'logout': 'لاگ آؤٹ',
      'sign_up': 'سائن اپ',
      'email': 'ای میل',
      'password': 'پاس ورڈ',
      'forgot_password': 'پاس ورڈ بھول گئے؟',
      'google_sign_in': 'گوگل سے لاگ ان',
      
      // Committee
      'committee': 'کمیٹی',
      'committees': 'کمیٹیاں',
      'create_committee': 'کمیٹی بنائیں',
      'join_committee': 'کمیٹی میں شامل ہوں',
      'committee_code': 'کمیٹی کوڈ',
      'contribution_amount': 'چندہ رقم',
      'members': 'اراکین',
      'payout': 'ادائیگی',
      'payment': 'ادائیگی',
      'payments': 'ادائیگیاں',
      'paid': 'ادا شدہ',
      'unpaid': 'غیر ادا شدہ',
      'pending': 'زیر التواء',
      'collected': 'جمع شدہ',
      
      // Payment Sheet
      'payment_sheet': 'ادائیگی شیٹ',
      'mark_payment': 'ادائیگی لگائیں',
      'total_paid': 'کل ادا شدہ',
      'total_pending': 'کل زیر التواء',
      
      // Profile
      'member_since': 'رکنیت کی تاریخ',
      'change_password': 'پاس ورڈ تبدیل کریں',
      'delete_account': 'اکاؤنٹ حذف کریں',
      'security_settings': 'سیکورٹی ترتیبات',
      'biometric_lock': 'بائیو میٹرک لاک',
      'language': 'زبان',
      
      // Settings
      'app_settings': 'ایپ ترتیبات',
      'select_language': 'زبان منتخب کریں',
    },
  };
}

/// Extension to easily access translations
extension TranslationExtension on String {
  String get tr => LocalizationService().translate(this);
}
