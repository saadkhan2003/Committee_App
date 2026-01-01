import 'package:flutter/material.dart';
import '../../utils/app_theme.dart';

class PrivacyPolicyScreen extends StatelessWidget {
  const PrivacyPolicyScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.darkBg,
      appBar: AppBar(
        title: const Text('Privacy Policy'),
        backgroundColor: AppTheme.darkCard,
      ),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          _buildHeader(),
          const SizedBox(height: 20),
          _buildSection(
            'Introduction',
            'Kameti - Committee Manager ("we", "our", or "us") respects your privacy. '
            'This Privacy Policy explains how we collect, use, and protect your information '
            'when you use our mobile application.',
          ),
          _buildSection(
            'Information We Collect',
            '• Email address (for authentication)\n'
            '• Display name (optional)\n'
            '• Committee data you create\n'
            '• Member names and payment records\n'
            '• Payment dates and amounts\n'
            '• Device type and app version',
          ),
          _buildSection(
            'How We Use Your Information',
            '• Provide and maintain the app service\n'
            '• Sync your data across devices\n'
            '• Send important notifications\n'
            '• Improve our app and user experience',
          ),
          _buildSection(
            'Data Storage',
            'Your data is stored securely in Firebase Cloud Firestore (Google\'s secure cloud database) '
            'and local device storage for offline access. All data transmission uses encrypted HTTPS connections.',
          ),
          _buildSection(
            'Data Sharing',
            'We do NOT:\n'
            '• Sell your personal information\n'
            '• Share your data with third parties\n'
            '• Use your data for advertising\n'
            '• Track your location',
          ),
          _buildSection(
            'Your Rights',
            'You have the right to:\n'
            '• Access your personal data\n'
            '• Delete your account and all data\n'
            '• Export your committee data\n'
            '• Opt out of notifications',
          ),
          _buildSection(
            'Data Deletion',
            'To delete your data:\n'
            '1. Go to Profile in the app\n'
            '2. Select "Delete Account"\n'
            '3. Confirm deletion\n\n'
            'All your data will be permanently removed within 30 days.',
          ),
          _buildSection(
            'Third-Party Services',
            'We use the following Firebase services:\n'
            '• Firebase Auth (authentication)\n'
            '• Firebase Firestore (database)\n'
            '• Firebase Analytics (usage statistics)\n\n'
            'These services have their own privacy policies governed by Google.',
          ),
          _buildSection(
            'Children\'s Privacy',
            'Our app is not intended for children under 13. We do not knowingly collect data from children.',
          ),
          _buildSection(
            'Contact Us',
            'If you have questions about this Privacy Policy, contact us at:\n\n'
            'Email: msaad.official6@gmail.com',
          ),
          const SizedBox(height: 16),
          Center(
            child: Text(
              'Last updated: December 2025',
              style: TextStyle(color: Colors.grey[600], fontSize: 12),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildHeader() {
    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        gradient: const LinearGradient(
          colors: [AppTheme.primaryColor, AppTheme.primaryDark],
        ),
        borderRadius: BorderRadius.circular(16),
      ),
      child: const Row(
        children: [
          Icon(Icons.privacy_tip, color: Colors.white, size: 40),
          SizedBox(width: 16),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Privacy Policy',
                  style: TextStyle(
                    color: Colors.white,
                    fontSize: 20,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                Text(
                  'Your privacy matters to us',
                  style: TextStyle(color: Colors.white70, fontSize: 14),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildSection(String title, String content) {
    return Container(
      margin: const EdgeInsets.only(bottom: 16),
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: AppTheme.darkCard,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: Colors.white.withAlpha(13)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            title,
            style: const TextStyle(
              color: Colors.white,
              fontSize: 16,
              fontWeight: FontWeight.bold,
            ),
          ),
          const SizedBox(height: 10),
          Text(
            content,
            style: TextStyle(color: Colors.grey[300], height: 1.6),
          ),
        ],
      ),
    );
  }
}
