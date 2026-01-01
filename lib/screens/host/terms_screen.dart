import 'package:flutter/material.dart';
import '../../utils/app_theme.dart';

class TermsScreen extends StatelessWidget {
  const TermsScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.darkBg,
      appBar: AppBar(
        title: const Text('Terms & Conditions'),
        backgroundColor: AppTheme.darkCard,
      ),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          _buildHeader(),
          const SizedBox(height: 20),
          _buildSection(
            '1. Acceptance of Terms',
            'By downloading, installing, or using Kameti - Committee Manager ("the App"), '
            'you agree to be bound by these Terms and Conditions. If you do not agree to these terms, '
            'please do not use the App.',
          ),
          _buildSection(
            '2. Description of Service',
            'Kameti is a mobile application designed to help users manage rotating savings committees (ROSCA). '
            'The app allows hosts to create committees, add members, track payments, and share committee information.',
          ),
          _buildSection(
            '3. User Accounts',
            '• You must be at least 13 years old to use this App\n'
            '• You are responsible for maintaining the security of your account\n'
            '• You must provide accurate information during registration\n'
            '• You are responsible for all activities under your account',
          ),
          _buildSection(
            '4. User Responsibilities',
            '• Use the App only for lawful purposes\n'
            '• Do not attempt to gain unauthorized access\n'
            '• Do not upload harmful or malicious content\n'
            '• Respect the privacy of other users\n'
            '• Keep your login credentials secure',
          ),
          _buildSection(
            '5. Data and Privacy',
            '• You own your committee data\n'
            '• We store your data securely in the cloud\n'
            '• We do not sell or share your data with third parties\n'
            '• Please review our Privacy Policy for more details',
          ),
          _buildSection(
            '6. Intellectual Property',
            'The App, including its design, features, and content, is the intellectual property of the developer. '
            'You may not copy, modify, distribute, or create derivative works without permission.',
          ),
          _buildSection(
            '7. Disclaimer of Warranties',
            'The App is provided "AS IS" without warranties of any kind. We do not guarantee:\n'
            '• Uninterrupted or error-free service\n'
            '• Accuracy of calculations or data\n'
            '• Security against all possible threats\n\n'
            'You use the App at your own risk.',
          ),
          _buildSection(
            '8. Limitation of Liability',
            'We are not liable for:\n'
            '• Loss of data or information\n'
            '• Financial losses from using the App\n'
            '• Disputes between committee members\n'
            '• Any indirect or consequential damages',
          ),
          _buildSection(
            '9. Termination',
            '• You may stop using the App at any time\n'
            '• You may delete your account from the Profile section\n'
            '• We may terminate accounts for violations of these terms\n'
            '• Upon termination, your data will be deleted within 30 days',
          ),
          _buildSection(
            '10. Changes to Terms',
            'We may update these Terms and Conditions from time to time. Continued use of the App '
            'after changes constitutes acceptance of the new terms. We will notify users of significant changes.',
          ),
          _buildSection(
            '11. Governing Law',
            'These Terms shall be governed by and construed in accordance with the laws of Pakistan. '
            'Any disputes shall be resolved through appropriate legal channels.',
          ),
          _buildSection(
            '12. Contact Information',
            'For questions about these Terms, contact us at:\n\n'
            'Email: msaad.official6@gmail.com\n'
            'Phone: +92 321 8685488',
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
          Icon(Icons.article, color: Colors.white, size: 40),
          SizedBox(width: 16),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Terms & Conditions',
                  style: TextStyle(
                    color: Colors.white,
                    fontSize: 20,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                Text(
                  'Please read carefully',
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
