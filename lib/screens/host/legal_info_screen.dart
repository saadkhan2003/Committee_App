import 'package:flutter/material.dart';
import 'package:package_info_plus/package_info_plus.dart';
import '../../utils/app_theme.dart';

class LegalInfoScreen extends StatefulWidget {
  const LegalInfoScreen({super.key});

  @override
  State<LegalInfoScreen> createState() => _LegalInfoScreenState();
}

class _LegalInfoScreenState extends State<LegalInfoScreen> {
  String _version = '';

  @override
  void initState() {
    super.initState();
    _loadVersion();
  }

  Future<void> _loadVersion() async {
    final info = await PackageInfo.fromPlatform();
    if (mounted) {
      setState(() => _version = info.version);
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.darkBg,
      appBar: AppBar(
        title: const Text('About & Legal'),
        backgroundColor: AppTheme.darkCard,
      ),
      body: ListView(
        padding: const EdgeInsets.all(16),
        children: [
          // App Info Card
          _buildCard(
            icon: Icons.info_outline,
            title: 'App Information',
            children: [
              _buildInfoRow('App Name', 'Kameti - Committee Manager'),
              _buildInfoRow('Version', _version.isEmpty ? 'Loading...' : _version),
              _buildInfoRow('Category', 'Finance'),
              _buildInfoRow('Developer', 'Saad Khan'),
            ],
          ),
          const SizedBox(height: 16),

          // Legal Section
          _buildCard(
            icon: Icons.gavel,
            title: 'Legal',
            children: [
              _buildLinkTile(
                icon: Icons.privacy_tip_outlined,
                title: 'Privacy Policy',
                onTap: () => _showLegalPage(context, 'Privacy Policy', _privacyPolicy),
              ),
              _buildLinkTile(
                icon: Icons.article_outlined,
                title: 'Terms of Service',
                onTap: () => _showLegalPage(context, 'Terms of Service', _termsOfService),
              ),
              _buildLinkTile(
                icon: Icons.security_outlined,
                title: 'Data Safety',
                onTap: () => _showLegalPage(context, 'Data Safety', _dataSafety),
              ),
            ],
          ),
          const SizedBox(height: 16),

          // Credits
          _buildCard(
            icon: Icons.favorite_outline,
            title: 'Credits',
            children: [
              _buildInfoRow('Framework', 'Flutter'),
              _buildInfoRow('Backend', 'Firebase'),
              _buildInfoRow('Made with', '❤️ in Pakistan'),
            ],
          ),
          const SizedBox(height: 24),

          // Copyright
          Center(
            child: Text(
              '© 2024 Committee App. All rights reserved.',
              style: TextStyle(color: Colors.grey[600], fontSize: 12),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildCard({
    required IconData icon,
    required String title,
    required List<Widget> children,
  }) {
    return Container(
      decoration: BoxDecoration(
        color: AppTheme.darkCard,
        borderRadius: BorderRadius.circular(16),
        border: Border.all(color: Colors.white.withAlpha(13)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Padding(
            padding: const EdgeInsets.all(16),
            child: Row(
              children: [
                Icon(icon, color: AppTheme.primaryColor, size: 24),
                const SizedBox(width: 12),
                Text(
                  title,
                  style: const TextStyle(
                    color: Colors.white,
                    fontSize: 18,
                    fontWeight: FontWeight.bold,
                  ),
                ),
              ],
            ),
          ),
          const Divider(height: 1, color: Colors.white10),
          ...children,
        ],
      ),
    );
  }

  Widget _buildInfoRow(String label, String value) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: TextStyle(color: Colors.grey[400])),
          Text(value, style: const TextStyle(color: Colors.white)),
        ],
      ),
    );
  }

  Widget _buildLinkTile({
    required IconData icon,
    required String title,
    required VoidCallback onTap,
  }) {
    return InkWell(
      onTap: onTap,
      child: Padding(
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
        child: Row(
          children: [
            Icon(icon, color: Colors.grey[400], size: 20),
            const SizedBox(width: 12),
            Expanded(
              child: Text(title, style: const TextStyle(color: Colors.white)),
            ),
            Icon(Icons.chevron_right, color: Colors.grey[600]),
          ],
        ),
      ),
    );
  }

  void _showLegalPage(BuildContext context, String title, String content) {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (context) => Scaffold(
          backgroundColor: AppTheme.darkBg,
          appBar: AppBar(
            title: Text(title),
            backgroundColor: AppTheme.darkCard,
          ),
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(20),
            child: Text(
              content,
              style: TextStyle(color: Colors.grey[300], height: 1.6),
            ),
          ),
        ),
      ),
    );
  }
}

// Legal Content
const String _privacyPolicy = '''
Privacy Policy

Last updated: December 2025

Introduction
Committee - Payment Tracker respects your privacy. This Privacy Policy explains how we collect, use, and protect your information.

Information We Collect
• Email address (for authentication)
• Display name (optional)
• Committee data you create
• Payment records

How We Use Your Information
• Provide and maintain the app
• Sync your data across devices
• Improve our service

Data Storage
Your data is stored securely in Firebase Cloud Firestore with encrypted connections.

Data Sharing
We do NOT sell, share, or use your data for advertising.

Your Rights
• Access your personal data
• Delete your account and data
• Export your committee data

Contact
For questions, contact the developer through the app store listing.
''';

const String _termsOfService = '''
Terms of Service

Last updated: December 2025

1. Acceptance of Terms
By using Committee - Payment Tracker, you agree to these terms.

2. Use of Service
• You must be 13 or older
• You are responsible for your account
• Do not misuse the service

3. User Content
• You own your committee data
• You grant us license to store and sync it
• We don't claim ownership of your data

4. Limitations
• Service provided "as is"
• We are not liable for data loss
• Backup your important data

5. Termination
• You can delete your account anytime
• We may terminate for violations

6. Changes
We may update these terms. Continued use means acceptance.
''';

const String _dataSafety = '''
Data Safety Information

Data Collection
✓ Email address - For account management
✓ Display name - For personalization
✓ Committee data - For app functionality

Data Security
✓ All data encrypted in transit (HTTPS)
✓ Secure Firebase authentication
✓ No third-party data sharing

Data Deletion
You can delete all your data:
1. Go to Profile
2. Select "Delete Account"
3. Confirm deletion

Your data will be permanently removed within 30 days.

Third-Party Services
• Firebase Auth (authentication)
• Firebase Firestore (database)
• Firebase Analytics (usage statistics)

We do NOT:
✗ Sell your personal information
✗ Share data with advertisers
✗ Track your location
''';
