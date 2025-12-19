import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:ota_update/ota_update.dart';
import 'package:http/http.dart' as http;
import 'package:package_info_plus/package_info_plus.dart';
import '../utils/app_theme.dart';

class UpdateService {
  // GitHub Releases URL for version.json
  static const String _versionCheckUrl = 
      'https://github.com/saadkhan2003/Committee_app_personal/releases/latest/download/version.json';
  
  // version.json format:
  // {
  //   "version": "1.0.1",
  //   "apkUrl": "https://your-server.com/committee-app/app-release.apk",
  //   "releaseNotes": "Bug fixes and improvements"
  // }

  static Future<void> checkForUpdate(BuildContext context) async {
    // Only check on Android (not web)
    if (kIsWeb) return;

    try {
      final packageInfo = await PackageInfo.fromPlatform();
      final currentVersion = packageInfo.version;

      // Fetch version info from server
      final response = await http.get(Uri.parse(_versionCheckUrl));
      if (response.statusCode != 200) return;

      final versionData = json.decode(response.body);
      final latestVersion = versionData['version'] as String;
      final apkUrl = versionData['apkUrl'] as String;
      final releaseNotes = versionData['releaseNotes'] as String? ?? 'Bug fixes';

      // Compare versions
      if (_isNewerVersion(currentVersion, latestVersion)) {
        if (context.mounted) {
          _showUpdateDialog(context, latestVersion, releaseNotes, apkUrl);
        }
      }
    } catch (e) {
      // Silently fail - don't interrupt user if update check fails
      debugPrint('Update check failed: $e');
    }
  }

  static bool _isNewerVersion(String current, String latest) {
    final currentParts = current.split('.').map(int.parse).toList();
    final latestParts = latest.split('.').map(int.parse).toList();

    for (int i = 0; i < 3; i++) {
      final currentPart = i < currentParts.length ? currentParts[i] : 0;
      final latestPart = i < latestParts.length ? latestParts[i] : 0;
      
      if (latestPart > currentPart) return true;
      if (latestPart < currentPart) return false;
    }
    return false;
  }

  static void _showUpdateDialog(
    BuildContext context,
    String version,
    String releaseNotes,
    String apkUrl,
  ) {
    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (context) => AlertDialog(
        backgroundColor: AppTheme.darkCard,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        title: Row(
          children: [
            Container(
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: AppTheme.primaryColor.withAlpha(30),
                borderRadius: BorderRadius.circular(8),
              ),
              child: const Icon(
                Icons.system_update_rounded,
                color: AppTheme.primaryColor,
              ),
            ),
            const SizedBox(width: 12),
            const Text('Update Available', style: TextStyle(color: Colors.white)),
          ],
        ),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              'Version $version is available',
              style: const TextStyle(color: Colors.white, fontWeight: FontWeight.w600),
            ),
            const SizedBox(height: 8),
            Text(
              releaseNotes,
              style: TextStyle(color: Colors.grey[400], fontSize: 13),
            ),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: const Text('Later'),
          ),
          ElevatedButton.icon(
            onPressed: () {
              Navigator.pop(context);
              _downloadAndInstall(context, apkUrl);
            },
            icon: const Icon(Icons.download_rounded, size: 18),
            label: const Text('Update Now'),
          ),
        ],
      ),
    );
  }

  static void _downloadAndInstall(BuildContext context, String apkUrl) {
    double progress = 0;
    
    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (context) => StatefulBuilder(
        builder: (context, setState) => AlertDialog(
          backgroundColor: AppTheme.darkCard,
          content: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              const CircularProgressIndicator(),
              const SizedBox(height: 16),
              Text(
                'Downloading update... ${(progress * 100).toInt()}%',
                style: const TextStyle(color: Colors.white),
              ),
              const SizedBox(height: 8),
              LinearProgressIndicator(value: progress),
            ],
          ),
        ),
      ),
    );

    try {
      OtaUpdate().execute(apkUrl).listen(
        (event) {
          if (event.status == OtaStatus.DOWNLOADING) {
            // Update progress (this is a simplified approach)
            debugPrint('Download progress: ${event.value}%');
          } else if (event.status == OtaStatus.INSTALLING) {
            // APK is being installed
            debugPrint('Installing...');
          }
        },
        onError: (e) {
          if (context.mounted) {
            Navigator.pop(context);
            ScaffoldMessenger.of(context).showSnackBar(
              SnackBar(
                content: Text('Update failed: $e'),
                backgroundColor: AppTheme.errorColor,
              ),
            );
          }
        },
      );
    } catch (e) {
      if (context.mounted) {
        Navigator.pop(context);
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('Update failed: $e'),
            backgroundColor: AppTheme.errorColor,
          ),
        );
      }
    }
  }
}
