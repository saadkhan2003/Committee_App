import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:ota_update/ota_update.dart';
import 'package:http/http.dart' as http;
import 'package:package_info_plus/package_info_plus.dart';
import 'package:url_launcher/url_launcher.dart';
import '../utils/app_theme.dart';

class UpdateService {
  // GitHub Releases URL for version.json
  static const String _versionCheckUrl = 
      'https://github.com/saadkhan2003/Committee_app_personal/releases/latest/download/version.json';
  
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
    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (dialogContext) => AlertDialog(
        backgroundColor: AppTheme.darkCard,
        content: const Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            CircularProgressIndicator(),
            SizedBox(height: 16),
            Text(
              'Downloading update...',
              style: TextStyle(color: Colors.white),
            ),
            SizedBox(height: 8),
            Text(
              'Please wait, the installer will open automatically.',
              style: TextStyle(color: Colors.grey, fontSize: 12),
              textAlign: TextAlign.center,
            ),
          ],
        ),
      ),
    );

    try {
      OtaUpdate().execute(apkUrl).listen(
        (event) {
          debugPrint('OTA Status: ${event.status}, Value: ${event.value}');
          if (event.status == OtaStatus.INSTALLING) {
            if (context.mounted) {
              Navigator.of(context, rootNavigator: true).pop();
            }
          }
        },
        onError: (e) {
          debugPrint('OTA Error: $e');
          if (context.mounted) {
            Navigator.of(context, rootNavigator: true).pop();
            _showFallbackDialog(context, apkUrl, e.toString());
          }
        },
      );
    } catch (e) {
      debugPrint('OTA Exception: $e');
      if (context.mounted) {
        Navigator.of(context, rootNavigator: true).pop();
        _showFallbackDialog(context, apkUrl, e.toString());
      }
    }
  }

  static void _showFallbackDialog(BuildContext context, String apkUrl, String error) {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        backgroundColor: AppTheme.darkCard,
        title: const Text('Download Manually', style: TextStyle(color: Colors.white)),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            const Icon(Icons.download_rounded, size: 48, color: AppTheme.primaryColor),
            const SizedBox(height: 16),
            const Text(
              'Auto-update failed. Please download and install the update manually.',
              style: TextStyle(color: Colors.white),
              textAlign: TextAlign.center,
            ),
            const SizedBox(height: 8),
            Text(
              'Error: $error',
              style: TextStyle(color: Colors.grey[600], fontSize: 10),
              textAlign: TextAlign.center,
            ),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: const Text('Cancel'),
          ),
          ElevatedButton.icon(
            onPressed: () async {
              Navigator.pop(context);
              final uri = Uri.parse(apkUrl);
              if (await canLaunchUrl(uri)) {
                await launchUrl(uri, mode: LaunchMode.externalApplication);
              }
            },
            icon: const Icon(Icons.open_in_browser, size: 18),
            label: const Text('Download'),
          ),
        ],
      ),
    );
  }
}
