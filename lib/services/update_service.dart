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
      'https://raw.githubusercontent.com/saadkhan2003/Committee_App/main/version.json';
  
  static Future<void> checkForUpdate(BuildContext context) async {
    // Only check on Android (not web)
    if (kIsWeb) {
      debugPrint('Update check: Skipping on web');
      return;
    }

    try {
      debugPrint('Update check: Starting...');
      final packageInfo = await PackageInfo.fromPlatform();
      final currentVersion = packageInfo.version;
      debugPrint('Update check: Current version = $currentVersion');

      // Fetch version info from server
      final response = await http.get(Uri.parse(_versionCheckUrl));
      debugPrint('Update check: Fetched version.json, status = ${response.statusCode}');
      if (response.statusCode != 200) return;

      final versionData = json.decode(response.body);
      final latestVersion = versionData['version'] as String;
      final apkUrl = versionData['apkUrl'] as String;
      final releaseNotes = versionData['releaseNotes'] as String? ?? 'Bug fixes';
      debugPrint('Update check: Latest version = $latestVersion');

      // Compare versions
      if (_isNewerVersion(currentVersion, latestVersion)) {
        debugPrint('Update check: Update available! Showing dialog...');
        if (context.mounted) {
          _showUpdateDialog(context, latestVersion, releaseNotes, apkUrl);
        }
      } else {
        debugPrint('Update check: Already up to date');
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
              _showDownloadProgress(context, apkUrl);
            },
            icon: const Icon(Icons.download_rounded, size: 18),
            label: const Text('Update Now'),
          ),
        ],
      ),
    );
  }

  static void _showDownloadProgress(BuildContext context, String apkUrl) {
    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (dialogContext) => _DownloadProgressDialog(
        apkUrl: apkUrl,
        onError: (error) => _showFallbackDialog(context, apkUrl, error),
      ),
    );
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

// Stateful widget for download progress
class _DownloadProgressDialog extends StatefulWidget {
  final String apkUrl;
  final Function(String) onError;

  const _DownloadProgressDialog({required this.apkUrl, required this.onError});

  @override
  State<_DownloadProgressDialog> createState() => _DownloadProgressDialogState();
}

class _DownloadProgressDialogState extends State<_DownloadProgressDialog> {
  double _progress = 0.0;
  String _status = 'Starting download...';

  @override
  void initState() {
    super.initState();
    _startDownload();
  }

  void _startDownload() {
    try {
      OtaUpdate().execute(widget.apkUrl).listen(
        (event) {
          debugPrint('OTA Status: ${event.status}, Value: ${event.value}');
          
          if (!mounted) return;
          
          switch (event.status) {
            case OtaStatus.DOWNLOADING:
              final progress = double.tryParse(event.value ?? '0') ?? 0;
              setState(() {
                _progress = progress / 100;
                _status = 'Downloading... ${progress.toInt()}%';
              });
              break;
            case OtaStatus.INSTALLING:
              setState(() {
                _progress = 1.0;
                _status = 'Installing...';
              });
              Future.delayed(const Duration(milliseconds: 500), () {
                if (mounted) Navigator.of(context).pop();
              });
              break;
            case OtaStatus.DOWNLOAD_ERROR:
              Navigator.of(context).pop();
              widget.onError(event.value ?? 'Download failed');
              break;
            default:
              break;
          }
        },
        onError: (e) {
          debugPrint('OTA Error: $e');
          if (mounted) {
            Navigator.of(context).pop();
            widget.onError(e.toString());
          }
        },
      );
    } catch (e) {
      debugPrint('OTA Exception: $e');
      if (mounted) {
        Navigator.of(context).pop();
        widget.onError(e.toString());
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      backgroundColor: AppTheme.darkCard,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          const Icon(
            Icons.download_rounded,
            size: 48,
            color: AppTheme.primaryColor,
          ),
          const SizedBox(height: 20),
          Text(
            _status,
            style: const TextStyle(
              color: Colors.white,
              fontSize: 16,
              fontWeight: FontWeight.w600,
            ),
          ),
          const SizedBox(height: 16),
          ClipRRect(
            borderRadius: BorderRadius.circular(8),
            child: LinearProgressIndicator(
              value: _progress > 0 ? _progress : null,
              backgroundColor: Colors.grey[800],
              valueColor: const AlwaysStoppedAnimation<Color>(AppTheme.primaryColor),
              minHeight: 8,
            ),
          ),
          const SizedBox(height: 16),
          Text(
            'The installer will open automatically',
            style: TextStyle(color: Colors.grey[600], fontSize: 11),
            textAlign: TextAlign.center,
          ),
        ],
      ),
    );
  }
}
