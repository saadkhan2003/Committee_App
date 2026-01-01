import 'package:firebase_remote_config/firebase_remote_config.dart';
import 'package:package_info_plus/package_info_plus.dart';
import 'package:flutter/foundation.dart';

class RemoteConfigService {
  static final RemoteConfigService _instance = RemoteConfigService._internal();
  factory RemoteConfigService() => _instance;
  RemoteConfigService._internal();

  final FirebaseRemoteConfig _remoteConfig = FirebaseRemoteConfig.instance;

  Future<void> initialize() async {
    try {
      await _remoteConfig.setConfigSettings(RemoteConfigSettings(
        fetchTimeout: const Duration(minutes: 1),
        minimumFetchInterval: const Duration(hours: 1), // Cache for 1 hour
      ));
      
      // Default values
      await _remoteConfig.setDefaults({
        'min_supported_build_number': 0,
        'update_message': 'A critical update is available. Please update to continue using the app.',
        'play_store_url': 'https://play.google.com/store/apps/details?id=com.kameti.app',
      });

      await _remoteConfig.fetchAndActivate();
      debugPrint('Remote Config fetched. Min Build: ${_remoteConfig.getInt('min_supported_build_number')}');
    } catch (e) {
      debugPrint('Remote Config init error: $e');
    }
  }

  Future<bool> isUpdateRequired() async {
    try {
      final packageInfo = await PackageInfo.fromPlatform();
      final currentBuild = int.tryParse(packageInfo.buildNumber) ?? 0;
      final minBuild = _remoteConfig.getInt('min_supported_build_number');
      
      if (kDebugMode) {
        print('Update Check: Current=$currentBuild, Min=$minBuild');
      }
      
      return currentBuild < minBuild;
    } catch (e) {
      debugPrint('Update check error: $e');
      return false;
    }
  }
  
  String get updateMessage => _remoteConfig.getString('update_message');
  String get playStoreUrl => _remoteConfig.getString('play_store_url');
}
