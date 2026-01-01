import 'package:flutter/material.dart';
import 'package:firebase_core/firebase_core.dart';
import 'firebase_options.dart';
import 'services/database_service.dart';
import 'utils/app_theme.dart';
import 'screens/splash_screen.dart';
import 'screens/lock_screen.dart';
import 'services/biometric_service.dart';
import 'services/auth_service.dart';

final GlobalKey<NavigatorState> navigatorKey = GlobalKey<NavigatorState>();

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  // Initialize Hive
  await DatabaseService.initialize();

  // Initialize Firebase
  await Firebase.initializeApp(
    options: DefaultFirebaseOptions.currentPlatform,
  );

  runApp(const CommitteeApp());
}

class CommitteeApp extends StatefulWidget {
  const CommitteeApp({super.key});

  @override
  State<CommitteeApp> createState() => _CommitteeAppState();
}

class _CommitteeAppState extends State<CommitteeApp> with WidgetsBindingObserver {
  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    if (state == AppLifecycleState.resumed) {
      _checkBiometricLock();
    }
  }

  Future<void> _checkBiometricLock() async {
    // Check if lock is enabled
    final isEnabled = await BiometricService.isBiometricLockEnabled();
    // Only lock if enabled and not already showing lock screen
    if (isEnabled && !LockScreen.isShown) {
      final user = AuthService().currentUser;
      final isRealHost = user != null && !user.isAnonymous && user.email != null;
      
      navigatorKey.currentState?.pushReplacement(
        PageRouteBuilder(
          pageBuilder: (_, __, ___) => LockScreen(isHost: isRealHost),
          transitionDuration: Duration.zero,
        ),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      navigatorKey: navigatorKey,
      title: 'Kameti',
      debugShowCheckedModeBanner: false,
      theme: AppTheme.lightTheme,
      darkTheme: AppTheme.darkTheme,
      themeMode: ThemeMode.dark,
      home: const SplashScreen(),
    );
  }
}
