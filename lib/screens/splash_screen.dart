import 'package:flutter/material.dart';
import '../utils/app_theme.dart';
import '../services/auth_service.dart';
import '../services/sync_service.dart';
import '../services/update_service.dart';
import '../screens/home_screen.dart';
import '../screens/host/host_dashboard_screen.dart';

class SplashScreen extends StatefulWidget {
  const SplashScreen({super.key});

  @override
  State<SplashScreen> createState() => _SplashScreenState();
}

class _SplashScreenState extends State<SplashScreen> with SingleTickerProviderStateMixin {
  late AnimationController _controller;
  late Animation<double> _fadeAnimation;
  late Animation<double> _scaleAnimation;
  
  String _status = 'Initializing...';
  bool _isComplete = false;

  @override
  void initState() {
    super.initState();
    _controller = AnimationController(
      duration: const Duration(milliseconds: 1500),
      vsync: this,
    );
    
    _fadeAnimation = Tween<double>(begin: 0.0, end: 1.0).animate(
      CurvedAnimation(parent: _controller, curve: Curves.easeIn),
    );
    
    _scaleAnimation = Tween<double>(begin: 0.8, end: 1.0).animate(
      CurvedAnimation(parent: _controller, curve: Curves.elasticOut),
    );
    
    _controller.forward();
    _initializeApp();
  }

  Future<void> _initializeApp() async {
    try {
      // Step 1: Database
      setState(() => _status = 'Loading database...');
      await Future.delayed(const Duration(milliseconds: 300));
      
      // Step 2: Check auth
      setState(() => _status = 'Checking authentication...');
      await Future.delayed(const Duration(milliseconds: 300));
      
      final authService = AuthService();
      final user = authService.currentUser;
      
      // Step 3: Sync if logged in AND is a real host (not anonymous)
      final isRealHost = user != null && !user.isAnonymous && user.email != null;
      
      if (isRealHost) {
        setState(() => _status = 'Syncing data...');
        try {
          final syncService = SyncService();
          await syncService.syncAll(user.uid).timeout(
            const Duration(seconds: 5),
            onTimeout: () => SyncResult(success: true, message: 'Timeout'),
          );
        } catch (e) {
          debugPrint('Sync error: $e');
        }
      }
      
      // Step 4: Ready
      setState(() {
        _status = 'Ready!';
        _isComplete = true;
      });
      
      await Future.delayed(const Duration(milliseconds: 500));
      
      if (mounted) {
        // Navigate: Real hosts go to Dashboard, everyone else (including anonymous) goes to HomeScreen
        Navigator.of(context).pushReplacement(
          PageRouteBuilder(
            pageBuilder: (context, animation, secondaryAnimation) =>
                isRealHost ? const HostDashboardScreen() : const HomeScreen(),
            transitionsBuilder: (context, animation, secondaryAnimation, child) {
              return FadeTransition(opacity: animation, child: child);
            },
            transitionDuration: const Duration(milliseconds: 500),
          ),
        );
        
        // Check for updates after navigation (for real hosts only)
        if (isRealHost) {
          Future.delayed(const Duration(seconds: 2), () {
            if (mounted) {
              UpdateService.checkForUpdate(context);
            }
          });
        }
      }
    } catch (e) {
      debugPrint('Splash init error: $e');
      setState(() => _status = 'Error: $e');
    }
  }

  @override
  void dispose() {
    _controller.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.darkBg,
      body: Center(
        child: AnimatedBuilder(
          animation: _controller,
          builder: (context, child) {
            return Opacity(
              opacity: _fadeAnimation.value,
              child: Transform.scale(
                scale: _scaleAnimation.value,
                child: Column(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    // App Icon/Logo
                    Container(
                      width: 120,
                      height: 120,
                      decoration: BoxDecoration(
                        gradient: LinearGradient(
                          colors: [
                            AppTheme.primaryColor,
                            AppTheme.primaryColor.withValues(alpha: 0.7),
                          ],
                          begin: Alignment.topLeft,
                          end: Alignment.bottomRight,
                        ),
                        borderRadius: BorderRadius.circular(28),
                        boxShadow: [
                          BoxShadow(
                            color: AppTheme.primaryColor.withValues(alpha: 0.4),
                            blurRadius: 30,
                            spreadRadius: 5,
                          ),
                        ],
                      ),
                      child: const Icon(
                        Icons.account_balance_wallet_rounded,
                        size: 60,
                        color: Colors.white,
                      ),
                    ),
                    const SizedBox(height: 32),
                    
                    // App Name
                    const Text(
                      'Committee',
                      style: TextStyle(
                        fontSize: 36,
                        fontWeight: FontWeight.bold,
                        color: Colors.white,
                        letterSpacing: 2,
                      ),
                    ),
                    const SizedBox(height: 8),
                    Text(
                      'Payment Tracker',
                      style: TextStyle(
                        fontSize: 16,
                        color: Colors.grey[400],
                        letterSpacing: 4,
                      ),
                    ),
                    const SizedBox(height: 60),
                    
                    // Loading indicator
                    if (!_isComplete)
                      SizedBox(
                        width: 200,
                        child: Column(
                          children: [
                            ClipRRect(
                              borderRadius: BorderRadius.circular(4),
                              child: LinearProgressIndicator(
                                backgroundColor: Colors.grey[800],
                                valueColor: const AlwaysStoppedAnimation<Color>(
                                  AppTheme.primaryColor,
                                ),
                                minHeight: 4,
                              ),
                            ),
                            const SizedBox(height: 16),
                            Text(
                              _status,
                              style: TextStyle(
                                fontSize: 12,
                                color: Colors.grey[500],
                              ),
                            ),
                          ],
                        ),
                      )
                    else
                      const Icon(
                        Icons.check_circle_rounded,
                        size: 32,
                        color: AppTheme.primaryColor,
                      ),
                  ],
                ),
              ),
            );
          },
        ),
      ),
    );
  }
}
