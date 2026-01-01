import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';
import '../utils/app_theme.dart';
import '../services/biometric_service.dart';
import 'home_screen.dart';
import 'host/host_dashboard_screen.dart';


class LockScreen extends StatefulWidget {
  final bool isHost;
  static bool isShown = false;
  
  const LockScreen({super.key, required this.isHost});

  @override
  State<LockScreen> createState() => _LockScreenState();
}

class _LockScreenState extends State<LockScreen> with WidgetsBindingObserver {
  bool _isAuthenticating = false;
  String _error = '';
  String _biometricType = 'Biometric';

  @override
  void initState() {
    super.initState();
    LockScreen.isShown = true;
    WidgetsBinding.instance.addObserver(this);
    _initBiometricType();
    _authenticate();
  }

  @override
  void dispose() {
    LockScreen.isShown = false;
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    // Auto-authenticate when app resumes
    if (state == AppLifecycleState.resumed && !_isAuthenticating) {
      _authenticate();
    }
  }

  Future<void> _initBiometricType() async {
    final types = await BiometricService.getAvailableBiometrics();
    if (mounted) {
      setState(() {
        _biometricType = BiometricService.getBiometricTypeName(types);
      });
    }
  }

  Future<void> _authenticate() async {
    if (_isAuthenticating) return;
    
    setState(() {
      _isAuthenticating = true;
      _error = '';
    });

    try {
      final success = await BiometricService.authenticate(
        reason: 'Unlock Committee App',
      );

      if (success && mounted) {
        _navigateToApp();
      } else if (mounted) {
        setState(() {
          _error = 'Authentication failed. Tap to try again.';
          _isAuthenticating = false;
        });
      }
    } catch (e) {
      if (mounted) {
        setState(() {
          _error = 'Error: $e';
          _isAuthenticating = false;
        });
      }
    }
  }

  void _navigateToApp() {
    Navigator.of(context).pushReplacement(
      MaterialPageRoute(
        builder: (_) => widget.isHost 
            ? const HostDashboardScreen() 
            : const HomeScreen(),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: AppTheme.darkBg,
      body: SafeArea(
        child: Center(
          child: Padding(
            padding: const EdgeInsets.all(32),
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                // App Icon
                Container(
                  width: 100,
                  height: 100,
                  decoration: BoxDecoration(
                    gradient: const LinearGradient(
                      colors: [AppTheme.primaryColor, AppTheme.primaryDark],
                    ),
                    borderRadius: BorderRadius.circular(24),
                    boxShadow: [
                      BoxShadow(
                        color: AppTheme.primaryColor.withAlpha(80),
                        blurRadius: 20,
                        spreadRadius: 2,
                      ),
                    ],
                  ),
                  child: const Icon(
                    Icons.account_balance_wallet_rounded,
                    size: 50,
                    color: Colors.white,
                  ),
                ),
                const SizedBox(height: 32),
                
                // Title
                Text(
                  'Kameti',
                  style: GoogleFonts.inter(
                    fontSize: 28,
                    fontWeight: FontWeight.bold,
                    color: Colors.white,
                  ),
                ),
                const SizedBox(height: 8),
                Text(
                  'Locked',
                  style: GoogleFonts.inter(
                    fontSize: 16,
                    color: Colors.grey[500],
                  ),
                ),
                const SizedBox(height: 60),
                
                // Biometric Button
                GestureDetector(
                  onTap: _isAuthenticating ? null : _authenticate,
                  child: Container(
                    width: 80,
                    height: 80,
                    decoration: BoxDecoration(
                      color: AppTheme.darkCard,
                      shape: BoxShape.circle,
                      border: Border.all(
                        color: _isAuthenticating 
                            ? AppTheme.primaryColor 
                            : Colors.grey[700]!,
                        width: 2,
                      ),
                    ),
                    child: _isAuthenticating
                        ? const Padding(
                            padding: EdgeInsets.all(20),
                            child: CircularProgressIndicator(
                              color: AppTheme.primaryColor,
                              strokeWidth: 3,
                            ),
                          )
                        : Icon(
                            _biometricType == 'Face ID' 
                                ? Icons.face_rounded 
                                : Icons.fingerprint_rounded,
                            size: 40,
                            color: AppTheme.primaryColor,
                          ),
                  ),
                ),
                const SizedBox(height: 16),
                
                Text(
                  _isAuthenticating 
                      ? 'Authenticating...' 
                      : 'Tap to unlock with $_biometricType',
                  style: GoogleFonts.inter(
                    fontSize: 14,
                    color: Colors.grey[400],
                  ),
                ),
                
                // Error Message
                if (_error.isNotEmpty) ...[
                  const SizedBox(height: 24),
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
                    decoration: BoxDecoration(
                      color: Colors.red.withAlpha(30),
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(color: Colors.red.withAlpha(50)),
                    ),
                    child: Text(
                      _error,
                      style: GoogleFonts.inter(
                        fontSize: 13,
                        color: Colors.red[300],
                      ),
                      textAlign: TextAlign.center,
                    ),
                  ),
                ],
              ],
            ),
          ),
        ),
      ),
    );
  }
}
