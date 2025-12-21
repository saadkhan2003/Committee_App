import 'package:flutter/material.dart';
import '../utils/app_theme.dart';

enum ToastType { success, error, warning, info }

class ToastService {
  static void show(
    BuildContext context,
    String message, {
    ToastType type = ToastType.info,
    Duration duration = const Duration(seconds: 3),
  }) {
    final config = _getConfig(type);
    
    ScaffoldMessenger.of(context).hideCurrentSnackBar();
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Row(
          children: [
            Container(
              padding: const EdgeInsets.all(8),
              decoration: BoxDecoration(
                color: config.iconBgColor,
                borderRadius: BorderRadius.circular(8),
              ),
              child: Icon(config.icon, color: Colors.white, size: 20),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Text(
                message,
                style: const TextStyle(
                  color: Colors.white,
                  fontWeight: FontWeight.w500,
                  fontSize: 14,
                ),
              ),
            ),
          ],
        ),
        backgroundColor: config.bgColor,
        behavior: SnackBarBehavior.floating,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
        margin: const EdgeInsets.all(16),
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
        duration: duration,
        dismissDirection: DismissDirection.horizontal,
      ),
    );
  }

  static void success(BuildContext context, String message) {
    show(context, message, type: ToastType.success);
  }

  static void error(BuildContext context, String message) {
    show(context, message, type: ToastType.error, duration: const Duration(seconds: 4));
  }

  static void warning(BuildContext context, String message) {
    show(context, message, type: ToastType.warning);
  }

  static void info(BuildContext context, String message) {
    show(context, message, type: ToastType.info);
  }

  static _ToastConfig _getConfig(ToastType type) {
    switch (type) {
      case ToastType.success:
        return _ToastConfig(
          icon: Icons.check_circle_rounded,
          bgColor: const Color(0xFF1E3A2F),
          iconBgColor: AppTheme.secondaryColor,
        );
      case ToastType.error:
        return _ToastConfig(
          icon: Icons.error_rounded,
          bgColor: const Color(0xFF3A1E1E),
          iconBgColor: AppTheme.errorColor,
        );
      case ToastType.warning:
        return _ToastConfig(
          icon: Icons.warning_rounded,
          bgColor: const Color(0xFF3A351E),
          iconBgColor: AppTheme.warningColor,
        );
      case ToastType.info:
        return _ToastConfig(
          icon: Icons.info_rounded,
          bgColor: const Color(0xFF1E2A3A),
          iconBgColor: AppTheme.primaryColor,
        );
    }
  }
}

class _ToastConfig {
  final IconData icon;
  final Color bgColor;
  final Color iconBgColor;

  _ToastConfig({
    required this.icon,
    required this.bgColor,
    required this.iconBgColor,
  });
}
