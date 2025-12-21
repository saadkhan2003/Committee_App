# Committee App

A Flutter application for managing committee payments and tracking member contributions.

## Features

- **Host Dashboard**: Create and manage committees
- **Member Management**: Add, edit, and track members
- **Payment Tracking**: Record and monitor payments
- **Payment Sheet**: Visual payment calendar and status
- **Member Dashboard**: Personal view for committee members
- **Cloud Sync**: Real-time sync with Firebase
- **Auto Updates**: OTA update system with optional force updates

## Getting Started

### Prerequisites

- Flutter SDK (^3.7.2)
- Firebase project with Authentication and Firestore enabled
- Android Studio / VS Code

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/saadkhan2003/Committee_App.git
   cd Committee_App
   ```

2. Install dependencies:
   ```bash
   flutter pub get
   ```

3. Run the app:
   ```bash
   flutter run
   ```

### Build Release APK

```bash
flutter build apk --release
```

The APK will be generated at: `build/app/outputs/flutter-apk/app-release.apk`

## Update System

The app includes an OTA update system controlled via `version.json`:

```json
{
  "version": "1.0.3",
  "apkUrl": "https://github.com/.../app-release.apk",
  "releaseNotes": "• Bug fixes and improvements",
  "forceUpdate": false,
  "minVersion": "1.0.0"
}
```

| Field | Description |
|-------|-------------|
| `version` | Latest version available |
| `apkUrl` | Download URL for the APK |
| `releaseNotes` | What's new in this version |
| `forceUpdate` | If `true`, all users must update |
| `minVersion` | Users below this version must update |

## Project Structure

```
lib/
├── models/          # Data models (Committee, Member, Payment)
├── screens/         # UI screens
│   ├── auth/        # Login/signup screens
│   ├── host/        # Host dashboard and management
│   └── viewer/      # Member view screens
├── services/        # Business logic and API services
└── utils/           # Themes and utilities
```

## Firebase Setup

1. Create a Firebase project
2. Enable Authentication (Email/Password)
3. Enable Cloud Firestore
4. Download `google-services.json` and place in `android/app/`

## License

This project is private.
