/**
 * DFO Raid Planner - Firebase Configuration
 *
 * Steps to set up:
 * 1. Go to https://console.firebase.google.com/
 * 2. Create a new project (or use an existing one)
 * 3. Go to Project Settings > General > Your apps > Add app (Web)
 * 4. Copy the firebaseConfig object below and replace the placeholder values
 * 5. In Firebase Console:
 *    - Go to Firestore Database > Create database (start in test mode for now)
 *    - Go to Authentication > Sign-in method > Enable "Anonymous"
 * 6. Deploy firestore.rules from this project folder:
 *    Run: firebase deploy --only firestore:rules
 *    (Requires Firebase CLI: npm install -g firebase-tools)
 */
const firebaseConfig = {
    apiKey: "AIzaSyDWMj6G6-yl57pNaZ3dqYTAadWySk2hbj0",
    authDomain: "dfoplanner.firebaseapp.com",
    projectId: "dfoplanner",
    storageBucket: "dfoplanner.firebasestorage.app",
    messagingSenderId: "673330430304",
    appId: "1:673330430304:web:fcb3c64e8269afb24e1e05",
    measurementId: "G-3RZ7H6P4C2"
};
