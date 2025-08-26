// apiKey: "AIzaSyDSnENpfbQVD5i-o_Dq1bbh04YHCFz7dpw",
//   authDomain: "data-edit-adbcb.firebaseapp.com",
//   projectId: "data-edit-adbcb",
//   storageBucket: "data-edit-adbcb.firebasestorage.app",
//   messagingSenderId: "89538533125",
//   appId: "1:89538533125:web:b0f24d3d44ebd50b16896c",
//   measurementId: "G-ZH1BCCKTGF"


// import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
// import { getFirestore, doc, setDoc, getDoc } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

// 🔹 Your Firebase Config (replace with your own from Firebase Console)// ✅ Firebase config
const firebaseConfig = {
  apiKey: "AIzaSyDSnENpfbQVD5i-o_Dq1bbh04YHCFz7dpw",
  authDomain: "data-edit-adbcb.firebaseapp.com",
  projectId: "data-edit-adbcb",
  storageBucket: "data-edit-adbcb.firebasestorage.app",
  messagingSenderId: "89538533125",
  appId: "1:89538533125:web:b0f24d3d44ebd50b16896c",
  measurementId: "G-ZH1BCCKTGF"
};

// ✅ Initialize Firebase
if (!firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
}
const db = firebase.firestore();

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnAdd").onclick = addDataToFirebase;
  }
});

// 🔹 Function to add structured data
async function addDataToFirebase() {
  // Collect values from the input fields
  const equipment = document.getElementById("equipment").value.trim();
  const sectionNo = document.getElementById("sectionNo").value.trim();
  const sections = document.getElementById("sections").value.trim();
  const serial = document.getElementById("serial").value.trim();
  const keywords = document.getElementById("keywords").value.trim();

  if (!equipment || !sectionNo || !sections || !serial || !keywords) {
    document.getElementById("output").innerText = "⚠️ Please fill out all fields before submitting.";
    return;
  }

  try {
    // Check if serial already exists
    const docRef = db.collection("excel").doc(serial);
    const docSnap = await docRef.get();

    if (docSnap.exists) {
      document.getElementById("output").innerText = "⚠️ This serial already entered.";
      return;
    }

    // If not exists → save new document
    await docRef.set({
      "Equippment / Subsections": equipment,
      "Section No.": sectionNo,
      "Sections": sections,
      "Serial": serial,
      "System Default / Keywords": keywords
    });

    document.getElementById("output").innerText = "✅ Data successfully added to Firebase!";

    // Clear inputs after submission
    document.getElementById("equipment").value = "";
    document.getElementById("sectionNo").value = "";
    document.getElementById("sections").value = "";
    document.getElementById("serial").value = "";
    document.getElementById("keywords").value = "";

  } catch (err) {
    document.getElementById("output").innerText = "❌ Error adding data: " + err.message;
    console.error(err);
  }
}