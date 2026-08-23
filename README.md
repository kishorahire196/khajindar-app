# खजिनदार वही — स्वतःचं Hosting

हा तुमच्या Claude artifact चाच code आहे, पण आता खऱ्या website म्हणून चालण्यासाठी तयार केलाय.
Data आता **Firebase** मध्ये (मोफत) साठतो — Claude बंद केलं / नवीन version आलं तरी data कायम राहतो.

---

## पायरी 1: Firebase प्रोजेक्ट तयार करा (5 मिनिटं)

1. [console.firebase.google.com](https://console.firebase.google.com) उघडा, Google login करा
2. **"Add project"** दाबा → नाव द्या (उदा. `khajindar-vahi`) → पुढे जा (Analytics नको असेल तर बंद करा)
3. डाव्या मेनूत **Build → Firestore Database** → **"Create database"**
   - Location कुठलंही जवळचं निवडा (उदा. `asia-south1`)
   - **"Start in test mode"** निवडा (नंतर rules बदलू)
4. वरती ⚙️ (Project settings) → खाली **"Your apps"** → **`</>`  (Web)** आयकॉन दाबा
   - App चं नाव द्या → **"Register app"**
   - जो code दिसेल त्यातला `firebaseConfig` object कॉपी करा (यात `apiKey`, `projectId` वगैरे असतं)

5. या प्रोजेक्टमधल्या **`src/firebase.js`** फाईलमध्ये, कॉपी केलेला config paste करून `firebaseConfig` बदला.

6. Firestore च्या **"Rules"** tab मध्ये जाऊन, इथल्या **`firestore.rules`** फाईलमधला मजकूर paste करून **"Publish"** दाबा.

---

## पायरी 2: GitHub वर टाका (मोबाईलवरून सुद्धा)

1. [github.com](https://github.com) वर नवीन repository तयार करा (उदा. `khajindar-app`)
2. GitHub च्या mobile app किंवा website वरून हे सगळे files त्या repo मध्ये upload करा
   (`src/firebase.js` आता तुमच्या स्वतःच्या config सकट असल्याची खात्री करा)

---

## पायरी 3: Netlify ला जोडा (auto-deploy)

1. [netlify.com](https://netlify.com) वर Google/GitHub नी login करा
2. **"Add new site" → "Import an existing project"** → GitHub निवडा → तुमचा `khajindar-app` repo निवडा
3. Build settings आपोआप भरलेली दिसतील (Vite प्रोजेक्ट असल्याने):
   - Build command: `npm run build`
   - Publish directory: `dist`
4. **"Deploy site"** दाबा — 1-2 मिनिटांत खरी website URL मिळेल (उदा. `khajindar-app.netlify.app`)

यानंतर, GitHub repo मध्ये कधीही बदल केलात (मोबाईलवरूनही) की Netlify आपोआप नवीन version deploy करेल.

---

## पर्यायी: स्वतः build करून drag-drop

Computer असेल तर:
```
npm install
npm run build
```
यामुळे तयार झालेला **`dist`** फोल्डर सरळ [netlify.com](https://app.netlify.com/drop) वर drag-drop करा — लगेच live होईल.

---

## लक्षात ठेवा

- `src/firebase.js` मधला config **कधीही "YOUR_API_KEY" असाच ठेवू नका** — तो बदलल्याशिवाय data save होणार नाही
- प्रत्येक अकाउंटचा (अध्यक्ष/खजिनदार/सेक्रेटरी/कार्यकर्ता) पासवर्ड पहिल्यांदा कोणीतरी लॉगिन करताना तयार होतो — आता तो **कायमचा** राहील (Claude artifact सारखा प्रत्येक वेळी रीसेट होणार नाही)
- Firebase चा मोफत plan (Spark) या आकाराच्या वापरासाठी पुरेसा आहे
