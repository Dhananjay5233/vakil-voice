====== VAKIL VOICE ======

सेटअप (Mac):
1. cd ~/Downloads/vakil-voice-clean
2. python3 -m venv venv
3. source venv/bin/activate
4. pip install -r requirements.txt
5. python app.py

फिर Chrome में खोलें: http://localhost:5000

Features:
✓ 7 letter templates (कानूनी नोटिस, वकालतनामा, etc)
✓ Hindi voice typing (सीधे Chrome का)
✓ Save & search letters
✓ Word export (.docx)
✓ Local database (फॉल्डर में)

Templates:
- कानूनी नोटिस
- नोटिस का उत्तर
- वकालतनामा (Power of Attorney)
- लिखित बयान
- जमानत आवेदन
- याचिका
- मांग पत्र

Troubleshooting:
1. Mic नहीं चल रहा? Chrome में खोलें, address bar के lock icon → Microphone → Allow
2. pip error? python3 use करें (python नहीं)
3. Port 5000 busy? app.py के आखिरी line में port=5001 करो

Simple रखा है - सिर्फ काम करने वाली चीज़ें हैं।
