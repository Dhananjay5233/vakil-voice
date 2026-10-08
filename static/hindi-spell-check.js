// Simple Hindi spell checker - common mistakes
const HINDI_CORRECTIONS = {
  // Common typos
  "कीसे": "किसे",
  "जिसे": "जिसे",
  "किधर": "किधर",
  "वहां": "वहाँ",
  "यहां": "यहाँ",
  "जहां": "जहाँ",
  "कभी": "कभी",
  "सदा": "सदा",
  "कुछ": "कुछ",
  "सब": "सब",
  "बहुत": "बहुत",
  "थोडा": "थोड़ा",
  "पडोसी": "पड़ोसी",
  "मालिक": "मालिक",
  "भोजन": "भोजन",
  
  // Double space and punctuation
  "।।": "।",
  "।,": ",",
  "।।।": "।",
  
  // Common phrase mistakes
  "पर न्तु": "परंतु",
  "अत: ": "अतः ",
  "मुझे को": "मुझे",
  "आपको को": "आपको",
  "हमे": "हमें",
  "तुमे": "तुम्हें",
  "जिसके": "जिसके",
  "किसके": "किसके"
};

function spellCheck(text) {
  let corrected = text;
  const mistakes = [];
  
  Object.keys(HINDI_CORRECTIONS).forEach(wrong => {
    const right = HINDI_CORRECTIONS[wrong];
    const regex = new RegExp(wrong, 'g');
    if (regex.test(text)) {
      mistakes.push({original: wrong, suggestion: right});
      corrected = corrected.replace(regex, right);
    }
  });
  
  return {corrected, mistakes};
}
