from deep_translator import GoogleTranslator

# Translate from Bangla to English
translated_text = GoogleTranslator(source="bn", target="en").translate("সবাত শ্বসনে কত শক্তি খরচ হয়?")
print("Translated:", translated_text)
