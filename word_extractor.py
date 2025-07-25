import torch
import librosa
from transformers import Wav2Vec2Processor, Wav2Vec2ForCTC

# Load processor and model (recommended way)
processor = Wav2Vec2Processor.from_pretrained("facebook/wav2vec2-base-960h")
model = Wav2Vec2ForCTC.from_pretrained("facebook/wav2vec2-base-960h")

# Optional: handle words like "seven" → "7"
word_to_digit = {
    "zero": "0", "oh": "0", "o": "0","zeril":'0',"zerel":'0','zereol':'0','zelo':'0','zemo':'0','zere':'0','zarol':'0','zyril':'0','zyral':'0','zeral':'0',
    "one": "1", "won": "1", "juan": "1","when":'1','wen':'1','ene':'1','wone':"1",'on':'1',
    "two": "2", "to": "2", "too": "2","thwo":'2','tro':'2',
    "three": "3", "tree": "3", "fhree": "3","hree":"3","free":"3",'thrty':'3','thre':'3','threw':'3','thirty':'3',
    "four": "4", "for": "4", "fore": "4",
    "five": "5", "fine": "5", "far": "5", "fire": "5",'fie':'5','fine':'5',
    "six": "6", "sics": "6", "sex": "6", "siks": "6", "siic": "6","suics":"6",'seics':'6','sixs':'6','sixh':'6','sucs':'6','sircs':'6',
    "seven": "7", "sevin": "7", "savin": "7","semen":"7","sevn":"7","sem":"7","servan":"7","sam":"7","seve":'7','servant':'7','seveny':'7','several':'7','saysin':'7','siven':'7',
    "eight": "8", "ate": "8", "ait": "8",'eit':'8','ay':'8',
    "nine": "9", "nien": "9", "none": "9", "mine": "9","narn":"9"
}

def transcribe_wav_to_digits(wav_path: str) -> str:
    print(f"[📂] Loading audio from: {wav_path}")
    
    # Load audio (Wav2Vec2 expects 16kHz mono)
    audio, sr = librosa.load(wav_path, sr=16000)

    # Preprocess with processor
    inputs = processor(audio, sampling_rate=16000, return_tensors="pt")

    # Forward pass
    with torch.no_grad():
        logits = model(**inputs).logits

    # Get predicted token IDs
    predicted_ids = torch.argmax(logits, dim=-1)

    # Decode predicted text
    transcription = processor.batch_decode(predicted_ids)[0]
    print(f"[🧠 Transcription]: {transcription}")

    # Try to extract digits directly
    digits = ''.join([c for c in transcription if c.isdigit()])

    # Fallback: convert word digits to numbers
    if len(digits) < 6:
        words = transcription.lower().split()
        for word in words:
            if word in word_to_digit:
                digits += word_to_digit[word]
            if len(digits) == 6:
                break

    print(f"[✅ Extracted Digits]: {digits}")
    return digits
